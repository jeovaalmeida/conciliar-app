using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using ConciliarApp.Extensions;
using ConciliarApp.Models;

namespace ConciliarApp.Services
{
    public class ConciliacaoService
    {
        // Métodos de extração
        
        public List<LancamentoExtrato> ExtrairLancamentosDoExtrato(string caminhoArquivo)
        {
            try
            {
                var linhas = File.ReadAllLines(caminhoArquivo);
                List<LancamentoExtrato> lancamentosTxt = new List<LancamentoExtrato>();

                foreach (var linha in linhas)
                {
                    if (LinhaEhValida(linha))
                    {
                        string parteData = linha.Substring(0, 10).Trim();
                        string descricao = linha.Substring(10, linha.Length - 30).Trim();
                        string parteValor = linha.Substring(linha.Length - 20, 10).Trim();

                        if (LancamentoEhValido(parteData, parteValor, out DateTime dataConvertida, out decimal valorConvertido))
                        {
                            lancamentosTxt.Add(new LancamentoExtrato
                            {
                                Data = dataConvertida,
                                Valor = valorConvertido,
                                Descricao = descricao.Truncate(47).PadRight(50, ' '),
                                ExisteNoExcel = false
                            });
                        }
                    }
                }

                return lancamentosTxt;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar o arquivo TXT: {ex.Message}");
                return new List<LancamentoExtrato>();
            }
        }

        public HashSet<LancamentoExcel> ExtrairLancamentosDoExcel(string caminhoArquivo, string cartao, string nomePlanilha, out int linhaInicial, out int linhaInsercao)
        {
            FileInfo arquivoInfo = new FileInfo(caminhoArquivo);
            linhaInicial = 0;
            linhaInsercao = 0;

            using (ExcelPackage pacote = new ExcelPackage(arquivoInfo))
            {
                ExcelWorksheet planilha = pacote.Workbook.Worksheets[nomePlanilha];
                int qtdLinhas = planilha.Dimension.Rows;
                bool encontrouCartaoDeCredito = false;
                HashSet<LancamentoExcel> lancamentosExcel = new();

                for (int linha = 1; linha <= qtdLinhas; linha++)
                {
                    string valor1aCelula = planilha.Cells[linha, 1].Text;
                    string valor2aCelula = planilha.Cells[linha, 2].Text;
                    string valor3aCelula = planilha.Cells[linha, 3].Text;

                    if (valor1aCelula.Equals($"CARTÃO DE CRÉDITO: {cartao}", StringComparison.OrdinalIgnoreCase))
                    {
                        encontrouCartaoDeCredito = true;
                        linhaInicial = linha + 1;
                    }
                    else if (encontrouCartaoDeCredito && valor2aCelula.Equals($"TOTAL ({cartao}):", StringComparison.OrdinalIgnoreCase))
                    {
                        encontrouCartaoDeCredito = false;
                    }
                    else if (encontrouCartaoDeCredito)
                    {
                        string valor = planilha.Cells[linha, 6].Text;
                        string data = planilha.Cells[linha, 7].Text;
                        string valor4aCelula = planilha.Cells[linha, 4].Text;
                        var descricao = string.IsNullOrEmpty(valor4aCelula) ? $"{valor2aCelula} - {valor3aCelula}" : $"{valor3aCelula} - {valor4aCelula}";

                        if (LancamentoEhValido(data, valor, out DateTime dataConvertida, out decimal valorConvertido))
                        {
                            lancamentosExcel.Add(new LancamentoExcel
                            {
                                Data = dataConvertida,
                                Valor = valorConvertido,
                                Descricao = descricao.Truncate(47).PadRight(50, ' '),
                                DiferencaDePequenoValor = false,
                                NaoExisteNoExtrato = false
                            });
                        }

                        // Identificar a linha de inserção
                        if (valor2aCelula == "(novo - copiar/colar)")
                        {
                            linhaInsercao = linha;
                        }
                    }
                }

                return lancamentosExcel;
            }
        }

        public void ExibirLancamentosDoExtrato(List<LancamentoExtrato> lancamentosTxt, string nomeArquivoExtrato)
        {
            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS DO EXTRATO: {lancamentosTxt.Count} | Arquivo: {nomeArquivoExtrato}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosTxt)
            {
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {lancamento.Descricao}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        // Métodos de exibição
        public void ExibirLancamentosDoExcel(HashSet<LancamentoExcel> lancamentosExcel, string cartao, int linhaInicial)
        {
            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS DO EXCEL: {lancamentosExcel.Count} | CARTÃO {cartao}");
            Console.WriteLine($"Iniciando a leitura dos lançamentos na linha {linhaInicial}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosExcel)
            {
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {lancamento.Descricao}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public void ExibirLancamentosNoExtratoENaoNoExcel(List<LancamentoExtrato> lancamentosExtrato)
        {
            var lancamentosNaoNoExcel = lancamentosExtrato.Where(l => !l.ExisteNoExcel).ToList();

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS NO EXTRATO E NÃO NO EXCEL: {lancamentosNaoNoExcel.Count}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosNaoNoExcel)
            {
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {lancamento.Descricao}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public void ExibirLancamentosNoExcelENaoNoExtrato(HashSet<LancamentoExcel> lancamentosExcel)
        {
            var lancamentosNaoNoExtrato = lancamentosExcel.Where(l => l.NaoExisteNoExtrato).ToList();

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS NO EXCEL E NÃO NO EXTRATO: {lancamentosNaoNoExtrato.Count}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosNaoNoExtrato)
            {
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {lancamento.Descricao}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public void ExibirDiferencaEntreExtratoEExcel(int qtdLancamentosTxt, int qtdLancamentosExcel, decimal totalTxt, decimal totalExcel)
        {
            Console.WriteLine();
            Console.WriteLine($"DIFERENÇA ENTRE EXTRATO x EXCEL (Extrato: {qtdLancamentosTxt}, Excel: {qtdLancamentosExcel})");
            var diferenca = qtdLancamentosTxt - qtdLancamentosExcel;
            var sinal = diferenca < 0 ? "-" : diferenca > 0 ? "+" : "";
            Console.WriteLine($"  Qtde de lançamentos: {sinal}{diferenca}");
            Console.WriteLine($"  Valor: {(totalTxt - totalExcel).ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public void ExibirLancamentosComPequenaDiferenca(List<(DateTime Data, string Descricao, decimal ValorExcel, decimal ValorExtrato)> lancamentosComPequenaDiferenca)
        {
            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS COM PEQUENA DIFERENÇA: {lancamentosComPequenaDiferenca.Count}");
            decimal valorTotalExcel = 0;
            decimal valorTotalExtrato = 0;
            foreach (var lancamento in lancamentosComPequenaDiferenca)
            {
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {lancamento.Descricao}, Valor Excel: {lancamento.ValorExcel.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}, Valor Extrato: {lancamento.ValorExtrato.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotalExcel += lancamento.ValorExcel;
                valorTotalExtrato += lancamento.ValorExtrato;
            }
            Console.WriteLine($"Total Geral Excel: {valorTotalExcel.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
            Console.WriteLine($"Total Geral Extrato: {valorTotalExtrato.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        private bool LancamentoEhValido(string data, string valor, out DateTime dataConvertida, out decimal valorConvertido)
        {
            dataConvertida = DateTime.MinValue;
            valorConvertido = 0;

            if (DateTime.TryParse(data, out dataConvertida) && decimal.TryParse(valor, out valorConvertido))
            {
                return true;
            }

            return false;
        }

        private bool LinhaEhValida(string linha)
        {
            // Ignorar linhas em branco, cabeçalhos e linhas que contêm "PGTO DEBITO CONTA"
            if (string.IsNullOrWhiteSpace(linha) || !char.IsDigit(linha[0]) || linha.Contains("PGTO DEBITO CONTA"))
            {
                return false;
            }

            return true;
        }

        public LancamentosProcessados ExtrairEMarcarLancamentos(string caminhoArquivoExcel, string caminhoArquivoExtrato, string cartao, string nomePlanilha)
        {
            var lancamentosProcessados = new LancamentosProcessados
            {
                LancamentosExcel = ExtrairLancamentosDoExcel(caminhoArquivoExcel, cartao, nomePlanilha, out int linhaInicial, out int linhaInsercao),
                LancamentosExtrato = ExtrairLancamentosDoExtrato(caminhoArquivoExtrato),
                LancamentosComPequenaDiferenca = new List<(DateTime Data, string Descricao, decimal ValorExcel, decimal ValorExtrato)>(),
                LinhaInsercao = linhaInsercao
            };

            foreach (var lancamentoTxt in lancamentosProcessados.LancamentosExtrato)
            {
                var lancamentoExcel = lancamentosProcessados.LancamentosExcel.FirstOrDefault(e => e.Data == lancamentoTxt.Data && Math.Abs(e.Valor - lancamentoTxt.Valor) <= 0.15m);
                if (lancamentoExcel != null)
                {
                    lancamentoTxt.ExisteNoExcel = (lancamentoExcel.Valor == lancamentoTxt.Valor);
                    lancamentoExcel.DiferencaDePequenoValor = (lancamentoExcel.Valor != lancamentoTxt.Valor);
                    if (lancamentoExcel.DiferencaDePequenoValor )
                    {
                        lancamentosProcessados.LancamentosComPequenaDiferenca.Add((lancamentoExcel.Data, lancamentoExcel.Descricao, lancamentoExcel.Valor, lancamentoTxt.Valor));
                    }
                }
            }

            foreach (var lancamentoExcel in lancamentosProcessados.LancamentosExcel)
            {
                if (!lancamentosProcessados.LancamentosExtrato.Any(t => t.Data == lancamentoExcel.Data && t.Valor == lancamentoExcel.Valor) && !lancamentoExcel.DiferencaDePequenoValor)
                {
                    lancamentoExcel.NaoExisteNoExtrato = true;
                }
            }

            return lancamentosProcessados;
        }

        public void InserirLancamentosNoExcel(string caminhoArquivoExcel, List<LancamentoExtrato> lancamentosNaoNoExcel, int linhaInsercao, string nomePlanilha)
        {
            FileInfo arquivoInfo = new FileInfo(caminhoArquivoExcel);

            using (ExcelPackage pacote = new ExcelPackage(arquivoInfo))
            {
                ExcelWorksheet planilha = pacote.Workbook.Worksheets[nomePlanilha];

                if (linhaInsercao == 0)
                {
                    Console.WriteLine("Linha de inserção não encontrada.");
                    return;
                }

                CriarLinhasDuplicandoADeCopiarColar(lancamentosNaoNoExcel, linhaInsercao, planilha);

                // Inserir os lançamentos do extrato
                for (int i = 0; i < lancamentosNaoNoExcel.Count; i++)
                {
                    var lancamento = lancamentosNaoNoExcel[i];
                    int linhaAtual = linhaInsercao + i;

                    var (categoria, fornecedor) = ObterCategoriaEFornecedor(lancamento);

                    planilha.Cells[linhaAtual, 2].Value = categoria ?? string.Empty;
                    planilha.Cells[linhaAtual, 3].Value = fornecedor ?? lancamento.Descricao;
                    planilha.Cells[linhaAtual, 6].Value = lancamento.Valor;
                    planilha.Cells[linhaAtual, 7].Value = lancamento.Data.ToString("dd/MM/yyyy");

                    // Pintar a célula da linha inserida com fundo amarelo
                    planilha.Cells[linhaAtual, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    planilha.Cells[linhaAtual, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                }

                // Salvar as alterações no arquivo Excel
                pacote.Save();

                Console.WriteLine($"\r\n{lancamentosNaoNoExcel.Count} lançamento(s) do extrato nao existente(s) no excel inserido(s) na planilha na linha {linhaInsercao}.");
            }
        }

        private static void CriarLinhasDuplicandoADeCopiarColar(List<LancamentoExtrato> lancamentosNaoNoExcel, int linhaInsercao, ExcelWorksheet planilha)
        {
            // Inserir novas linhas
            planilha.InsertRow(linhaInsercao, lancamentosNaoNoExcel.Count, linhaInsercao);

            // Copiar o conteúdo da linha original para as novas linhas
            for (int i = 0; i < lancamentosNaoNoExcel.Count; i++)
            {
                int linhaAtual = linhaInsercao + i;
                for (int col = 1; col <= planilha.Dimension.Columns; col++)
                {
                    planilha.Cells[linhaAtual, col].Value = planilha.Cells[linhaInsercao - 1, col].Value;
                    planilha.Cells[linhaAtual, col].StyleID = planilha.Cells[linhaInsercao - 1, col].StyleID;

                    // Ajustar as fórmulas para referenciar a linha correta
                    if (!string.IsNullOrEmpty(planilha.Cells[linhaInsercao - 1, col].Formula))
                    {
                        planilha.Cells[linhaAtual, col].FormulaR1C1 = planilha.Cells[linhaInsercao - 1, col].FormulaR1C1;
                    }
                }
            }
        }

        private (string, string) ObterCategoriaEFornecedor(LancamentoExtrato lancamento)
        {
            if (lancamento.Descricao.Contains("RDSAUDE"))
                return ("Farmácia - Remédios - ", "Drogasil");
            else if (lancamento.Descricao.Contains("MSCAP"))
                return ("Loteria", "Ms Cap");    
            else if (lancamento.Descricao.Contains("PAG POKO"))
                return ("Mercado", "Pag Poko");
            else if (lancamento.Descricao.Contains("ASSAI"))
                return ("Mercado", "Assaí");

            return (null, null);
        }
    }

    public class LancamentosProcessados
    {
        public HashSet<LancamentoExcel> LancamentosExcel { get; set; }
        public List<LancamentoExtrato> LancamentosExtrato { get; set; }
        public List<(DateTime Data, string Descricao, decimal ValorExcel, decimal ValorExtrato)> LancamentosComPequenaDiferenca { get; set; }
        public int LinhaInsercao { get; set; }
    }
}