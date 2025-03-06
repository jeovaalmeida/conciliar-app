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
        public HashSet<LancamentoExcel> ExtrairLancamentosDoExcel(string caminhoArquivo, string cartao, out int linhaInicial)
        {
            FileInfo arquivoInfo = new FileInfo(caminhoArquivo);
            linhaInicial = 0;

            using (ExcelPackage pacote = new ExcelPackage(arquivoInfo))
            {
                ExcelWorksheet planilha = pacote.Workbook.Worksheets["2025-03"];
                int qtdLinhas = planilha.Dimension.Rows;
                bool encontrouCartaoDeCredito = false;
                HashSet<LancamentoExcel> lancamentosExcel = new HashSet<LancamentoExcel>();

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
                        string descricao = planilha.Cells[linha, 4].Text;
                        descricao = string.IsNullOrEmpty(descricao) ? $"{valor2aCelula} - {valor3aCelula}" : descricao;

                        if (LancamentoEhValido(data, valor, out DateTime dataConvertida, out decimal valorConvertido))
                        {
                            lancamentosExcel.Add(new LancamentoExcel
                            {
                                Data = dataConvertida,
                                Valor = valorConvertido,
                                Descricao = descricao,
                                DiferencaDePequenoValor = false,
                                NaoExisteNoExtrato = true
                            });
                        }
                    }
                }

                return lancamentosExcel;
            }
        }

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
                                Descricao = descricao,
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

        // Métodos de exibição
        public void ExibirLancamentosDoExcel(HashSet<LancamentoExcel> lancamentosExcel, string cartao, int linhaInicial)
        {
            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS DO EXCEL - CARTÃO {cartao}: {lancamentosExcel.Count}");
            Console.WriteLine($"Iniciando a leitura dos lançamentos na linha {linhaInicial}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosExcel)
            {
                string descricaoTruncada = lancamento.Descricao.Truncate(50);
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {descricaoTruncada}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public void ExibirLancamentosDoExtrato(List<LancamentoExtrato> lancamentosTxt)
        {
            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS DO EXTRATO: {lancamentosTxt.Count}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosTxt)
            {
                string descricaoTruncada = lancamento.Descricao.Truncate(50);
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {descricaoTruncada}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        // Métodos existentes
        public (int, decimal, HashSet<LancamentoExcel>) ProcessarArquivoExcel(string caminhoArquivo, string cartao)
        {
            var lancamentosExcel = ExtrairLancamentosDoExcel(caminhoArquivo, cartao, out int linhaInicial);
            int qtdLancamentosValidos = lancamentosExcel.Count;
            decimal valorTotal = lancamentosExcel.Sum(l => l.Valor);

            Console.WriteLine($"Total de lançamentos válidos lidos: {qtdLancamentosValidos}");
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");

            return (qtdLancamentosValidos, valorTotal, lancamentosExcel);
        }

        public (int, decimal, List<LancamentoExtrato>) ProcessarArquivoTxt(string caminhoArquivo)
        {
            var lancamentosTxt = ExtrairLancamentosDoExtrato(caminhoArquivo);
            int qtdLancamentosValidos = lancamentosTxt.Count;
            decimal valorTotal = lancamentosTxt.Sum(l => l.Valor);

            Console.WriteLine($"Total de lançamentos válidos lidos: {qtdLancamentosValidos}");
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");

            return (qtdLancamentosValidos, valorTotal, lancamentosTxt);
        }

        public void ExibirLancamentosNoExtratoENaoNoExcel(List<LancamentoExtrato> lancamentosTxt, HashSet<LancamentoExcel> lancamentosExcel)
        {
            var lancamentosNaoNoExcel = new List<LancamentoExtrato>();

            foreach (var lancamento in lancamentosTxt)
            {
                if (lancamento.Descricao.Contains("ANUIDADE DIFERENCIADA") || lancamento.Descricao.Contains("DESC AUTOMATICO ANUD") || EhStreaming(lancamento.Descricao))
                {
                    if (!lancamentosExcel.Any(e => e.Valor == lancamento.Valor))
                    {
                        lancamentosNaoNoExcel.Add(lancamento);
                    }
                }
                else
                {
                    if (!lancamentosExcel.Any(e => e.Data == lancamento.Data && e.Valor == lancamento.Valor))
                    {
                        lancamentosNaoNoExcel.Add(lancamento);
                    }
                }
            }

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS NO EXTRATO E NÃO NO EXCEL: {lancamentosNaoNoExcel.Count}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosNaoNoExcel)
            {
                string descricaoTruncada = lancamento.Descricao.Truncate(50); // Truncar a descrição para 50 caracteres
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {descricaoTruncada}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public void ExibirLancamentosNoExcelENaoNoExtrato(HashSet<LancamentoExcel> lancamentosExcel, List<LancamentoExtrato> lancamentosTxt, List<LancamentoExcel> lancamentosComPequenaDiferenca)
        {
            var lancamentosNaoNoExtrato = new List<LancamentoExcel>();

            foreach (var lancamento in lancamentosExcel)
            {
                if (!lancamentosTxt.Any(t => t.Data == lancamento.Data && t.Valor == lancamento.Valor) && !lancamento.DiferencaDePequenoValor)
                {
                    lancamento.NaoExisteNoExtrato = true;
                    lancamentosNaoNoExtrato.Add(lancamento);
                }
            }

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS NO EXCEL E NÃO NO EXTRATO: {lancamentosNaoNoExtrato.Count}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosNaoNoExtrato)
            {
                string descricaoTruncada = lancamento.Descricao.Truncate(50); // Truncar a descrição para 50 caracteres
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {descricaoTruncada}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
        }

        public List<LancamentoExcel> ExibirLancamentosComPequenaDiferenca(List<LancamentoExtrato> lancamentosTxt, HashSet<LancamentoExcel> lancamentosExcel)
        {
            var lancamentosComPequenaDiferenca = new List<LancamentoExcel>();

            foreach (var lancamentoTxt in lancamentosTxt)
            {
                var lancamentoExcel = lancamentosExcel.FirstOrDefault(e => e.Data == lancamentoTxt.Data && Math.Abs(e.Valor - lancamentoTxt.Valor) <= 0.15m);
                if (lancamentoExcel != null)
                {
                    lancamentoExcel.DiferencaDePequenoValor = true;
                    lancamentoTxt.ExisteNoExcel = true;
                    lancamentosComPequenaDiferenca.Add(lancamentoExcel);
                }
            }

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS COM PEQUENA DIFERENÇA: {lancamentosComPequenaDiferenca.Count}");
            decimal valorTotal = 0;
            foreach (var lancamento in lancamentosComPequenaDiferenca)
            {
                string descricaoTruncada = lancamento.Descricao.Truncate(50); // Truncar a descrição para 50 caracteres
                Console.WriteLine($"Data: {lancamento.Data.ToString("dd/MM/yyyy")}, Descrição: {descricaoTruncada}, Valor: {lancamento.Valor.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                valorTotal += lancamento.Valor;
            }
            Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");

            return lancamentosComPequenaDiferenca;
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

        private bool EhStreaming(string descricao)
        {
            var servicosStreaming = new List<string> { "NETFLIX", "YOUTUBE", "AMAZONPRIME" };
            return servicosStreaming.Any(servico => descricao.Contains(servico, StringComparison.OrdinalIgnoreCase));
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

        private string TruncarDescricao(string descricao, int maxLength = 80)
        {
            if (descricao.Length <= maxLength)
            {
                return descricao;
            }
            return descricao.Substring(0, maxLength) + "...";
        }

        private bool EstaNaListaDePequenaDiferenca((DateTime, decimal, string) lancamento, List<(DateTime, decimal, string)> lancamentosComPequenaDiferenca)
        {
            return lancamentosComPequenaDiferenca.Any(l => l.Item1 == lancamento.Item1 && l.Item2 == lancamento.Item2 && l.Item3 == lancamento.Item3);
        }
    }
}