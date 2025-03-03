using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace ConciliarApp.Services
{
    public class ConciliacaoService
    {
        public (int, decimal, HashSet<(DateTime, decimal)>) ProcessarArquivoExcel(string caminhoArquivo, string cartao)
        {
            FileInfo arquivoInfo = new FileInfo(caminhoArquivo);

            using (ExcelPackage pacote = new ExcelPackage(arquivoInfo))
            {
                ExcelWorksheet planilha = pacote.Workbook.Worksheets["2025-03"];
                int qtdLinhas = planilha.Dimension.Rows;
                bool encontrouCartaoDeCredito = false;
                int qtdLancamentosValidos = 0;
                decimal valorTotal = 0;
                int linhaInicial = 0;
                HashSet<(DateTime, decimal)> lancamentosExcel = new HashSet<(DateTime, decimal)>();

                Console.WriteLine($"LANÇAMENTOS DO EXCEL - CARTÃO: {cartao}");

                for (int linha = 1; linha <= qtdLinhas; linha++)
                {
                    string valor1aCelula = planilha.Cells[linha, 1].Text;
                    string valor2aCelula = planilha.Cells[linha, 2].Text;

                    if (valor1aCelula.Equals($"CARTÃO DE CRÉDITO: {cartao}", StringComparison.OrdinalIgnoreCase))
                    {
                        encontrouCartaoDeCredito = true;
                        linhaInicial = linha + 1; // A leitura começa na próxima linha
                        Console.WriteLine($"Iniciando a leitura dos lançamentos na linha {linhaInicial}");
                    }
                    else if (encontrouCartaoDeCredito && valor2aCelula.Equals($"TOTAL ({cartao}):", StringComparison.OrdinalIgnoreCase))
                    {
                        // Se encontrar a linha de total do cartão, parar a leitura
                        encontrouCartaoDeCredito = false;
                    }
                    else if (encontrouCartaoDeCredito)
                    {
                        string valor = planilha.Cells[linha, 6].Text;
                        string data = planilha.Cells[linha, 7].Text;

                        if (LancamentoEhValido(data, valor, out DateTime dataConvertida, out decimal valorConvertido))
                        {
                            lancamentosExcel.Add((dataConvertida, valorConvertido));
                            Console.WriteLine($"Data: {dataConvertida.ToShortDateString()}, Valor: {valorConvertido.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                            valorTotal += valorConvertido;
                            qtdLancamentosValidos++;
                        }
                    }
                }

                Console.WriteLine($"Total de lançamentos válidos lidos: {qtdLancamentosValidos}");
                Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");

                return (qtdLancamentosValidos, valorTotal, lancamentosExcel);
            }
        }

        public (int, decimal, List<(DateTime, decimal, string)>) ProcessarArquivoTxt(string caminhoArquivo)
        {
            try
            {
                var linhas = File.ReadAllLines(caminhoArquivo);
                int qtdLancamentosValidos = 0;
                decimal valorTotal = 0;
                List<(DateTime, decimal, string)> lancamentosTxt = new List<(DateTime, decimal, string)>();
                Console.WriteLine();
                Console.WriteLine("LANÇAMENTOS DO EXTRATO");

                foreach (var linha in linhas)
                {
                    if (LinhaEhValida(linha))
                    {
                        // Extrair data, descrição e valor
                        string parteData = linha.Substring(0, 10).Trim();
                        string descricao = linha.Substring(10, linha.Length - 30).Trim(); // Ajuste para capturar a descrição
                        string parteValor = linha.Substring(linha.Length - 20, 10).Trim(); // Ajuste para capturar o valor em reais

                        if (LancamentoEhValido(parteData, parteValor, out DateTime dataConvertida, out decimal valorConvertido))
                        {
                            lancamentosTxt.Add((dataConvertida, valorConvertido, descricao));
                            Console.WriteLine($"Data: {dataConvertida.ToString("dd/MM/yyyy")}, Descrição: {descricao}, Valor: {valorConvertido.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
                            valorTotal += valorConvertido;
                            qtdLancamentosValidos++;
                        }
                    }
                }

                Console.WriteLine($"Total de lançamentos válidos lidos: {qtdLancamentosValidos}");
                Console.WriteLine($"Total Geral: {valorTotal.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");

                return (qtdLancamentosValidos, valorTotal, lancamentosTxt);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar o arquivo TXT: {ex.Message}");
                return (0, 0, new List<(DateTime, decimal, string)>());
            }
        }

        public void ExibirLancamentosNoExtratoENaoNoExcel(List<(DateTime, decimal, string)> lancamentosTxt, HashSet<(DateTime, decimal)> lancamentosExcel)
        {
            var lancamentosNaoNoExcel = new List<(DateTime, decimal, string)>();

            foreach (var lancamento in lancamentosTxt)
            {
                if (lancamento.Item3.Contains("ANUIDADE DIFERENCIADA") || lancamento.Item3.Contains("DESC AUTOMATICO ANUD") || EhStreaming(lancamento.Item3))
                {
                    // Comparar apenas pelo valor
                    if (!lancamentosExcel.Any(e => e.Item2 == lancamento.Item2))
                    {
                        lancamentosNaoNoExcel.Add(lancamento);
                    }
                }
                else
                {
                    // Comparar data e valor
                    if (!lancamentosExcel.Contains((lancamento.Item1, lancamento.Item2)))
                    {
                        lancamentosNaoNoExcel.Add(lancamento);
                    }
                }
            }

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS NO EXTRATO E NÃO NO EXCEL: {lancamentosNaoNoExcel.Count}");
            foreach (var lancamento in lancamentosNaoNoExcel)
            {
                Console.WriteLine($"Data: {lancamento.Item1.ToString("dd/MM/yyyy")}, Descrição: {TruncarDescricao(lancamento.Item3)}, Valor: {lancamento.Item2.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
            }
        }

        public void ExibirLancamentosComPequenaDiferenca(List<(DateTime, decimal, string)> lancamentosTxt, HashSet<(DateTime, decimal)> lancamentosExcel, decimal diferencaMaxima = 0.15m)
        {
            var lancamentosComPequenaDiferenca = new List<(DateTime, decimal, string, decimal)>();

            foreach (var lancamento in lancamentosTxt)
            {
                var lancamentoCorrespondente = lancamentosExcel.FirstOrDefault(e => e.Item1 == lancamento.Item1 && Math.Abs(e.Item2 - lancamento.Item2) <= diferencaMaxima && Math.Abs(e.Item2 - lancamento.Item2) > 0);
                if (lancamentoCorrespondente != default)
                {
                    var diferenca = lancamento.Item2 - lancamentoCorrespondente.Item2;
                    lancamentosComPequenaDiferenca.Add((lancamento.Item1, lancamento.Item2, lancamento.Item3, diferenca));
                }
            }

            Console.WriteLine();
            Console.WriteLine($"LANÇAMENTOS COM PEQUENA DIFERENÇA DE VALOR (ATÉ {diferencaMaxima.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}): {lancamentosComPequenaDiferenca.Count}");
            foreach (var lancamento in lancamentosComPequenaDiferenca)
            {
                Console.WriteLine($"{lancamento.Item1.ToString("dd/MM/yyyy")}, {TruncarDescricao(lancamento.Item3)}, Extrato: {lancamento.Item2.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}, Excel: {(lancamento.Item2 - lancamento.Item4).ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}, Diferença: {lancamento.Item4.ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");
            }
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
    }
}