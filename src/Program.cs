using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using ConciliarApp.Services; // Certifique-se de que esta linha está presente

namespace ConciliarApp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Por favor, informe o nome do cartão (ex: MASTER, VISA) e o nome do arquivo TXT.");
                return;
            }

            string cartao = args[0].ToUpper();
            string nomeArquivoTxt = args[1];
            string caminhoArquivoTxt = Path.Combine(@"C:\Users\jeova\OneDrive\FileSync\Extratos", nomeArquivoTxt);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Definindo o contexto da licença

            string caminhoArquivoExcel = @"C:\Users\jeova\OneDrive\FileSync\_ControleFin\RecDesp-2025.xlsx";

            var conciliacaoService = new ConciliacaoService();

            // Processar o arquivo Excel
            (int qtdLancamentosExcel, decimal totalExcel, HashSet<(DateTime, decimal)> lancamentosExcel) = conciliacaoService.ProcessarArquivoExcel(caminhoArquivoExcel, cartao);

            // Processar o arquivo TXT
            (int qtdLancamentosTxt, decimal totalTxt, List<(DateTime, decimal, string)> lancamentosTxt) = conciliacaoService.ProcessarArquivoTxt(caminhoArquivoTxt);

            // Exibir a diferença entre o Excel e o TXT
            Console.WriteLine();
            Console.WriteLine($"Diferença entre Extrato x Excel");
            var diferenca = qtdLancamentosTxt - qtdLancamentosExcel;
            var sinal = diferenca < 0 ? "-" : diferenca > 0 ? "+" : "";
            Console.WriteLine($"  Qtde de lançamentos: {sinal}{diferenca}");
            Console.WriteLine($"  Valor: {(totalTxt - totalExcel).ToString("C", CultureInfo.GetCultureInfo("pt-BR"))}");

            // Exibir lançamentos que estão no extrato e não estão no Excel
            conciliacaoService.ExibirLancamentosNoExtratoENaoNoExcel(lancamentosTxt, lancamentosExcel);

            // Exibir lançamentos que estão em ambos, mas têm diferença de valor de até 15 centavos
            conciliacaoService.ExibirLancamentosComPequenaDiferenca(lancamentosTxt, lancamentosExcel);
        }
    }
}