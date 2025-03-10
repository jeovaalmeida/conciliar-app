using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using ConciliarApp.Services;
using ConciliarApp.Models;

namespace ConciliarApp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Por favor, informe o nome do cartão (ex: MASTER, VISA) e o nome do arquivo de extrato (TXT).");
                return;
            }

            string cartao = args[0].ToUpper();
            string nomeArquivoExtrato = args[1];
            string caminhoArquivoExtrato = Path.Combine(@"C:\Users\jeova\OneDrive\FileSync\Extratos", nomeArquivoExtrato);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string caminhoArquivoExcel = @"C:\Users\jeova\OneDrive\FileSync\_ControleFin\RecDesp-2025.xlsx";

            var conciliacaoService = new ConciliacaoService();

            Console.WriteLine($"\r\n{DateTime.Now} | Iniciando processamento");

            var lancamentosProcessados = conciliacaoService.ExtrairEMarcarLancamentos(caminhoArquivoExcel, caminhoArquivoExtrato, cartao);

            // Exibir todos os lançamentos do extrato
            conciliacaoService.ExibirLancamentosDoExtrato(lancamentosProcessados.LancamentosExtrato, nomeArquivoExtrato);

            // Exibir todos os lançamentos do Excel
            conciliacaoService.ExibirLancamentosDoExcel(lancamentosProcessados.LancamentosExcel, cartao, 0);

            // Exibir a diferença entre o Excel e o TXT
            conciliacaoService.ExibirDiferencaEntreExtratoEExcel(lancamentosProcessados.LancamentosExtrato.Count, lancamentosProcessados.LancamentosExcel.Count, lancamentosProcessados.LancamentosExtrato.Sum(l => l.Valor), lancamentosProcessados.LancamentosExcel.Sum(l => l.Valor));

            // Exibir lançamentos que estão no extrato e não estão no Excel
            conciliacaoService.ExibirLancamentosNoExtratoENaoNoExcel(lancamentosProcessados.LancamentosExtrato);

            // Exibir lançamentos que estão no Excel e não estão no extrato
            conciliacaoService.ExibirLancamentosNoExcelENaoNoExtrato(lancamentosProcessados.LancamentosExcel);

            // Exibir lançamentos que estão em ambos, mas têm diferença de valor de até 15 centavos
            conciliacaoService.ExibirLancamentosComPequenaDiferenca(lancamentosProcessados.LancamentosComPequenaDiferenca);
        }
    }
}