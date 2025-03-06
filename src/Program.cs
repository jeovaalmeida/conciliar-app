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
                Console.WriteLine("Por favor, informe o nome do cartão (ex: MASTER, VISA) e o nome do arquivo TXT.");
                return;
            }

            string cartao = args[0].ToUpper();
            string nomeArquivoTxt = args[1];
            string caminhoArquivoTxt = Path.Combine(@"C:\Users\jeova\OneDrive\FileSync\Extratos", nomeArquivoTxt);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string caminhoArquivoExcel = @"C:\Users\jeova\OneDrive\FileSync\_ControleFin\RecDesp-2025.xlsx";

            var conciliacaoService = new ConciliacaoService();

            // Extrair lançamentos do extrato
            List<LancamentoExtrato> lancamentosTxt = conciliacaoService.ExtrairLancamentosDoExtrato(caminhoArquivoTxt);

            // Extrair lançamentos do Excel
            int linhaInicial;
            HashSet<LancamentoExcel> lancamentosExcel = conciliacaoService.ExtrairLancamentosDoExcel(caminhoArquivoExcel, cartao, out linhaInicial);

            // Exibir todos os lançamentos do extrato
            conciliacaoService.ExibirLancamentosDoExtrato(lancamentosTxt);

            // Exibir todos os lançamentos do Excel
            conciliacaoService.ExibirLancamentosDoExcel(lancamentosExcel, cartao, linhaInicial);

            // Exibir a diferença entre o Excel e o TXT
            conciliacaoService.ExibirDiferencaEntreExtratoEExcel(lancamentosTxt.Count, lancamentosExcel.Count, lancamentosTxt.Sum(l => l.Valor), lancamentosExcel.Sum(l => l.Valor));

            // Exibir lançamentos que estão no extrato e não estão no Excel
            conciliacaoService.ExibirLancamentosNoExtratoENaoNoExcel(lancamentosTxt, lancamentosExcel);

            // Exibir lançamentos que estão no Excel e não estão no extrato
            conciliacaoService.ExibirLancamentosNoExcelENaoNoExtrato(lancamentosExcel, lancamentosTxt, new List<LancamentoExcel>());

            // Exibir lançamentos que estão em ambos, mas têm diferença de valor de até 15 centavos
            var lancamentosComPequenaDiferenca = conciliacaoService.ExibirLancamentosComPequenaDiferenca(lancamentosTxt, lancamentosExcel);
        }
    }
}