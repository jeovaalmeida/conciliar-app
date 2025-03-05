using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using ConciliarApp.Services;
using ConciliarApp.Models; // Adicione esta linha para usar as classes LancamentoExcel e LancamentoExtrato

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
            (int qtdLancamentosExcel, decimal totalExcel, HashSet<LancamentoExcel> lancamentosExcel) = conciliacaoService.ProcessarArquivoExcel(caminhoArquivoExcel, cartao);

            // Processar o arquivo TXT
            (int qtdLancamentosTxt, decimal totalTxt, List<LancamentoExtrato> lancamentosTxt) = conciliacaoService.ProcessarArquivoTxt(caminhoArquivoTxt);

            // Exibir a diferença entre o Excel e o TXT
            conciliacaoService.ExibirDiferencaEntreExtratoEExcel(qtdLancamentosTxt, qtdLancamentosExcel, totalTxt, totalExcel);

            // Exibir lançamentos que estão no extrato e não estão no Excel
            conciliacaoService.ExibirLancamentosNoExtratoENaoNoExcel(lancamentosTxt, lancamentosExcel);

            // Exibir lançamentos que estão em ambos, mas têm diferença de valor de até 15 centavos
            var lancamentosComPequenaDiferenca = conciliacaoService.ExibirLancamentosComPequenaDiferenca(lancamentosTxt, lancamentosExcel);

            // Exibir lançamentos que estão no Excel e não estão no extrato
            conciliacaoService.ExibirLancamentosNoExcelENaoNoExtrato(lancamentosExcel, lancamentosTxt, lancamentosComPequenaDiferenca);
        }
    }
}