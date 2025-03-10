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
            if (args.Length < 3)
            {
                Console.WriteLine("Por favor, informe o nome do cartão (ex: MASTER, VISA), o nome do arquivo de extrato (TXT) e o nome da planilha (ex: 2025-03).");
                return;
            }

            string cartao = args[0].ToUpper();
            string nomeArquivoExtrato = args[1];
            string nomePlanilha = args[2];
            bool inserirLancamentos = args.Length > 3 && args[3].ToUpper() == "INSERIR";
            string caminhoOneDrive = @"C:\Users\jeova\OneDrive\FileSync\";
            string caminhoArquivoExtrato = Path.Combine(caminhoOneDrive + "Extratos", nomeArquivoExtrato);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string caminhoArquivoExcel = caminhoOneDrive + @"_ControleFin\RecDesp-Fixas-Parc-2018-2025.xlsx";

            var conciliacaoService = new ConciliacaoService();

            Console.WriteLine($"\r\n{DateTime.Now} | Iniciando processamento");

            var lancamentosProcessados = conciliacaoService.ExtrairEMarcarLancamentos(caminhoArquivoExcel, caminhoArquivoExtrato, cartao, nomePlanilha);

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

            // Inserir lançamentos do extrato que não estão no Excel, se o parâmetro INSERIR for fornecido
            if (inserirLancamentos)
            {
                var lancamentosNaoNoExcel = lancamentosProcessados.LancamentosExtrato.Where(l => !l.ExisteNoExcel).ToList();
                conciliacaoService.InserirLancamentosNoExcel(caminhoArquivoExcel, lancamentosNaoNoExcel, lancamentosProcessados.LinhaInsercao, nomePlanilha);
            }
        }
    }
}