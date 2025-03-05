using System;

namespace ConciliarApp.Models
{
    public class LancamentoExtrato
    {
        public DateTime Data { get; set; }
        public decimal Valor { get; set; }
        public string Descricao { get; set; }
        public bool ExisteNoExcel { get; set; }
    }
}