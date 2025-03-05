using System;

namespace ConciliarApp.Models
{
    public class LancamentoExcel
    {
        public DateTime Data { get; set; }
        public decimal Valor { get; set; }
        public string Descricao { get; set; }
        public bool DiferencaDePequenoValor { get; set; }
        public bool NaoExisteNoExtrato { get; set; }
    }
}