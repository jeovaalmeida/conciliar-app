using System;

namespace ConciliarApp.Models
{
    public class LancamentoExtrato
    {
        public DateTime Data { get; set; }
        public decimal Valor { get; set; }
        private string _descricao;
        public string Descricao 
        { 
            get => _descricao; 
            set
            {
                _descricao = value;
                var pos = _descricao.LastIndexOf("   CAMPO GRANDE"); // remover   CAMPO GRANDE MS e CAMPO GRANDE BRA muito comum nas descrições
                if (pos > 0)
                    _descricao = _descricao.Substring(0, pos).Trim();
            } 
        }
        public bool ExisteNoExcel { get; set; }
    }
}