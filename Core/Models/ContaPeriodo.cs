using System;

namespace AnaliseH3.Core.Models
{
    public class ContaPeriodo
    {
        public string Codigo { get; }
        public string Descricao { get; }
        public double SaldoAtual { get; }
        public bool EhRedutora { get; set; }

        public ContaPeriodo(string codigo, string descricao, double saldoAtual)
        {
            if (string.IsNullOrWhiteSpace(codigo))
                throw new ArgumentException("Código da conta inválido.");

            Codigo = codigo.Trim();
            Descricao = descricao?.Trim() ?? string.Empty;
            SaldoAtual = saldoAtual;
        }
    }
}