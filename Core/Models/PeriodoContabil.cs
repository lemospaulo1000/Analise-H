using System;

namespace AnaliseH3.Core.Models
{
    public class PeriodoContabil
    {
        public int Mes { get; }
        public int Ano { get; }

        public PeriodoContabil(int mes, int ano)
        {
            if (mes < 1 || mes > 14)
                throw new ArgumentException("Mês contábil deve estar entre 1 e 14.");

            if (ano < 1900 || ano > 2100)
                throw new ArgumentException("Ano inválido.");

            Mes = mes;
            Ano = ano;
        }

        public override string ToString()
        {
            return $"{Mes:00}/{Ano}";
        }
    }
}