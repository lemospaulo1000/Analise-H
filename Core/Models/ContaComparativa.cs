using System;

namespace AnaliseH3.Core.Models
{
    public class ContaComparativa
    {
        public string Codigo { get; }
        public string Descricao { get; }

        public double MediaBase { get; }
        public double SaldoAtual { get; }
        public double Variacao { get; }

        public double VariacaoPercentual { get; }

        public bool Material { get; }
        public bool EhRedutora { get; set; }
        public bool Selecionada { get; set; }

        public string ClassificacaoRisco { get; set; }
        public string Observacao { get; set; }

        public ContaComparativa(
            string codigo,
            string descricao,
            double saldoAtual,
            double mediaBase,
            double limiteMaterialidade)
        {
            Codigo = codigo;
            Descricao = descricao;

            SaldoAtual = saldoAtual;
            MediaBase = mediaBase;

            Variacao = saldoAtual - mediaBase;

            if (mediaBase != 0)
                VariacaoPercentual = Math.Abs(Variacao) / Math.Abs(mediaBase);
            else
                VariacaoPercentual = 0;

            Material = Math.Abs(SaldoAtual) >= limiteMaterialidade;

            EhRedutora = false;
            Selecionada = false;
            ClassificacaoRisco = "";
            Observacao = "";
        }
    }
}