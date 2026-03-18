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
        public bool ContaNova { get; set; }
        public bool ContaRemovida { get; set; } // 🔹 NOVO

        public string Observacao { get; set; }

        public int ScoreRisco { get; set; }

        // 🔹 AGORA DERIVADO (não setável manualmente)
        public string ClassificacaoRisco
        {
            get
            {
                if (ScoreRisco >= 70) return "Alto";
                if (ScoreRisco >= 40) return "Médio";
                return "Baixo";
            }
        }

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
            ContaNova = false;
            ContaRemovida = false; // 🔹 NOVO

            Observacao = "";
            ScoreRisco = 0;
        }
    }
}