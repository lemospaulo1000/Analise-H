namespace AnaliseH3.Core.Models
{
    public class ReclassificacaoConta
    {
        public string ContaAntiga { get; set; }
        public string ContaNova { get; set; }

        public string Descricao { get; set; }

        public double SaldoAntigo { get; set; }
        public double SaldoNovo { get; set; }
    }
}