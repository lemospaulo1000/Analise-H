using System.Collections.Generic;
using System.Linq;
using AnaliseH3.Core.Models;

namespace AnaliseH3.Core.Services
{
    public class DetectorReclassificacoes
    {
        public List<ReclassificacaoConta> Detectar(
            List<ContaPeriodo> contasRemovidas,
            List<ContaComparativa> comparativo)
        {
            var contasNovas = comparativo
                .Where(c => c.ContaNova)
                .ToList();

            var novasUtilizadas = new HashSet<string>();

            var resultado = new List<ReclassificacaoConta>();

            foreach (var removida in contasRemovidas)
            {
                if (string.IsNullOrWhiteSpace(removida.Codigo))
                    continue;

                if (removida.Codigo.Length < 7)
                    continue;

                string subtituloRemovido = removida.Codigo.Substring(0, 5);
                string itemRemovido = removida.Codigo.Substring(5, 2);

                var candidata = contasNovas
                    .FirstOrDefault(n =>
                        !novasUtilizadas.Contains(n.Codigo) &&
                        n.Codigo.Length >= 7 &&

                        // mesmo subtítulo
                        n.Codigo.Substring(0, 5) == subtituloRemovido &&

                        // item diferente
                        n.Codigo.Substring(5, 2) != itemRemovido &&

                        // descrição semelhante
                        DescricoesSemelhantes(removida.Descricao, n.Descricao));

                if (candidata != null)
                {
                    resultado.Add(new ReclassificacaoConta
                    {
                        ContaAntiga = removida.Codigo,
                        ContaNova = candidata.Codigo,
                        Descricao = candidata.Descricao,
                        SaldoAntigo = removida.SaldoAtual,
                        SaldoNovo = candidata.SaldoAtual
                    });

                    novasUtilizadas.Add(candidata.Codigo);
                }
            }

            return resultado;
        }

        private bool DescricoesSemelhantes(string a, string b)
        {
            var d1 = Normalizar(a);
            var d2 = Normalizar(b);

            if (d1 == d2)
                return true;

            if (d1.StartsWith(d2) || d2.StartsWith(d1))
                return true;

            var palavras1 = d1.Split(' ');
            var palavras2 = d2.Split(' ');

            int iguais = palavras1.Intersect(palavras2).Count();

            return iguais >= 2;
        }

        private string Normalizar(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return "";

            return string.Join(" ",
                texto
                    .Trim()
                    .ToLower()
                    .Split(' ')
                    .Where(p => !string.IsNullOrWhiteSpace(p))
            );
        }
    }
}