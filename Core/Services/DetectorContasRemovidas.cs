using System.Collections.Generic;
using System.Linq;
using AnaliseH3.Core.Models;

namespace AnaliseH3.Core.Services
{
    public class DetectorContasRemovidas
    {
        public List<ContaPeriodo> ObterContasRemovidas(
            Balancete periodoAtual,
            List<Balancete> periodosBase)
        {
            var removidas = new List<ContaPeriodo>();

            foreach (var periodo in periodosBase)
            {
                foreach (var conta in periodo.Contas.Values)
                {
                    if (!periodoAtual.ContemConta(conta.Codigo))
                    {
                        removidas.Add(conta);
                    }
                }
            }

            return removidas
                .GroupBy(c => c.Codigo)
                .Select(g => g.First())
                .OrderBy(c => c.Codigo)
                .ToList();
        }
    }
}