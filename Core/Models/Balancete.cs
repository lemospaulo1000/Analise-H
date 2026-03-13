using System;
using System.Collections.Generic;

namespace AnaliseH3.Core.Models
{
    public class Balancete
    {
        public string Identificador { get; }
        public PeriodoContabil Competencia { get; }

        private readonly Dictionary<string, ContaPeriodo> _contas;

        public IReadOnlyDictionary<string, ContaPeriodo> Contas => _contas;

        public Balancete(string identificador, PeriodoContabil competencia)
        {
            if (string.IsNullOrWhiteSpace(identificador))
                throw new ArgumentException("Identificador do período inválido.");

            Identificador = identificador.Trim();
            Competencia = competencia ?? throw new ArgumentNullException(nameof(competencia));

            _contas = new Dictionary<string, ContaPeriodo>();
        }

        public void AdicionarConta(ContaPeriodo conta)
        {
            if (conta == null)
                throw new ArgumentNullException(nameof(conta));

            _contas[conta.Codigo] = conta;
        }

        public bool ContemConta(string codigo)
        {
            return _contas.ContainsKey(codigo);
        }

        public ContaPeriodo ObterConta(string codigo)
        {
            return _contas.TryGetValue(codigo, out var conta) ? conta : null;
        }
    }
}