using System;
using System.Collections.Generic;
using System.Linq;
using AnaliseH3.Core.Models;

namespace AnaliseH3.Core.Services
{
    public class AnaliseComparativa
    {
        private readonly List<Balancete> _periodos;
        private Balancete _periodoAtual;
        private double? _limiteMaterialidade;

        public IReadOnlyList<Balancete> Periodos => _periodos.AsReadOnly();
        public Balancete PeriodoAtual => _periodoAtual;
        public double? LimiteMaterialidade => _limiteMaterialidade;

        public AnaliseComparativa()
        {
            _periodos = new List<Balancete>();
        }

        public void AdicionarBalancete(Balancete balancete)
        {
            if (balancete == null)
                throw new ArgumentNullException(nameof(balancete));

            _periodos.Add(balancete);
        }

        public void DefinirPeriodoAtual(string identificador)
        {
            if (string.IsNullOrWhiteSpace(identificador))
                throw new ArgumentException("Identificador inválido.");

            var periodo = _periodos.FirstOrDefault(p => p.Identificador == identificador);

            if (periodo == null)
                throw new InvalidOperationException("Período não encontrado.");

            _periodoAtual = periodo;
        }

        public void DefinirLimiteMaterialidade(double valor)
        {
            if (valor < 0)
                throw new ArgumentException("Limite de materialidade não pode ser negativo.");

            _limiteMaterialidade = valor;
        }

        public List<ContaComparativa> ExecutarComparacao()
        {
            if (_periodos.Count < 2)
                throw new InvalidOperationException("É necessário pelo menos dois períodos para comparar.");

            if (_periodoAtual == null)
                throw new InvalidOperationException("Período atual não definido.");

            if (!_limiteMaterialidade.HasValue)
                throw new InvalidOperationException("Limite de materialidade não definido.");

            var periodosBase = _periodos.Where(p => p != _periodoAtual).ToList();

            var todosCodigos = new HashSet<string>();

            foreach (var periodo in _periodos)
                foreach (var codigo in periodo.Contas.Keys)
                    todosCodigos.Add(codigo);

            var resultado = new List<ContaComparativa>();

            foreach (var codigo in todosCodigos)
            {
                bool existeAtual = _periodoAtual.ContemConta(codigo);

                double saldoAtual = existeAtual
                    ? _periodoAtual.ObterConta(codigo).SaldoAtual
                    : 0.0;

                var saldosBase = periodosBase
                    .Select(p => p.ContemConta(codigo) ? p.ObterConta(codigo).SaldoAtual : 0.0)
                    .ToList();

                double mediaBase = saldosBase.Count > 0
                    ? saldosBase.Average()
                    : 0.0;

                string descricao = existeAtual
                    ? _periodoAtual.ObterConta(codigo).Descricao
                    : periodosBase
                        .Select(p => p.ObterConta(codigo))
                        .FirstOrDefault(c => c != null)?.Descricao ?? string.Empty;

                var contaComparativa = new ContaComparativa(
                    codigo,
                    descricao,
                    saldoAtual,
                    mediaBase,
                    _limiteMaterialidade.Value
                );

                // ---------------------------------
                // DETECÇÃO DE CONTA NOVA
                // ---------------------------------

                bool contaNova =
                    existeAtual &&
                    !periodosBase.Any(p => p.ContemConta(codigo));

                contaComparativa.ContaNova = contaNova;

                // ---------------------------------
                // REGRA DE SELEÇÃO
                // ---------------------------------

                bool variacaoRelevante =
                    contaComparativa.VariacaoPercentual >= 0.10;

                contaComparativa.Selecionada =
                    contaComparativa.Material && variacaoRelevante;

                resultado.Add(contaComparativa);
            }

            return resultado
                .OrderBy(c => c.Codigo)
                .ToList();
        }
    }
}