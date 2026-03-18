using System;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using AnaliseH3.ExcelLayer;
using AnaliseH3.Core.Services;
using AnaliseH3.Core.Models;
using System.Collections.Generic;

namespace AnaliseH3
{
    public class ExecutorAnalise
    {
        public static List<ContaComparativa> ResultadoAnalise { get; private set; }

        public void Executar(Excel.Application excelApp)
        {
            var app = excelApp;

            bool screen = app.ScreenUpdating;
            bool events = app.EnableEvents;
            bool alerts = app.DisplayAlerts;
            Excel.XlCalculation calc = app.Calculation;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                // =============================
                // CRIAR NOVO WORKBOOK PARA A ANÁLISE
                // =============================

                var wbNovo = app.Workbooks.Add();
                wbNovo.Activate();

                // =============================
                // SELECIONAR BALANCETES
                // =============================

                var caminhos = SelecionarArquivos();

                if (caminhos == null)
                    return;

                var balancetes = LerBalancetes(app, caminhos.Item1, caminhos.Item2);

                var analise = CriarAnalise(balancetes.Item1, balancetes.Item2);

                double saldoAtivo = ObterSaldoAtivo(balancetes.Item2);

                double materialidade = saldoAtivo * 0.01;

                analise.DefinirLimiteMaterialidade(materialidade);

                var resultado = analise.ExecutarComparacao();
                ResultadoAnalise = resultado;

                // =============================
                // GERAR PLANILHAS
                // =============================

                GerarComparativo(app, resultado);

                GerarMaterialidade(app, saldoAtivo);

                var contasRemovidas = DetectarContasRemovidas(analise);

                GerarPlanilhaRemovidas(app, contasRemovidas);

                var reclassificacoes = DetectarReclassificacoes(contasRemovidas, resultado);

                GerarPlanilhaReclassificacoes(app, reclassificacoes);

                // =============================
                // ORGANIZAR ABAS
                // =============================

                OrganizarAbas(app);

                AtivarComparativo(app);

                MostrarResultado(resultado.Count, contasRemovidas.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Erro");
            }
            finally
            {
                // restaurar ambiente do Excel
                app.ScreenUpdating = screen;
                app.EnableEvents = events;
                app.DisplayAlerts = alerts;
                app.Calculation = calc;
            }
        }

        private Tuple<string, string> SelecionarArquivos()
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Arquivos Excel (*.xls;*.xlsx)|*.xls;*.xlsx";

            dialog.Title = "Selecione o Balancete Anterior";

            if (dialog.ShowDialog() != DialogResult.OK)
                return null;

            var anterior = dialog.FileName;

            dialog.Title = "Selecione o Balancete Atual";

            if (dialog.ShowDialog() != DialogResult.OK)
                return null;

            var atual = dialog.FileName;

            if (anterior == atual)
            {
                MessageBox.Show(
                    "O balancete anterior e o atual são o mesmo arquivo.\n\nSelecione arquivos diferentes.",
                    "Análise-H 3.0"
                );
                return null;
            }

            return Tuple.Create(anterior, atual);
        }

        private Tuple<Balancete, Balancete> LerBalancetes(
            Excel.Application excelApp,
            string anterior,
            string atual)
        {
            var leitor = new LeitorBalanceteExcel();

            var balanceteAnterior = leitor.Ler(excelApp, anterior, "Anterior");
            var balanceteAtual = leitor.Ler(excelApp, atual, "Atual");

            return Tuple.Create(balanceteAnterior, balanceteAtual);
        }

        private AnaliseComparativa CriarAnalise(
            Balancete anterior,
            Balancete atual)
        {
            var analise = new AnaliseComparativa();

            analise.AdicionarBalancete(anterior);
            analise.AdicionarBalancete(atual);

            analise.DefinirPeriodoAtual("Atual");

            return analise;
        }

        private double ObterSaldoAtivo(Balancete balanceteAtual)
        {
            var contaAtivo = balanceteAtual.ObterConta("100000000");

            if (contaAtivo != null)
                return Math.Abs(contaAtivo.SaldoAtual);

            return 0;
        }

        private void GerarMaterialidade(
            Excel.Application excelApp,
            double saldoAtivo)
        {
            var geradorMaterialidade = new GeradorMaterialidade();
            geradorMaterialidade.Gerar(excelApp, saldoAtivo);
        }

        private void GerarComparativo(
            Excel.Application excelApp,
            List<ContaComparativa> resultado)
        {
            var gerador = new GeradorPlanilhaComparativo();
            gerador.Gerar(excelApp, resultado);
        }

        private List<ContaPeriodo> DetectarContasRemovidas(
            AnaliseComparativa analise)
        {
            var detectorRemovidas = new DetectorContasRemovidas();

            var periodosBase = analise.Periodos
                .Where(p => p != analise.PeriodoAtual)
                .ToList();

            return detectorRemovidas.ObterContasRemovidas(
                analise.PeriodoAtual,
                periodosBase
            );
        }

        private void GerarPlanilhaRemovidas(
            Excel.Application excelApp,
            List<ContaPeriodo> contasRemovidas)
        {
            var workbook = excelApp.ActiveWorkbook;

            var geradorRemovidas = new GeradorContasRemovidas();

            geradorRemovidas.Gerar(
                workbook,
                contasRemovidas
            );
        }

        private List<ReclassificacaoConta> DetectarReclassificacoes(
            List<ContaPeriodo> contasRemovidas,
            List<ContaComparativa> comparativo)
        {
            var detector = new DetectorReclassificacoes();

            return detector.Detectar(
                contasRemovidas,
                comparativo
            );
        }

        private void GerarPlanilhaReclassificacoes(
            Excel.Application excelApp,
            List<ReclassificacaoConta> reclassificacoes)
        {
            var workbook = excelApp.ActiveWorkbook;

            var gerador = new GeradorReclassificacoes();

            gerador.Gerar(
                workbook,
                reclassificacoes
            );
        }

        private void OrganizarAbas(Excel.Application excelApp)
        {
            var wb = excelApp.ActiveWorkbook;

            if (wb == null)
                return;

            try
            {
                Excel.Worksheet comparativo = null;
                Excel.Worksheet materialidade = null;
                Excel.Worksheet removidas = null;
                Excel.Worksheet reclassificacoes = null;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws.Name == "Comparativo")
                        comparativo = ws;

                    else if (ws.Name == "Materialidade")
                        materialidade = ws;

                    else if (ws.Name == "ContasRemovidas")
                        removidas = ws;

                    else if (ws.Name == "Reclassificacoes")
                        reclassificacoes = ws;
                }

                if (comparativo != null)
                    comparativo.Move(Before: wb.Worksheets[1]);

                if (materialidade != null && comparativo != null)
                    materialidade.Move(After: comparativo);

                if (removidas != null && materialidade != null)
                    removidas.Move(After: materialidade);

                if (reclassificacoes != null && removidas != null)
                    reclassificacoes.Move(After: removidas);
            }
            catch
            {
            }
        }

        private void AtivarComparativo(Excel.Application excelApp)
        {
            var wb = excelApp.ActiveWorkbook;

            if (wb == null)
                return;

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == "Comparativo")
                {
                    ws.Activate();
                    ws.Cells[1, 1].Select();
                    return;
                }
            }
        }

        private void MostrarResultado(
            int contasAnalisadas,
            int contasRemovidas)
        {
            MessageBox.Show(
                $"Comparação executada.\n\n" +
                $"Contas analisadas: {contasAnalisadas}\n" +
                $"Contas removidas: {contasRemovidas}",
                "Análise-H 3.0"
            );
        }
    }
}