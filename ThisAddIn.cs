using System;
using Excel = Microsoft.Office.Interop.Excel;
using AnaliseH3.ExcelLayer;
using AnaliseH3.Core.Models;
using System.Collections.Generic;

namespace AnaliseH3
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // inicialização do add-in
            this.Application.SheetBeforeDoubleClick += Application_SheetBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // =====================================================
        // EVENTO DE DUPLO CLIQUE NA PLANILHA
        // =====================================================

        private void Application_SheetBeforeDoubleClick(
            object Sh,
            Excel.Range Target,
            ref bool Cancel)
        {
            try
            {
                if (Target == null)
                    return;

                if (Target.Cells.Count > 1)
                    return;

                Excel.Worksheet ws = Sh as Excel.Worksheet;

                if (ws == null)
                    return;

                // Apenas na planilha Comparativo
                if (!string.Equals(ws.Name, "Comparativo", StringComparison.OrdinalIgnoreCase))
                    return;

                // Apenas coluna Conta
                if (Target.Column != 1)
                    return;

                // Ignorar cabeçalho
                if (Target.Row < 2)
                    return;

                var valor = Target.Value2;

                if (valor == null)
                    return;

                string codigo = valor.ToString();

                if (string.IsNullOrWhiteSpace(codigo))
                    return;

                // Cancelar comportamento padrão do Excel
                Cancel = true;

                var contas = ExecutorAnalise.ResultadoAnalise;

                // Caso o usuário tenha reaberto o arquivo
                if (contas == null)
                {
                    contas = ReconstruirContasDaPlanilha(ws);
                }

                if (contas == null || contas.Count == 0)
                    return;

                var gerador = new GeradorDrillDownConta();

                gerador.Gerar(
                    this.Application.ActiveWorkbook,
                    codigo,
                    contas
                );
            }
            catch
            {
                // evitar erro de automação do Excel
            }
        }

        // =====================================================
        // RECONSTRUIR CONTAS CASO ARQUIVO TENHA SIDO REABERTO
        // =====================================================

        private List<ContaComparativa> ReconstruirContasDaPlanilha(Excel.Worksheet ws)
        {
            var lista = new List<ContaComparativa>();

            int linha = 2;

            while (true)
            {
                var codigo = ws.Cells[linha, 1].Value2;

                if (codigo == null)
                    break;

                string conta = codigo.ToString();
                string descricao = ws.Cells[linha, 2].Value2?.ToString() ?? "";

                double saldoAnterior = 0;
                double saldoAtual = 0;

                var v1 = ws.Cells[linha, 3].Value2;
                var v2 = ws.Cells[linha, 4].Value2;

                if (v1 != null)
                    saldoAnterior = Convert.ToDouble(v1);

                if (v2 != null)
                    saldoAtual = Convert.ToDouble(v2);

                var c = new ContaComparativa(
                    conta,
                    descricao,
                    saldoAtual,
                    saldoAnterior,
                    0
                );

                lista.Add(c);

                linha++;
            }

            return lista;
        }

        #region Código gerado por VSTO

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}