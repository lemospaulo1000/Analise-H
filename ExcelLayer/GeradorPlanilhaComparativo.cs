using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AnaliseH3.Core.Models;

namespace AnaliseH3.ExcelLayer
{
    public class GeradorPlanilhaComparativo
    {
        public void Gerar(Application excelApp, List<ContaComparativa> contas)
        {
            Workbook wb = excelApp.ActiveWorkbook;

            if (wb == null)
                wb = excelApp.Workbooks.Add();

            Worksheet ws = null;

            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Comparativo")
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws != null)
                ws.Delete();

            ws = wb.Worksheets.Add();
            ws.Name = "Comparativo";

            ws.Cells[1, 1] = "Conta";
            ws.Cells[1, 2] = "Descrição";
            ws.Cells[1, 3] = "Saldo Anterior";
            ws.Cells[1, 4] = "Saldo Atual";
            ws.Cells[1, 5] = "Variação";
            ws.Cells[1, 6] = "Variação %";
            ws.Cells[1, 7] = "Material";
            ws.Cells[1, 8] = "Redutora";
            ws.Cells[1, 9] = "Selecionada";
            ws.Cells[1, 10] = "Classificação Risco";
            ws.Cells[1, 11] = "Observação";

            int total = contas.Count;

            object[,] dados = new object[total, 11];

            for (int i = 0; i < total; i++)
            {
                var c = contas[i];

                bool ehAnalitica = !c.Codigo.EndsWith("00");

                dados[i, 0] = c.Codigo;
                dados[i, 1] = c.Descricao;
                dados[i, 2] = c.MediaBase;
                dados[i, 3] = c.SaldoAtual;
                dados[i, 4] = c.Variacao;
                dados[i, 5] = c.VariacaoPercentual;

                dados[i, 6] = ehAnalitica ? (c.Material ? "Sim" : "Não") : "";
                dados[i, 7] = c.EhRedutora ? "Sim" : "Não";
                dados[i, 8] = ehAnalitica ? (c.Selecionada ? "Sim" : "Não") : "";

                dados[i, 9] = c.ClassificacaoRisco;
                dados[i, 10] = c.Observacao;
            }

            Range destino = ws.Range["A2"].Resize[total, 11];

            ws.Range["C:E"].NumberFormat = "#,##0.00";
            ws.Range["F:F"].NumberFormat = "0.00%";

            destino.Value2 = dados;

            ws.Columns[2].ColumnWidth = 70;

            Range header = ws.Range["A1:K1"];

            header.Font.Bold = true;
            header.Interior.ColorIndex = 15;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // ===== ZEBRADO =====

            for (int i = 2; i <= total + 1; i += 2)
            {
                Range linha = ws.Range["A" + i, "K" + i];
                linha.Interior.Color = 242 + (242 * 256) + (242 * 65536);
            }

            // congelar cabeçalho
            ws.Cells[1, 1].Select();

            excelApp.ActiveWindow.SplitRow = 1;
            excelApp.ActiveWindow.FreezePanes = true;

            header.AutoFilter(1);

            // AutoFit após filtro
            ws.Columns["A:A"].AutoFit();
            ws.Columns["C:F"].AutoFit();
            ws.Columns["G:K"].AutoFit();
        }
    }
}