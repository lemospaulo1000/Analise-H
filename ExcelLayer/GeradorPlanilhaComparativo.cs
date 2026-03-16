using AnaliseH3.Core.Models;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

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
            ws.Cells[1, 10] = "ContaNova";
            ws.Cells[1, 11] = "Classificação Risco";
            ws.Cells[1, 12] = "Observação";

            int total = contas.Count;

            object[,] dados = new object[total, 12];

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
                dados[i, 6] = c.Material ? "Sim" : "Não";
                dados[i, 7] = c.EhRedutora ? "Sim" : "Não";
                dados[i, 8] = ehAnalitica ? (c.Selecionada ? "Sim" : "Não") : "";

                dados[i, 9] = c.ContaNova ? "Sim" : "Não";

                dados[i, 10] = c.ClassificacaoRisco;
                dados[i, 11] = c.Observacao;
            }

            Range destino = ws.Range["A2"].Resize[total, 12];

            ws.Range["C:E"].NumberFormat = "#,##0.00";
            ws.Range["F:F"].NumberFormat = "0.00%";

            destino.Value2 = dados;

            ws.Columns[2].ColumnWidth = 70;

            Range header = ws.Range["A1:L1"];

            header.Font.Bold = true;
            header.Interior.ColorIndex = 15;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // ===== ZEBRADO =====

            for (int i = 2; i <= total + 1; i += 2)
            {
                Range linha = ws.Range["A" + i, "L" + i];
                linha.Interior.Color = 242 + (242 * 256) + (242 * 65536);
            }

            // congelar cabeçalho
            ws.Cells[1, 1].Select();

            excelApp.ActiveWindow.SplitRow = 1;
            excelApp.ActiveWindow.FreezePanes = true;

            header.AutoFilter(1);

            // agrupara por nivel hierarquico

            for (int i = 0; i < total; i++)
            {
                var conta = contas[i].Codigo;

                int zeros = conta.Reverse().TakeWhile(c => c == '0').Count();

                int nivel;

                switch (zeros)
                {
                    case 8: nivel = 1; break; // Classe
                    case 7: nivel = 2; break; // Grupo
                    case 6: nivel = 3; break; // Subgrupo
                    case 5: nivel = 4; break; // Título
                    case 4: nivel = 5; break; // Subtítulo
                    case 3: nivel = 6; break; // Item
                    default: nivel = 7; break; // Subitem
                }

                int linha = i + 2;

                ws.Rows[linha].OutlineLevel = nivel;
            }

            ws.Outline.ShowLevels(RowLevels: 3);

            // AutoFit após filtro
            ws.Columns["A:A"].AutoFit();
            ws.Columns["C:F"].AutoFit();
            ws.Columns["G:L"].AutoFit();
        }
    }
}