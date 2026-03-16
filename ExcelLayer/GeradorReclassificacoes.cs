using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AnaliseH3.Core.Models;

namespace AnaliseH3.ExcelLayer
{
    public class GeradorReclassificacoes
    {
        public void Gerar(Workbook wb, List<ReclassificacaoConta> reclassificacoes)
        {
            Worksheet ws = null;

            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Reclassificacoes")
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws != null)
                ws.Delete();

            ws = wb.Worksheets.Add();
            ws.Name = "Reclassificacoes";

            ws.Cells[1, 1] = "Conta Antiga";
            ws.Cells[1, 2] = "Conta Nova";
            ws.Cells[1, 3] = "Descrição";
            ws.Cells[1, 4] = "Saldo Antigo";
            ws.Cells[1, 5] = "Saldo Novo";

            int total = reclassificacoes.Count;

            object[,] dados = new object[total, 5];

            for (int i = 0; i < total; i++)
            {
                var r = reclassificacoes[i];

                dados[i, 0] = r.ContaAntiga;
                dados[i, 1] = r.ContaNova;
                dados[i, 2] = r.Descricao;
                dados[i, 3] = r.SaldoAntigo;
                dados[i, 4] = r.SaldoNovo;
            }

            if (total > 0)
            {
                Range destino = ws.Range["A2"].Resize[total, 5];
                destino.Value2 = dados;
            }

            ws.Range["D:E"].NumberFormat = "#,##0.00";

            Range header = ws.Range["A1:E1"];

            header.Font.Bold = true;
            header.Interior.ColorIndex = 15;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            ws.Columns.AutoFit();
        }
    }
}