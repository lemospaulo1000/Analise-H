using Microsoft.Office.Interop.Excel;

namespace AnaliseH3.ExcelLayer
{
    public class GeradorMaterialidade
    {
        public void Gerar(Application excelApp, double totalAtivo)
        {
            Workbook wb = excelApp.ActiveWorkbook;

            if (wb == null)
                wb = excelApp.Workbooks.Add();

            Worksheet ws = null;

            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Materialidade")
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws != null)
                ws.Delete();

            ws = wb.Worksheets.Add();
            ws.Name = "Materialidade";

            // cabeçalhos
            ws.Range["A1"].Value = "Conta";
            ws.Range["B1"].Value = "Saldo";

            ws.Range["A2"].Value = "Total do Ativo";
            ws.Range["A3"].Value = "Materialidade (1%)";

            ws.Range["B2"].Value = totalAtivo;

            // fórmula
            ws.Range["B3"].Formula = "=B2*1%";

            // formatação cabeçalho
            var header = ws.Range["A1:B1"];

            header.Font.Bold = true;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            header.Interior.ColorIndex = 15;

            // formato numérico
            ws.Range["B2"].NumberFormat = "#,##0.00";
            ws.Range["B3"].NumberFormat = "#,##0.00";

            // largura fixa (evita bug de AutoFit no Interop)
            ws.Columns["A"].ColumnWidth = 18;
            ws.Columns["B"].ColumnWidth = 18;
        }
    }
}