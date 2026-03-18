using Microsoft.Office.Interop.Excel;
using System;

namespace AnaliseH3.ExcelLayer
{
    public class ComandoDestacarSubgrupos
    {
        public void Executar(Application excelApp)
        {
            Workbook wb = excelApp.ActiveWorkbook;

            if (wb == null)
                return;

            Worksheet comparativo = null;
            Worksheet materialidade = null;

            foreach (Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == "Comparativo")
                    comparativo = ws;

                if (ws.Name == "Materialidade")
                    materialidade = ws;
            }

            if (comparativo == null || materialidade == null)
                return;

            // ==========================
            // Ler materialidade
            // ==========================

            double limiteMaterialidade;

            try
            {
                limiteMaterialidade = Convert.ToDouble(
                    materialidade.Range["B3"].Value2
                );
            }
            catch
            {
                return;
            }

            int ultimaLinha =
                comparativo.Cells[comparativo.Rows.Count, 1]
                .End(XlDirection.xlUp)
                .Row;

            int total = ultimaLinha - 1;

            // ==========================
            // Ler dados em bloco
            // ==========================

            Range codigosRange = comparativo.Range["A2"].Resize[total, 1];
            Range saldoRange = comparativo.Range["D2"].Resize[total, 1];
            Range materialRange = comparativo.Range["G2"].Resize[total, 1];

            object[,] codigos = codigosRange.Value2;
            object[,] saldos = saldoRange.Value2;
            object[,] materialColuna = new object[total, 1];

            // ==========================
            // Recalcular materialidade
            // ==========================

            for (int i = 1; i <= total; i++)
            {
                if (saldos[i, 1] == null)
                {
                    materialColuna[i - 1, 0] = "";
                    continue;
                }

                double saldoAtual = Convert.ToDouble(saldos[i, 1]);

                bool material = Math.Abs(saldoAtual) >= limiteMaterialidade;

                materialColuna[i - 1, 0] = material ? "Sim" : "Não";
            }

            // escrever coluna material de volta
            materialRange.Value2 = materialColuna;

            // ==========================
            // Resetar formatação
            // ==========================

            Range linhas = comparativo.Range["A2"].Resize[total, 12];

            linhas.Font.Bold = false;
            linhas.Interior.Pattern = XlPattern.xlPatternNone;

            // ==========================
            // Recolher hierarquia
            // ==========================

            try
            {
                comparativo.Outline.ShowLevels(1);
                comparativo.Outline.ShowLevels(2);
                comparativo.Outline.ShowLevels(3);
            }
            catch
            {
            }

            // ==========================
            // Destacar subgrupos materiais
            // ==========================

            for (int i = 1; i <= total; i++)
            {
                if (codigos[i, 1] == null)
                    continue;

                string codigo = codigos[i, 1].ToString();

                int zeros = 0;

                for (int z = codigo.Length - 1; z >= 0; z--)
                {
                    if (codigo[z] == '0')
                        zeros++;
                    else
                        break;
                }

                int nivel;

                switch (zeros)
                {
                    case 8: nivel = 1; break;
                    case 7: nivel = 2; break;
                    case 6: nivel = 3; break;
                    case 5: nivel = 4; break;
                    case 4: nivel = 5; break;
                    case 3:
                    case 2: nivel = 6; break;
                    default: nivel = 7; break;
                }

                if (nivel != 3)
                    continue;

                if (materialColuna[i - 1, 0]?.ToString() != "Sim")
                    continue;

                Range linhaExcel = comparativo.Rows[i + 1];

                linhaExcel.Font.Bold = true;
                linhaExcel.Interior.Color = 13434879;
            }

            try
            {
                comparativo.Outline.ShowLevels(RowLevels: 3);
            }
            catch
            {
            }

            comparativo.Activate();
        }
    }
}