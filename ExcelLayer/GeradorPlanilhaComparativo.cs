using AnaliseH3.Core.Models;
using Microsoft.Office.Interop.Excel;
using System;
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
            ws.Cells[1, 11] = "Score Risco";
            ws.Cells[1, 12] = "Classificação Risco";
            ws.Cells[1, 13] = "Observação";

            int total = contas.Count;

            object[,] dados = new object[total, 13];

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
                dados[i, 10] = c.ScoreRisco;
                dados[i, 11] = c.ClassificacaoRisco;
                dados[i, 12] = c.Observacao;
            }

            Range destino = ws.Range["A2"].Resize[total, 13];

            ws.Range["C:E"].NumberFormat = "#,##0.00";
            ws.Range["F:F"].NumberFormat = "0.00%";

            destino.Value2 = dados;

            ws.Columns[2].ColumnWidth = 70;

            Range header = ws.Range["A1:M1"];

            header.Font.Bold = true;
            header.Interior.ColorIndex = 15;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // ===== ZEBRADO =====
            for (int i = 2; i <= total + 1; i += 2)
            {
                Range linhaZebra = ws.Range["A" + i, "M" + i];
                linhaZebra.Interior.Color = 242 + (242 * 256) + (242 * 65536);
            }

            // congelar cabeçalho
            ws.Cells[1, 1].Select();

            excelApp.ActiveWindow.SplitRow = 1;
            excelApp.ActiveWindow.FreezePanes = true;

            header.AutoFilter(1);

            // ===== AGRUPAMENTO HIERÁRQUICO =====
            for (int i = 0; i < total; i++)
            {
                var contaObj = contas[i];
                string conta = contaObj.Codigo;

                int zeros = conta.Reverse().TakeWhile(c => c == '0').Count();

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

                int linha = i + 2;

                ws.Rows[linha].OutlineLevel = nivel;

                if (nivel == 3 && contaObj.Material)
                {
                    Range linhaExcel = ws.Rows[linha];
                    linhaExcel.Font.Bold = true;
                    linhaExcel.Interior.Color = 13434879;
                }
            }

            // expandir subgrupos
            for (int linha = 2; linha <= total + 1; linha++)
            {
                Range linhaExcel = ws.Rows[linha];

                if (linhaExcel.Font.Bold)
                {
                    try
                    {
                        linhaExcel.ShowDetail = true;
                    }
                    catch { }
                }
            }

            ws.Outline.ShowLevels(RowLevels: 3);

            ws.Columns["A:A"].AutoFit();
            ws.Columns["C:F"].AutoFit();
            ws.Columns["G:M"].AutoFit();

            // ===== RADAR =====
            GerarRadarDistorcoes(wb, contas);
        }

        private void GerarRadarDistorcoes(Workbook wb, List<ContaComparativa> contas)
        {
            Worksheet ws = null;

            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "RadarDistorcoes")
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws != null)
                ws.Delete();

            ws = wb.Worksheets.Add();
            ws.Name = "RadarDistorcoes";

            var topContas = contas
                .OrderByDescending(c => c.ScoreRisco)
                .ThenByDescending(c => Math.Abs(c.Variacao))
                .Take(20)
                .ToList();

            ws.Cells[1, 1] = "Conta";
            ws.Cells[1, 2] = "Descrição";
            ws.Cells[1, 3] = "Saldo Atual";
            ws.Cells[1, 4] = "Variação";
            ws.Cells[1, 5] = "Score Risco";
            ws.Cells[1, 6] = "Classificação";

            int total = topContas.Count;

            object[,] dados = new object[total, 6];

            for (int i = 0; i < total; i++)
            {
                var c = topContas[i];

                dados[i, 0] = c.Codigo;
                dados[i, 1] = c.Descricao;
                dados[i, 2] = c.SaldoAtual;
                dados[i, 3] = c.Variacao;
                dados[i, 4] = c.ScoreRisco;
                dados[i, 5] = c.ClassificacaoRisco;
            }

            Range destino = ws.Range["A2"].Resize[total, 6];
            destino.Value2 = dados;

            ws.Range["C:D"].NumberFormat = "#,##0.00";

            Range header = ws.Range["A1:F1"];
            header.Font.Bold = true;
            header.Interior.ColorIndex = 15;

            ws.Columns["A:F"].AutoFit();

            // 🔹 CHAMA RADAR POR SUBGRUPO
            GerarRadarSubgrupos(wb, contas);
        }

        private void GerarRadarSubgrupos(Workbook wb, List<ContaComparativa> contas)
        {
            Worksheet ws = null;

            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "RadarSubgrupos")
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws != null)
                ws.Delete();

            ws = wb.Worksheets.Add();
            ws.Name = "RadarSubgrupos";

            var grupos = contas
                .Where(c => !string.IsNullOrEmpty(c.Codigo) && c.Codigo.Length >= 3)
                .GroupBy(c => c.Codigo.Substring(0, 3))
                .Select(g => new
                {
                    Subgrupo = g.Key,
                    ScoreTotal = g.Sum(x => x.ScoreRisco),
                    VariacaoTotal = g.Sum(x => x.Variacao),
                    QuantidadeContas = g.Count()
                })
                .OrderByDescending(g => g.ScoreTotal)
                .ToList();

            ws.Cells[1, 1] = "Subgrupo";
            ws.Cells[1, 2] = "Score Total";
            ws.Cells[1, 3] = "Variação Total";
            ws.Cells[1, 4] = "Qtd Contas";

            int total = grupos.Count;

            object[,] dados = new object[total, 4];

            for (int i = 0; i < total; i++)
            {
                var g = grupos[i];

                dados[i, 0] = g.Subgrupo;
                dados[i, 1] = g.ScoreTotal;
                dados[i, 2] = g.VariacaoTotal;
                dados[i, 3] = g.QuantidadeContas;
            }

            Range destino = ws.Range["A2"].Resize[total, 4];
            destino.Value2 = dados;

            ws.Range["C:C"].NumberFormat = "#,##0.00";

            Range header = ws.Range["A1:D1"];
            header.Font.Bold = true;
            header.Interior.ColorIndex = 15;

            ws.Columns["A:D"].AutoFit();
        }
    }
}