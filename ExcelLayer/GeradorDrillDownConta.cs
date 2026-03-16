using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using AnaliseH3.Core.Models;

namespace AnaliseH3.ExcelLayer
{
    public class GeradorDrillDownConta
    {
        public void Gerar(
            Workbook workbook,
            string codigoBase,
            List<ContaComparativa> contas)
        {
            string prefixo = codigoBase.TrimEnd('0');

            if (string.IsNullOrEmpty(prefixo))
                prefixo = codigoBase.Substring(0, 1);

            var contasFiltradas = contas
                .Where(c => c.Codigo.StartsWith(prefixo))
                .OrderBy(c => c.Codigo)
                .ToList();

            if (contasFiltradas.Count == 0)
                return;

            string nomePlanilha = "Drill_" + codigoBase;

            Worksheet ws = null;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == nomePlanilha)
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws == null)
            {
                ws = workbook.Worksheets.Add();
                ws.Name = nomePlanilha;
            }
            else
            {
                ws.Cells.Clear();
            }

            // =============================
            // BOTÃO VOLTAR
            // =============================

            ws.Cells[1, 1] = "← Voltar ao Comparativo";

            ws.Hyperlinks.Add(
                ws.Cells[1, 1],
                "",
                "'Comparativo'!A1",
                "",
                "← Voltar ao Comparativo"
            );

            ws.Range["A1"].Font.Bold = true;
            ws.Range["A1"].Font.Color = 16711680;

            // =============================
            // CABEÇALHO
            // =============================

            ws.Cells[3, 1] = "Conta";
            ws.Cells[3, 2] = "Descrição";
            ws.Cells[3, 3] = "Saldo Anterior";
            ws.Cells[3, 4] = "Saldo Atual";
            ws.Cells[3, 5] = "Variação";
            ws.Cells[3, 6] = "Variação %";
            ws.Cells[3, 7] = "Material";
            ws.Cells[3, 8] = "Redutora";
            ws.Cells[3, 9] = "Selecionada";

            int linha = 4;

            foreach (var c in contasFiltradas)
            {
                ws.Cells[linha, 1] = c.Codigo;
                ws.Cells[linha, 2] = c.Descricao;
                ws.Cells[linha, 3] = c.MediaBase;
                ws.Cells[linha, 4] = c.SaldoAtual;
                ws.Cells[linha, 5] = c.Variacao;
                ws.Cells[linha, 6] = c.VariacaoPercentual;
                ws.Cells[linha, 7] = c.Material ? "Sim" : "Não";
                ws.Cells[linha, 8] = c.EhRedutora ? "Sim" : "Não";
                ws.Cells[linha, 9] = c.Selecionada ? "Sim" : "Não";

                linha++;
            }

            ws.Columns["C:E"].NumberFormat = "#,##0.00";
            ws.Columns["F:F"].NumberFormat = "0.00%";

            ws.Columns.AutoFit();
        }
    }
}