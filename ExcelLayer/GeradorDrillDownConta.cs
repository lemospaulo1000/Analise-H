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
            var app = workbook.Application;

            bool screen = app.ScreenUpdating;
            app.ScreenUpdating = false;

            try
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

                // =============================
                // DADOS (ARRAY EM MEMÓRIA)
                // =============================

                int total = contasFiltradas.Count;

                object[,] dados = new object[total, 9];

                for (int i = 0; i < total; i++)
                {
                    var c = contasFiltradas[i];

                    dados[i, 0] = c.Codigo;
                    dados[i, 1] = c.Descricao;
                    dados[i, 2] = c.MediaBase;
                    dados[i, 3] = c.SaldoAtual;
                    dados[i, 4] = c.Variacao;
                    dados[i, 5] = c.VariacaoPercentual;
                    dados[i, 6] = c.Material ? "Sim" : "Não";
                    dados[i, 7] = c.EhRedutora ? "Sim" : "Não";
                    dados[i, 8] = c.Selecionada ? "Sim" : "Não";
                }

                Range destino = ws.Range["A4"].Resize[total, 9];

                destino.Value2 = dados;

                // =============================
                // FORMATAÇÃO
                // =============================

                ws.Columns["C:E"].NumberFormat = "#,##0.00";
                ws.Columns["F:F"].NumberFormat = "0.00%";

                ws.Columns.AutoFit();
            }
            finally
            {
                app.ScreenUpdating = screen;
            }
        }
    }
}