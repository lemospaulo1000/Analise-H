using AnaliseH3.Core.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;

namespace AnaliseH3.ExcelLayer
{
    public class LeitorBalanceteExcelInterop
    {
        public Balancete Ler(Application excelApp, string caminhoArquivo, string identificador)
        {
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(
                    caminhoArquivo,
                    ReadOnly: true
                );

                worksheet = workbook.Worksheets[1];

                // =============================
                // LER COMPETÊNCIA (D7)
                // =============================

                var competenciaRaw = worksheet.Range["D7"].Value2?.ToString()?.Trim();

                if (string.IsNullOrWhiteSpace(competenciaRaw))
                    throw new InvalidOperationException("Competência não encontrada na célula D7.");

                var partes = competenciaRaw.Split('/');

                if (partes.Length != 2)
                    throw new InvalidOperationException("Formato de competência inválido.");

                int mes = int.Parse(partes[0]);
                int ano = int.Parse(partes[1]);

                var periodo = new PeriodoContabil(mes, ano);

                var balancete = new Balancete(identificador, periodo);

                // =============================
                // LEITURA DAS CONTAS
                // =============================

                int ultimaLinha =
                    worksheet.Cells[worksheet.Rows.Count, 1]
                    .End(XlDirection.xlUp)
                    .Row;

                for (int linha = 10; linha <= ultimaLinha; linha++)
                {
                    var contaRaw = worksheet.Cells[linha, 1].Value2?.ToString()?.Trim();

                    if (string.IsNullOrWhiteSpace(contaRaw))
                        continue;

                    string[] partesConta = contaRaw.Split(new string[] { " - " }, StringSplitOptions.None);

                    string codigo = partesConta[0].Trim();

                    if (codigo.Length == 0 || !"1234".Contains(codigo[0]))
                        continue;

                    string descricao = partesConta.Length > 1 ? partesConta[1].Trim() : "";

                    double saldoAtual = 0.0;

                    var valor = worksheet.Cells[linha, 5].Value2;

                    if (valor != null)
                        saldoAtual = Convert.ToDouble(valor);

                    var dc = worksheet.Cells[linha, 6].Value2?.ToString()?.Trim();

                    if (dc == "D")
                        saldoAtual = -Math.Abs(saldoAtual);
                    else if (dc == "C")
                        saldoAtual = Math.Abs(saldoAtual);
                    else
                        continue;

                    var conta = new ContaPeriodo(codigo, descricao, saldoAtual);

                    balancete.AdicionarConta(conta);
                }

                return balancete;
            }
            finally
            {
                if (workbook != null)
                    workbook.Close(false);
            }
        }
    }
}