using AnaliseH3.Core.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace AnaliseH3.ExcelLayer
{
    public class LeitorBalanceteExcel
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

                // ----- Ler Competência (D7) -----

                var competenciaRaw = worksheet.Range["D7"].Value2?.ToString()?.Trim();

                if (string.IsNullOrWhiteSpace(competenciaRaw))
                    throw new InvalidOperationException("Competência (D7) não encontrada.");

                var partesData = competenciaRaw.Split('/');

                if (partesData.Length != 2)
                    throw new InvalidOperationException("Formato inválido de competência.");

                int mes = int.Parse(partesData[0]);
                int ano = int.Parse(partesData[1]);

                var periodoContabil = new PeriodoContabil(mes, ano);

                var balancete = new Balancete(identificador, periodoContabil);

                // --------------------------------------------------
                // DESCOBRIR ÚLTIMA LINHA REAL
                // --------------------------------------------------

                int ultimaLinha =
                    worksheet.Cells[worksheet.Rows.Count, 1]
                    .End(XlDirection.xlUp)
                    .Row;

                if (ultimaLinha < 10)
                    return balancete;

                // --------------------------------------------------
                // LER BLOCO DE DADOS EM UMA ÚNICA OPERAÇÃO
                // --------------------------------------------------

                Range range = worksheet.Range["A1", $"F{ultimaLinha}"];

                object[,] dados = range.Value2;

                for (int linha = 10; linha <= ultimaLinha; linha++)
                {
                    var contaRaw = dados[linha, 1]?.ToString()?.Trim();

                    if (string.IsNullOrWhiteSpace(contaRaw))
                        continue;

                    // Divide código e descrição pelo separador " - "
                    string[] partes = contaRaw.Split(new string[] { " - " }, StringSplitOptions.None);

                    string codigo = partes[0].Trim();

                    // garante que estamos nas classes contábeis 1,2,3 ou 4
                    if (codigo.Length == 0 || !"1234".Contains(codigo[0]))
                        continue;

                    string descricao = partes.Length > 1 ? partes[1].Trim() : "";

                    bool ehRedutora = descricao.Contains("(-)");

                    double saldoAtual = 0.0;

                    if (dados[linha, 5] != null)
                    {
                        if (dados[linha, 5] is double valor)
                            saldoAtual = valor;
                        else
                            saldoAtual = Convert.ToDouble(dados[linha, 5]);
                    }

                    var dc = dados[linha, 6]?.ToString()?.Trim();

                    if (dc == "D")
                        saldoAtual = -Math.Abs(saldoAtual);
                    else if (dc == "C")
                        saldoAtual = Math.Abs(saldoAtual);
                    else
                        continue;

                    var conta = new ContaPeriodo(codigo, descricao, saldoAtual);

                    conta.EhRedutora = ehRedutora;

                    balancete.AdicionarConta(conta);
                }

                return balancete;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
            }
        }
    }
}