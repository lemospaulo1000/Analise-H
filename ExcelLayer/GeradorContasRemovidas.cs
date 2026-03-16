using System.Collections.Generic;
using AnaliseH3.Core.Models;
using Microsoft.Office.Interop.Excel;

namespace AnaliseH3.ExcelLayer
{
    public class GeradorContasRemovidas
    {
        public void Gerar(Workbook workbook, List<ContaPeriodo> contasRemovidas)
        {
            Worksheet planilha = workbook.Worksheets.Add();
            planilha.Name = "ContasRemovidas";

            planilha.Cells[1, 1] = "Conta";
            planilha.Cells[1, 2] = "Descrição";
            planilha.Cells[1, 3] = "Saldo Anterior";

            int linha = 2;

            foreach (var conta in contasRemovidas)
            {
                planilha.Cells[linha, 1] = conta.Codigo;
                planilha.Cells[linha, 2] = conta.Descricao;
                planilha.Cells[linha, 3] = conta.SaldoAtual;

                linha++;
            }

            planilha.Columns.AutoFit();
        }
    }
}