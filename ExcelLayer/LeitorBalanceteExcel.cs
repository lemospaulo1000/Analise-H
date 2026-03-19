using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using SIGEFES.LimpaIntra.Modelos;

namespace SIGEFES.LimpaIntra.LeituraExcel
{
    public class LeitorBalanceteExcel
    {
        public Dictionary<long, ContaContabil> Ler(string caminho)
        {
            IWorkbook workbook = AbrirWorkbook(caminho);

            var sheet = workbook.GetSheetAt(0);

            var contas = new Dictionary<long, ContaContabil>();

            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);

                if (row == null)
                    continue;

                var contaCell = row.GetCell(0);

                if (contaCell == null)
                    continue;

                var contaTexto = contaCell.ToString().Trim();

                if (string.IsNullOrWhiteSpace(contaTexto))
                    continue;

                var partes = contaTexto.Split(new string[] { " - " }, StringSplitOptions.None);

                if (partes.Length == 0)
                    continue;

                if (!long.TryParse(partes[0], out long codigo))
                    continue;

                string descricao = partes.Length > 1 ? partes[1].Trim() : "";

                double saldoInicial = ObterValor(row.GetCell(2));
                double debito = ObterValor(row.GetCell(3));
                double credito = ObterValor(row.GetCell(4));
                double saldoAtual = ObterValor(row.GetCell(5));

                string dc = row.GetCell(6)?.ToString()?.Trim();

                var conta = new ContaContabil
                {
                    Codigo = codigo,
                    Descricao = descricao,
                    SaldoInicial = saldoInicial,
                    Debito = debito,
                    Credito = credito,
                    SaldoAtual = saldoAtual,
                    DC = dc
                };

                contas[codigo] = conta;
            }

            return contas;
        }

        private IWorkbook AbrirWorkbook(string caminho)
        {
            using (var stream = new FileStream(caminho, FileMode.Open, FileAccess.Read))
            {
                if (Path.GetExtension(caminho).ToLower() == ".xlsx")
                {
                    return new XSSFWorkbook(stream); // Excel moderno
                }
                else
                {
                    return new HSSFWorkbook(stream); // Excel antigo
                }
            }
        }

        private double ObterValor(ICell cell)
        {
            if (cell == null)
                return 0;

            if (cell.CellType == CellType.Numeric)
                return cell.NumericCellValue;

            if (double.TryParse(cell.ToString(), out double valor))
                return valor;

            return 0;
        }
    }
}