using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using SIGEFES.LimpaIntra.Modelos;

namespace AnaliseH3.ExcelLayer
{
    public class RelatorioLimpaIntraBuilder
    {
        public void Gerar(
            Excel.Application app,
            Dictionary<long, ContaContabil> original,
            Dictionary<long, ContaContabil> processado,
            List<ContaContabil> intra)
        {
            Excel.Workbook wb = app.Workbooks.Add();

            var wsOriginal = wb.Worksheets[1];
            wsOriginal.Name = "Balancete_Original";

            var wsSemIntra = wb.Worksheets.Add();
            wsSemIntra.Name = "Balancete_Sem_Intra";

            var wsDiff = wb.Worksheets.Add();
            wsDiff.Name = "Diferenca";

            var wsIntra = wb.Worksheets.Add();
            wsIntra.Name = "Intra_OFSS";

            var wsResumo = wb.Worksheets.Add();
            wsResumo.Name = "Resumo_Auditoria";

            var wsResumoGrupo = wb.Worksheets.Add();
            wsResumoGrupo.Name = "Resumo_Grupos";

            EscreverCabecalho(wsOriginal);
            EscreverCabecalho(wsSemIntra);
            EscreverCabecalho(wsDiff);
            EscreverCabecalho(wsIntra);

            int l1 = 2, l2 = 2, l3 = 2, l4 = 2;

            foreach (var c in original.Values.OrderBy(x => x.Codigo))
                EscreverConta(wsOriginal, l1++, c);

            foreach (var c in processado.Values.OrderBy(x => x.Codigo))
                EscreverConta(wsSemIntra, l2++, c);

            foreach (var c in original.Values.OrderBy(x => x.Codigo))
            {
                if (processado.ContainsKey(c.Codigo))
                {
                    var novo = processado[c.Codigo];

                    var diff = new ContaContabil
                    {
                        Codigo = c.Codigo,
                        Descricao = c.Descricao,
                        SaldoInicial = novo.SaldoInicial - c.SaldoInicial,
                        Debito = novo.Debito - c.Debito,
                        Credito = novo.Credito - c.Credito,
                        SaldoAtual = novo.SaldoAtual - c.SaldoAtual,
                        DC = c.DC
                    };

                    EscreverConta(wsDiff, l3++, diff);
                }
            }

            foreach (var c in intra.OrderBy(x => x.Codigo))
                EscreverConta(wsIntra, l4++, c);

            Formatar(wsOriginal);
            Formatar(wsSemIntra);
            Formatar(wsDiff);
            Formatar(wsIntra);

            wsIntra.UsedRange.Interior.Color = 0xFFFF99;

            GerarResumo(wsResumo, original, processado, intra);
            GerarResumoGrupos(wsResumoGrupo, original, intra);
        }

        // ================= RESUMO AUDITORIA =================

        private void GerarResumo(
    Excel.Worksheet ws,
    Dictionary<long, ContaContabil> original,
    Dictionary<long, ContaContabil> processado,
    List<ContaContabil> intra)
        {
            decimal somaIntraAtivo = intra.Where(x => (x.Codigo / 100000000) == 1).Sum(x => x.SaldoAtual);
            decimal somaIntraPassivo = intra.Where(x => (x.Codigo / 100000000) == 2).Sum(x => x.SaldoAtual);
            decimal somaIntraClasse3 = intra.Where(x => (x.Codigo / 100000000) == 3).Sum(x => x.SaldoAtual);
            decimal somaIntraClasse4 = intra.Where(x => (x.Codigo / 100000000) == 4).Sum(x => x.SaldoAtual);

            decimal ativoOriginal = original.ContainsKey(100000000) ? original[100000000].SaldoAtual : 0;
            decimal passivoOriginal = original.ContainsKey(200000000) ? original[200000000].SaldoAtual : 0;
            decimal classe3Original = original.ContainsKey(300000000) ? original[300000000].SaldoAtual : 0;
            decimal classe4Original = original.ContainsKey(400000000) ? original[400000000].SaldoAtual : 0;

            decimal ativoFinal = processado.ContainsKey(100000000) ? processado[100000000].SaldoAtual : 0;
            decimal passivoFinal = processado.ContainsKey(200000000) ? processado[200000000].SaldoAtual : 0;

            decimal classe3Final = classe3Original - somaIntraClasse3;
            decimal classe4Final = classe4Original - somaIntraClasse4;

            ws.Cells[1, 1] = "Resumo contas INTRA-OFSS";
            ws.Range["A1:D1"].Font.Bold = true;

            ws.Cells[3, 2] = "Original";
            ws.Cells[3, 3] = "INTRA";
            ws.Cells[3, 4] = "Após ajuste";
            ws.Range["B3:D3"].Font.Bold = true;

            // ATIVO
            ws.Cells[4, 1] = "ATIVO";
            ws.Cells[4, 2] = Math.Abs(ativoOriginal);
            ws.Cells[4, 3] = Math.Abs(somaIntraAtivo);
            ws.Cells[4, 4] = Math.Abs(ativoFinal);

            // PASSIVO
            ws.Cells[5, 1] = "Passivo e Patrimônio Líquido";
            ws.Cells[5, 2] = Math.Abs(passivoOriginal);
            ws.Cells[5, 3] = Math.Abs(somaIntraPassivo);
            ws.Cells[5, 4] = Math.Abs(passivoFinal);

            // DIFERENÇA ATIVO vs PASSIVO
            ws.Cells[6, 1] = "Diferença";
            ws.Cells[6, 3] = Math.Abs(somaIntraAtivo) - Math.Abs(somaIntraPassivo);

            ws.Cells[6, 1].Font.Color = 255;
            ws.Cells[6, 3].Font.Color = 255;

            // CLASSE 3
            ws.Cells[8, 1] = "Variação Patrimonial Diminutiva";
            ws.Cells[8, 2] = Math.Abs(classe3Original);
            ws.Cells[8, 3] = Math.Abs(somaIntraClasse3);
            ws.Cells[8, 4] = Math.Abs(classe3Final);

            // CLASSE 4
            ws.Cells[9, 1] = "Variação Patrimonial Aumentativa";
            ws.Cells[9, 2] = Math.Abs(classe4Original);
            ws.Cells[9, 3] = Math.Abs(somaIntraClasse4);
            ws.Cells[9, 4] = Math.Abs(classe4Final);

            // DIFERENÇA CLASSES
            ws.Cells[10, 1] = "Diferença";
            ws.Cells[10, 3] = Math.Abs(somaIntraClasse3) - Math.Abs(somaIntraClasse4);

            ws.Cells[10, 1].Font.Color = 255;
            ws.Cells[10, 3].Font.Color = 255;

            // FORMATAÇÃO CONTÁBIL
            ws.Range["B4:D10"].NumberFormat =
                "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

            // BORDAS
            Excel.Range tabela = ws.Range["A3:D10"];
            tabela.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            ws.Columns.AutoFit();
        }

        // ================= RESUMO GRUPOS =================

        private void GerarResumoGrupos(
    Excel.Worksheet ws,
    Dictionary<long, ContaContabil> original,
    List<ContaContabil> intra)
        {
            ws.Cells[1, 1] = "Grupo";
            ws.Cells[1, 2] = "Descrição";
            ws.Cells[1, 3] = "Saldo Original";
            ws.Cells[1, 4] = "Intra";
            ws.Cells[1, 5] = "Saldo Consolidado";

            ws.Range["A1:E1"].Font.Bold = true;

            var originalPorGrupo = new Dictionary<long, decimal>();
            var intraPorGrupo = new Dictionary<long, decimal>();

            foreach (var c in original.Values)
            {
                if (c.Codigo % 10000000 != 0) continue;
                if (c.Codigo % 100000000 == 0) continue;

                int classe = (int)(c.Codigo / 100000000);
                if (classe > 4) continue;

                long grupo = c.Codigo / 10000000;
                originalPorGrupo[grupo] = c.SaldoAtual;
            }

            foreach (var c in intra)
            {
                int classe = (int)(c.Codigo / 100000000);
                if (classe > 4) continue;

                long grupo = c.Codigo / 10000000;

                if (!intraPorGrupo.ContainsKey(grupo))
                    intraPorGrupo[grupo] = 0;

                intraPorGrupo[grupo] += c.SaldoAtual;
            }

            var grupos = originalPorGrupo.Keys
                .Union(intraPorGrupo.Keys)
                .OrderBy(x => x);

            int linha = 2;

            foreach (var g in grupos)
            {
                decimal originalV = originalPorGrupo.ContainsKey(g) ? originalPorGrupo[g] : 0;
                decimal intraV = intraPorGrupo.ContainsKey(g) ? intraPorGrupo[g] : 0;

                long codigoGrupo = g * 10000000;

                string descricao = original.ContainsKey(codigoGrupo)
                    ? original[codigoGrupo].Descricao
                    : "";

                ws.Cells[linha, 1] = g;
                ws.Cells[linha, 2] = descricao;
                ws.Cells[linha, 3] = originalV;
                ws.Cells[linha, 4] = intraV;
                ws.Cells[linha, 5] = originalV - intraV;

                linha++;
            }

            int ultimaLinha = linha - 1;

            // =============================
            // FORMATAÇÃO CONTÁBIL
            // =============================

            ws.Range["C2:E" + ultimaLinha].NumberFormat =
                "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

            // =============================
            // BORDAS
            // =============================

            Excel.Range tabela = ws.Range["A1:E" + ultimaLinha];
            tabela.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // =============================
            // ZEBRADO (OPCIONAL MAS RECOMENDADO)
            // =============================

            for (int i = 2; i <= ultimaLinha; i++)
            {
                if (i % 2 == 0)
                    ws.Range["A" + i, "E" + i].Interior.Color = 0xF2F2F2;
            }

            ws.Columns.AutoFit();
        }

        // ================= AUX =================

        private void EscreverCabecalho(Excel.Worksheet ws)
        {
            ws.Cells[1, 1] = "Conta";
            ws.Cells[1, 2] = "Descrição";
            ws.Cells[1, 3] = "Saldo Inicial";
            ws.Cells[1, 4] = "Débito";
            ws.Cells[1, 5] = "Crédito";
            ws.Cells[1, 6] = "Saldo Atual";
            ws.Cells[1, 7] = "D/C";

            ws.Range["A1:G1"].Font.Bold = true;
        }

        private void EscreverConta(Excel.Worksheet ws, int linha, ContaContabil c)
        {
            ws.Cells[linha, 1] = c.Codigo;
            ws.Cells[linha, 2] = c.Descricao;
            ws.Cells[linha, 3] = c.SaldoInicial;
            ws.Cells[linha, 4] = c.Debito;
            ws.Cells[linha, 5] = c.Credito;
            ws.Cells[linha, 6] = c.SaldoAtual;
            ws.Cells[linha, 7] = c.DC.ToString();
        }

        private void Formatar(Excel.Worksheet ws)
        {
            Excel.Range used = ws.UsedRange;
            int linhas = used.Rows.Count;

            Excel.Range numeros = ws.Range["C2:F" + linhas];
            numeros.NumberFormat =
                "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

            used.Columns.AutoFit();

            ws.Application.ActiveWindow.SplitRow = 1;
            ws.Application.ActiveWindow.FreezePanes = true;

            for (int i = 2; i <= linhas; i++)
            {
                if (i % 2 == 0)
                    ws.Range["A" + i, "G" + i].Interior.Color = 0xF2F2F2;

                object val = ws.Cells[i, 1].Value;

                if (val == null) continue;

                long codigo;
                if (!long.TryParse(val.ToString(), out codigo))
                    continue;

                int nivel = NivelConta(codigo);

                if (nivel == 1 || nivel == 2)
                    ws.Range["A" + i, "G" + i].Font.Bold = true;
            }
        }

        private int NivelConta(long codigo)
        {
            if (codigo % 100000000 == 0) return 1;
            if (codigo % 10000000 == 0) return 2;
            if (codigo % 1000000 == 0) return 3;
            if (codigo % 100000 == 0) return 4;
            return 5;
        }
    }
}