using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using AnaliseH3.ExcelLayer;
using AnaliseH3.Core.Services;

namespace AnaliseH3
{
    public class ExecutorAnalise
    {
        public void Executar(Excel.Application excelApp)
        {
            try
            {
                var leitor = new LeitorBalanceteExcel();

                // selecionar arquivos
                var dialog = new OpenFileDialog();
                dialog.Title = "Selecione o Balancete Anterior";
                dialog.Filter = "Arquivos Excel (*.xls;*.xlsx)|*.xls;*.xlsx";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                var caminhoAnterior = dialog.FileName;

                dialog.Title = "Selecione o Balancete Atual";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                var caminhoAtual = dialog.FileName;

                if (caminhoAnterior == caminhoAtual)
                {
                    MessageBox.Show(
                        "O balancete anterior e o atual são o mesmo arquivo.\n\nSelecione arquivos diferentes.",
                        "Análise-H 3.0"
                    );
                    return;
                }

                // ler balancetes
                var balanceteAnterior = leitor.Ler(excelApp, caminhoAnterior, "Anterior");
                var balanceteAtual = leitor.Ler(excelApp, caminhoAtual, "Atual");

                // criar análise
                var analise = new AnaliseComparativa();

                analise.AdicionarBalancete(balanceteAnterior);
                analise.AdicionarBalancete(balanceteAtual);

                analise.DefinirPeriodoAtual("Atual");

                // pegar saldo do ativo
                var contaAtivo = balanceteAtual.ObterConta("100000000");

                double saldoAtivo = 0;

                if (contaAtivo != null)
                    saldoAtivo = Math.Abs(contaAtivo.SaldoAtual);

                // gerar planilha Materialidade
                var geradorMaterialidade = new GeradorMaterialidade();
                geradorMaterialidade.Gerar(excelApp, saldoAtivo);

                // definir limite de materialidade
                double materialidade = saldoAtivo * 0.01;

                analise.DefinirLimiteMaterialidade(materialidade);

                // executar comparação
                var resultado = analise.ExecutarComparacao();

                var gerador = new GeradorPlanilhaComparativo();
                gerador.Gerar(excelApp, resultado);
                
                MessageBox.Show(
                    $"Comparação executada.\n\n" +
                    $"Contas analisadas: {resultado.Count}",
                    "Análise-H 3.0"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Erro");
            }
        }
    }
}