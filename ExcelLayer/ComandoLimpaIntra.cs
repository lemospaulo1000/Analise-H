using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AnaliseH3.ExcelLayer
{
    public class ComandoLimpaIntra
    {
        public void Executar(Excel.Application app)
        {
            try
            {
                var dialog = new OpenFileDialog();
                dialog.Filter = "Arquivos Excel (*.xls;*.xlsx)|*.xls;*.xlsx";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                string caminho = dialog.FileName;

                // Executa o motor
                var motor = new SIGEFES.LimpaIntra.Servicos.MotorLimpaIntra();
                var resultado = motor.Processar(caminho);

                // Gera o relatório (novo builder)
                var builder = new RelatorioLimpaIntraBuilder();

                builder.Gerar(
                    app,
                    resultado.Original,
                    resultado.Processado,
                    resultado.Intra
                );

                MessageBox.Show(
                    "Relatório gerado com sucesso.",
                    "Limpa Intra OFSS",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Erro ao executar Limpa Intra:\n" + ex.Message,
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}