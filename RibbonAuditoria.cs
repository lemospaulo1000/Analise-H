using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AnaliseH3.ExcelLayer;

namespace AnaliseH3
{
    public partial class RibbonAuditoria
    {
        private void RibbonAuditoria_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnExecutarAnalise_Click(object sender, RibbonControlEventArgs e)
        {
            var executor = new ExecutorAnalise();

            executor.Executar(
                Globals.ThisAddIn.Application
            );
        }
        private void btnDestacarSubgrupos_Click(object sender, RibbonControlEventArgs e)
        {
            var comando = new ComandoDestacarSubgrupos();

            comando.Executar(Globals.ThisAddIn.Application);
        }

        private void btnDestacarSubgrupos_Click_1(object sender, RibbonControlEventArgs e)
        {
            var comando = new AnaliseH3.ExcelLayer.ComandoDestacarSubgrupos();

            comando.Executar(Globals.ThisAddIn.Application);
        }

        private void btnLimpaIntra_Click(object sender, RibbonControlEventArgs e)
        {
            var comando = new AnaliseH3.ExcelLayer.ComandoLimpaIntra();

            comando.Executar(Globals.ThisAddIn.Application);
        }

    }
}
