using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
