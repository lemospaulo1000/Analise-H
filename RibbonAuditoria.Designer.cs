namespace AnaliseH3
{
    partial class RibbonAuditoria : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonAuditoria()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.TabAuditoria = this.Factory.CreateRibbonTab();
            this.groupAnalise = this.Factory.CreateRibbonGroup();
            this.btnExecutarAnalise = this.Factory.CreateRibbonButton();
            this.btnDestacarSubgrupos = this.Factory.CreateRibbonButton();
            this.TabAuditoria.SuspendLayout();
            this.groupAnalise.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabAuditoria
            // 
            this.TabAuditoria.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabAuditoria.Groups.Add(this.groupAnalise);
            this.TabAuditoria.Label = "Auditoria";
            this.TabAuditoria.Name = "TabAuditoria";
            // 
            // groupAnalise
            // 
            this.groupAnalise.Items.Add(this.btnExecutarAnalise);
            this.groupAnalise.Items.Add(this.btnDestacarSubgrupos);
            this.groupAnalise.Label = "Análise-H";
            this.groupAnalise.Name = "groupAnalise";
            // 
            // btnExecutarAnalise
            // 
            this.btnExecutarAnalise.Label = "Executar Análise-H";
            this.btnExecutarAnalise.Name = "btnExecutarAnalise";
            this.btnExecutarAnalise.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecutarAnalise_Click);
            // 
            // btnDestacarSubgrupos
            // 
            this.btnDestacarSubgrupos.Label = "Destacar Subgrupos";
            this.btnDestacarSubgrupos.Name = "btnDestacarSubgrupos";
            this.btnDestacarSubgrupos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDestacarSubgrupos_Click_1);
            // 
            // RibbonAuditoria
            // 
            this.Name = "RibbonAuditoria";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabAuditoria);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonAuditoria_Load);
            this.TabAuditoria.ResumeLayout(false);
            this.TabAuditoria.PerformLayout();
            this.groupAnalise.ResumeLayout(false);
            this.groupAnalise.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabAuditoria;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnalise;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecutarAnalise;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDestacarSubgrupos;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonAuditoria RibbonAuditoria
        {
            get { return this.GetRibbon<RibbonAuditoria>(); }
        }
    }
}
