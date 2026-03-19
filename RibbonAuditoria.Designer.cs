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
            this.Grp_limparIntra = this.Factory.CreateRibbonGroup();
            this.btnLimpaIntra = this.Factory.CreateRibbonButton();
            this.TabAuditoria.SuspendLayout();
            this.groupAnalise.SuspendLayout();
            this.Grp_limparIntra.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabAuditoria
            // 
            this.TabAuditoria.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabAuditoria.Groups.Add(this.groupAnalise);
            this.TabAuditoria.Groups.Add(this.Grp_limparIntra);
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
            this.btnExecutarAnalise.Label = "Análise-H";
            this.btnExecutarAnalise.Name = "btnExecutarAnalise";
            this.btnExecutarAnalise.OfficeImageId = "CharacterSpacingGallery";
            this.btnExecutarAnalise.ScreenTip = "Análise Horizontal";
            this.btnExecutarAnalise.ShowImage = true;
            this.btnExecutarAnalise.SuperTip = "Realiza análise horizontal a partir do balancete anterior e do exercício atual";
            this.btnExecutarAnalise.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecutarAnalise_Click);
            // 
            // btnDestacarSubgrupos
            // 
            this.btnDestacarSubgrupos.Label = "Destaca grupos";
            this.btnDestacarSubgrupos.Name = "btnDestacarSubgrupos";
            this.btnDestacarSubgrupos.OfficeImageId = "DrillInto";
            this.btnDestacarSubgrupos.ScreenTip = "Destaca Subgrupos materiais";
            this.btnDestacarSubgrupos.ShowImage = true;
            this.btnDestacarSubgrupos.SuperTip = "Destaca subgrupos materiais do balancete com base na aba Materialidade";
            this.btnDestacarSubgrupos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDestacarSubgrupos_Click_1);
            // 
            // Grp_limparIntra
            // 
            this.Grp_limparIntra.Items.Add(this.btnLimpaIntra);
            this.Grp_limparIntra.Label = "Limpar_Intra";
            this.Grp_limparIntra.Name = "Grp_limparIntra";
            // 
            // btnLimpaIntra
            // 
            this.btnLimpaIntra.Label = "Intra OFSS";
            this.btnLimpaIntra.Name = "btnLimpaIntra";
            this.btnLimpaIntra.OfficeImageId = "AccessRecycleBin";
            this.btnLimpaIntra.ScreenTip = "Excluir contas Intra OFSS";
            this.btnLimpaIntra.ShowImage = true;
            this.btnLimpaIntra.SuperTip = "Forneça balancete do SIGEFES que será excluídas Intra OFSS e apresentado resumos";
            this.btnLimpaIntra.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLimpaIntra_Click);
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
            this.Grp_limparIntra.ResumeLayout(false);
            this.Grp_limparIntra.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabAuditoria;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnalise;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecutarAnalise;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDestacarSubgrupos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Grp_limparIntra;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLimpaIntra;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonAuditoria RibbonAuditoria
        {
            get { return this.GetRibbon<RibbonAuditoria>(); }
        }
    }
}
