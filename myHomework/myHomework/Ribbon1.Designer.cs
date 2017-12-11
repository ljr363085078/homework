namespace myHomework
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnEncrypt = this.Factory.CreateRibbonButton();
            this.btnGetGontent = this.Factory.CreateRibbonButton();
            this.getTimeDiff = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnEncrypt);
            this.group1.Items.Add(this.btnGetGontent);
            this.group1.Items.Add(this.getTimeDiff);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnEncrypt
            // 
            this.btnEncrypt.Label = "加密";
            this.btnEncrypt.Name = "btnEncrypt";
            this.btnEncrypt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEncrypt_Click);
            // 
            // btnGetGontent
            // 
            this.btnGetGontent.Label = "网抓";
            this.btnGetGontent.Name = "btnGetGontent";
            this.btnGetGontent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetGontent_Click);
            // 
            // getTimeDiff
            // 
            this.getTimeDiff.Label = "获取时间";
            this.getTimeDiff.Name = "getTimeDiff";
            this.getTimeDiff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.getTimeDiff_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEncrypt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetGontent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton getTimeDiff;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
