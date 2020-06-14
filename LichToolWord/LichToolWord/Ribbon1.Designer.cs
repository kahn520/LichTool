namespace LichToolWord
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
            this.btnReplaceTitle = this.Factory.CreateRibbonButton();
            this.btnSaveAsDocx = this.Factory.CreateRibbonButton();
            this.btnFixNum = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnReplaceTitle);
            this.group1.Items.Add(this.btnSaveAsDocx);
            this.group1.Items.Add(this.btnFixNum);
            this.group1.Label = "lich";
            this.group1.Name = "group1";
            // 
            // btnReplaceTitle
            // 
            this.btnReplaceTitle.Label = "替换标题";
            this.btnReplaceTitle.Name = "btnReplaceTitle";
            this.btnReplaceTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceTitle_Click);
            // 
            // btnSaveAsDocx
            // 
            this.btnSaveAsDocx.Label = "保存为Docx";
            this.btnSaveAsDocx.Name = "btnSaveAsDocx";
            this.btnSaveAsDocx.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAsDocx_Click);
            // 
            // btnFixNum
            // 
            this.btnFixNum.Label = "修复试卷编号";
            this.btnFixNum.Name = "btnFixNum";
            this.btnFixNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFixNum_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAsDocx;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFixNum;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
