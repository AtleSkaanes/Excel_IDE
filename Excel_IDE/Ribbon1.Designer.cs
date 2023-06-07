namespace Excel_IDE
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.runButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.saveBtn = this.Factory.CreateRibbonButton();
            this.openBtn = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.importBtn = this.Factory.CreateRibbonButton();
            this.PythonIntBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "ExceIDE";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.runButton);
            this.group1.Label = "Run";
            this.group1.Name = "group1";
            // 
            // runButton
            // 
            this.runButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runButton.Description = "Run your fantastic code";
            this.runButton.Enabled = false;
            this.runButton.Image = global::Excel_IDE.Properties.Resources.runButton;
            this.runButton.Label = "run code";
            this.runButton.Name = "runButton";
            this.runButton.ScreenTip = "Run your code";
            this.runButton.ShowImage = true;
            this.runButton.SuperTip = "But don\'t do it if you have bugs :)";
            this.runButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.saveBtn);
            this.group2.Items.Add(this.openBtn);
            this.group2.Label = "Files";
            this.group2.Name = "group2";
            // 
            // saveBtn
            // 
            this.saveBtn.Label = "Save files";
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveBtn_Click);
            // 
            // openBtn
            // 
            this.openBtn.Label = "Open directory";
            this.openBtn.Name = "openBtn";
            this.openBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openBtn_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.importBtn);
            this.group3.Items.Add(this.PythonIntBtn);
            this.group3.Label = "Python";
            this.group3.Name = "group3";
            // 
            // importBtn
            // 
            this.importBtn.Enabled = false;
            this.importBtn.Label = "Add Package";
            this.importBtn.Name = "importBtn";
            this.importBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importBtn_Click);
            // 
            // PythonIntBtn
            // 
            this.PythonIntBtn.Label = "Add python intepreter";
            this.PythonIntBtn.Name = "PythonIntBtn";
            this.PythonIntBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PythonIntBtn_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PythonIntBtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
