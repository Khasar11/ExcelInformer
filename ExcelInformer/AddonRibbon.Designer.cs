namespace ExcelInformer
{
    partial class AddonRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddonRibbon()
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
            this.eiTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.updOnceBtn = this.Factory.CreateRibbonButton();
            this.startBtn = this.Factory.CreateRibbonButton();
            this.stopBtn = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.indicator = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.scannedCells = this.Factory.CreateRibbonLabel();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.scanStart = this.Factory.CreateRibbonLabel();
            this.scanEnd = this.Factory.CreateRibbonLabel();
            this.scanTime = this.Factory.CreateRibbonLabel();
            this.settingsGroup = this.Factory.CreateRibbonGroup();
            this.safeSwitch = this.Factory.CreateRibbonToggleButton();
            this.box5 = this.Factory.CreateRibbonBox();
            this.box4 = this.Factory.CreateRibbonBox();
            this.box3 = this.Factory.CreateRibbonBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.startRowB = this.Factory.CreateRibbonEditBox();
            this.endRowB = this.Factory.CreateRibbonEditBox();
            this.startColumnB = this.Factory.CreateRibbonEditBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.box2 = this.Factory.CreateRibbonBox();
            this.offsetValue = this.Factory.CreateRibbonEditBox();
            this.offsetTime = this.Factory.CreateRibbonEditBox();
            this.helpButton = this.Factory.CreateRibbonButton();
            this.eiTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.settingsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // eiTab
            // 
            this.eiTab.Groups.Add(this.group1);
            this.eiTab.Groups.Add(this.group2);
            this.eiTab.Groups.Add(this.group3);
            this.eiTab.Groups.Add(this.settingsGroup);
            this.eiTab.Label = "Excel Informer";
            this.eiTab.Name = "eiTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.updOnceBtn);
            this.group1.Items.Add(this.startBtn);
            this.group1.Items.Add(this.stopBtn);
            this.group1.Label = "Update";
            this.group1.Name = "group1";
            // 
            // updOnceBtn
            // 
            this.updOnceBtn.Label = "Update once";
            this.updOnceBtn.Name = "updOnceBtn";
            this.updOnceBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updOnceBtn_Click);
            // 
            // startBtn
            // 
            this.startBtn.Label = "Start";
            this.startBtn.Name = "startBtn";
            this.startBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startBtn_Click);
            // 
            // stopBtn
            // 
            this.stopBtn.Label = "Stop";
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.stopBtn_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.indicator);
            this.group2.Items.Add(this.label1);
            this.group2.Items.Add(this.scannedCells);
            this.group2.Label = "Indication";
            this.group2.Name = "group2";
            // 
            // indicator
            // 
            this.indicator.Image = global::ExcelInformer.Properties.Resources.indication_red;
            this.indicator.Label = "Not running";
            this.indicator.Name = "indicator";
            this.indicator.ShowImage = true;
            // 
            // label1
            // 
            this.label1.Label = "This scan upd";
            this.label1.Name = "label1";
            this.label1.ScreenTip = "Updated cells this scan";
            // 
            // scannedCells
            // 
            this.scannedCells.Label = "Scanned";
            this.scannedCells.Name = "scannedCells";
            // 
            // group3
            // 
            this.group3.Items.Add(this.scanStart);
            this.group3.Items.Add(this.scanEnd);
            this.group3.Items.Add(this.scanTime);
            this.group3.Label = "Timestamps";
            this.group3.Name = "group3";
            // 
            // scanStart
            // 
            this.scanStart.Label = "Scan Start";
            this.scanStart.Name = "scanStart";
            this.scanStart.ScreenTip = "Last scan start timestamp";
            // 
            // scanEnd
            // 
            this.scanEnd.Label = "Scan End";
            this.scanEnd.Name = "scanEnd";
            this.scanEnd.ScreenTip = "Last scan end timestamp";
            // 
            // scanTime
            // 
            this.scanTime.Label = "Scan Time";
            this.scanTime.Name = "scanTime";
            this.scanTime.ScreenTip = "Scan time total";
            // 
            // settingsGroup
            // 
            this.settingsGroup.Items.Add(this.safeSwitch);
            this.settingsGroup.Items.Add(this.box5);
            this.settingsGroup.Items.Add(this.box4);
            this.settingsGroup.Items.Add(this.box3);
            this.settingsGroup.Items.Add(this.separator1);
            this.settingsGroup.Items.Add(this.box1);
            this.settingsGroup.Items.Add(this.startRowB);
            this.settingsGroup.Items.Add(this.endRowB);
            this.settingsGroup.Items.Add(this.startColumnB);
            this.settingsGroup.Items.Add(this.separator2);
            this.settingsGroup.Items.Add(this.box2);
            this.settingsGroup.Items.Add(this.offsetValue);
            this.settingsGroup.Items.Add(this.offsetTime);
            this.settingsGroup.Items.Add(this.helpButton);
            this.settingsGroup.Label = "Settings";
            this.settingsGroup.Name = "settingsGroup";
            // 
            // safeSwitch
            // 
            this.safeSwitch.Checked = true;
            this.safeSwitch.Label = "Safe mode";
            this.safeSwitch.Name = "safeSwitch";
            this.safeSwitch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.safeSwitch_Click);
            // 
            // box5
            // 
            this.box5.Name = "box5";
            // 
            // box4
            // 
            this.box4.Name = "box4";
            // 
            // box3
            // 
            this.box3.Name = "box3";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // box1
            // 
            this.box1.Name = "box1";
            // 
            // startRowB
            // 
            this.startRowB.Label = "Start Row";
            this.startRowB.MaxLength = 6;
            this.startRowB.Name = "startRowB";
            this.startRowB.ScreenTip = "Number";
            this.startRowB.Text = "1";
            this.startRowB.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startRowB_TextChanged);
            // 
            // endRowB
            // 
            this.endRowB.Label = "End Row";
            this.endRowB.MaxLength = 6;
            this.endRowB.Name = "endRowB";
            this.endRowB.ScreenTip = "Number";
            this.endRowB.Text = "2000";
            this.endRowB.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endRowB_TextChanged);
            // 
            // startColumnB
            // 
            this.startColumnB.Label = "Column";
            this.startColumnB.MaxLength = 6;
            this.startColumnB.Name = "startColumnB";
            this.startColumnB.ScreenTip = "Number";
            this.startColumnB.Text = "1";
            this.startColumnB.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startColumnB_TextChanged);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // box2
            // 
            this.box2.Name = "box2";
            // 
            // offsetValue
            // 
            this.offsetValue.Label = "Offset Value";
            this.offsetValue.MaxLength = 6;
            this.offsetValue.Name = "offsetValue";
            this.offsetValue.ScreenTip = "Offset for value in columns";
            this.offsetValue.Text = "1";
            this.offsetValue.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.offsetValue_TextChanged);
            // 
            // offsetTime
            // 
            this.offsetTime.Label = "Offset Time";
            this.offsetTime.MaxLength = 6;
            this.offsetTime.Name = "offsetTime";
            this.offsetTime.ScreenTip = "Offset for time in columns";
            this.offsetTime.Text = "2";
            this.offsetTime.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.offsetTime_TextChanged);
            // 
            // helpButton
            // 
            this.helpButton.Label = "Help";
            this.helpButton.Name = "helpButton";
            this.helpButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.helpButton_Click);
            // 
            // AddonRibbon
            // 
            this.Name = "AddonRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.eiTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.EiRibbon_Load);
            this.eiTab.ResumeLayout(false);
            this.eiTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.settingsGroup.ResumeLayout(false);
            this.settingsGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab eiTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton startBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stopBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton indicator;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel scanStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel scanEnd;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel scanTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updOnceBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup settingsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton safeSwitch;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox startRowB;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox endRowB;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox startColumnB;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox offsetValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox offsetTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel scannedCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton helpButton;
    }

    partial class ThisRibbonCollection
    {
        internal AddonRibbon TineRibbon
        {
            get { return this.GetRibbon<AddonRibbon>(); }
        }
    }
}
