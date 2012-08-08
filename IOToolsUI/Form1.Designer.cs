namespace IOToolsUI
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabCompareTermLists = new System.Windows.Forms.TabPage();
            this.lblProgress1 = new System.Windows.Forms.Label();
            this.lblPreviousFilename1 = new System.Windows.Forms.Label();
            this.lblLatestFilename1 = new System.Windows.Forms.Label();
            this.numColStart2 = new System.Windows.Forms.NumericUpDown();
            this.lblCellStart2 = new System.Windows.Forms.Label();
            this.numRowStart2 = new System.Windows.Forms.NumericUpDown();
            this.selPreviousWorksheet = new System.Windows.Forms.ComboBox();
            this.btnLoadPrevTermList = new System.Windows.Forms.Button();
            this.btnCompareLists = new System.Windows.Forms.Button();
            this.numColStart1 = new System.Windows.Forms.NumericUpDown();
            this.lblCellStart1 = new System.Windows.Forms.Label();
            this.numRowStart1 = new System.Windows.Forms.NumericUpDown();
            this.selLatestWorksheet = new System.Windows.Forms.ComboBox();
            this.btnLoadNewTermList = new System.Windows.Forms.Button();
            this.tabCompareIOList = new System.Windows.Forms.TabPage();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtFatMarshallingPanels = new System.Windows.Forms.TextBox();
            this.btnFatAddInstToList = new System.Windows.Forms.Button();
            this.numFatInstColStart = new System.Windows.Forms.NumericUpDown();
            this.lblFatInstStartCell = new System.Windows.Forms.Label();
            this.numFatInstRowStart = new System.Windows.Forms.NumericUpDown();
            this.selFatInstWorksheets = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnFatAddTermToList = new System.Windows.Forms.Button();
            this.numFatTermColStart = new System.Windows.Forms.NumericUpDown();
            this.lblFatTermStartCell = new System.Windows.Forms.Label();
            this.numFatTermRowStart = new System.Windows.Forms.NumericUpDown();
            this.selFatTermWorksheets = new System.Windows.Forms.ComboBox();
            this.btnFatAddIOToList = new System.Windows.Forms.Button();
            this.numFatIOColStart = new System.Windows.Forms.NumericUpDown();
            this.lblFatIOStartCell = new System.Windows.Forms.Label();
            this.numFatIORowStart = new System.Windows.Forms.NumericUpDown();
            this.selFatIOWorksheets = new System.Windows.Forms.ComboBox();
            this.btnFatAddInstrumentList = new System.Windows.Forms.Button();
            this.btnFatAddTermList = new System.Windows.Forms.Button();
            this.btnFatAddIOList = new System.Windows.Forms.Button();
            this.lstInstrumentFilesList = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lstTermFilesList = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lstIOFilesList = new System.Windows.Forms.ListBox();
            this.btnCreateFatSheet = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.tabControl1.SuspendLayout();
            this.tabCompareTermLists.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColStart2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowStart2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColStart1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowStart1)).BeginInit();
            this.tabCompareIOList.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numFatInstColStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatInstRowStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatTermColStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatTermRowStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatIOColStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatIORowStart)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabCompareTermLists);
            this.tabControl1.Controls.Add(this.tabCompareIOList);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(638, 538);
            this.tabControl1.TabIndex = 1;
            // 
            // tabCompareTermLists
            // 
            this.tabCompareTermLists.Controls.Add(this.lblProgress1);
            this.tabCompareTermLists.Controls.Add(this.lblPreviousFilename1);
            this.tabCompareTermLists.Controls.Add(this.lblLatestFilename1);
            this.tabCompareTermLists.Controls.Add(this.numColStart2);
            this.tabCompareTermLists.Controls.Add(this.lblCellStart2);
            this.tabCompareTermLists.Controls.Add(this.numRowStart2);
            this.tabCompareTermLists.Controls.Add(this.selPreviousWorksheet);
            this.tabCompareTermLists.Controls.Add(this.btnLoadPrevTermList);
            this.tabCompareTermLists.Controls.Add(this.btnCompareLists);
            this.tabCompareTermLists.Controls.Add(this.numColStart1);
            this.tabCompareTermLists.Controls.Add(this.lblCellStart1);
            this.tabCompareTermLists.Controls.Add(this.numRowStart1);
            this.tabCompareTermLists.Controls.Add(this.selLatestWorksheet);
            this.tabCompareTermLists.Controls.Add(this.btnLoadNewTermList);
            this.tabCompareTermLists.Location = new System.Drawing.Point(4, 22);
            this.tabCompareTermLists.Name = "tabCompareTermLists";
            this.tabCompareTermLists.Padding = new System.Windows.Forms.Padding(3);
            this.tabCompareTermLists.Size = new System.Drawing.Size(630, 512);
            this.tabCompareTermLists.TabIndex = 0;
            this.tabCompareTermLists.Text = "Compare Termination List";
            this.tabCompareTermLists.UseVisualStyleBackColor = true;
            // 
            // lblProgress1
            // 
            this.lblProgress1.AutoSize = true;
            this.lblProgress1.Location = new System.Drawing.Point(207, 276);
            this.lblProgress1.Name = "lblProgress1";
            this.lblProgress1.Size = new System.Drawing.Size(0, 13);
            this.lblProgress1.TabIndex = 13;
            // 
            // lblPreviousFilename1
            // 
            this.lblPreviousFilename1.AutoSize = true;
            this.lblPreviousFilename1.Location = new System.Drawing.Point(207, 146);
            this.lblPreviousFilename1.Name = "lblPreviousFilename1";
            this.lblPreviousFilename1.Size = new System.Drawing.Size(151, 13);
            this.lblPreviousFilename1.TabIndex = 12;
            this.lblPreviousFilename1.Text = "Previous Termination Filename";
            this.lblPreviousFilename1.Visible = false;
            // 
            // lblLatestFilename1
            // 
            this.lblLatestFilename1.AutoSize = true;
            this.lblLatestFilename1.Location = new System.Drawing.Point(207, 39);
            this.lblLatestFilename1.Name = "lblLatestFilename1";
            this.lblLatestFilename1.Size = new System.Drawing.Size(125, 13);
            this.lblLatestFilename1.TabIndex = 11;
            this.lblLatestFilename1.Text = "Latest Termination Sheet";
            this.lblLatestFilename1.Visible = false;
            // 
            // numColStart2
            // 
            this.numColStart2.Location = new System.Drawing.Point(322, 171);
            this.numColStart2.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numColStart2.Name = "numColStart2";
            this.numColStart2.Size = new System.Drawing.Size(40, 20);
            this.numColStart2.TabIndex = 10;
            this.numColStart2.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numColStart2.Visible = false;
            // 
            // lblCellStart2
            // 
            this.lblCellStart2.AutoSize = true;
            this.lblCellStart2.Location = new System.Drawing.Point(146, 173);
            this.lblCellStart2.Name = "lblCellStart2";
            this.lblCellStart2.Size = new System.Drawing.Size(124, 13);
            this.lblCellStart2.TabIndex = 9;
            this.lblCellStart2.Text = "Data Starts at [Row][Col]";
            this.lblCellStart2.Visible = false;
            // 
            // numRowStart2
            // 
            this.numRowStart2.Location = new System.Drawing.Point(276, 171);
            this.numRowStart2.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numRowStart2.Name = "numRowStart2";
            this.numRowStart2.Size = new System.Drawing.Size(40, 20);
            this.numRowStart2.TabIndex = 8;
            this.numRowStart2.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
            this.numRowStart2.Visible = false;
            // 
            // selPreviousWorksheet
            // 
            this.selPreviousWorksheet.FormattingEnabled = true;
            this.selPreviousWorksheet.Location = new System.Drawing.Point(19, 170);
            this.selPreviousWorksheet.Name = "selPreviousWorksheet";
            this.selPreviousWorksheet.Size = new System.Drawing.Size(121, 21);
            this.selPreviousWorksheet.TabIndex = 7;
            this.selPreviousWorksheet.Text = "Select Worksheet";
            this.selPreviousWorksheet.Visible = false;
            this.selPreviousWorksheet.SelectedIndexChanged += new System.EventHandler(this.selPreviousWorksheet_SelectedIndexChanged);
            // 
            // btnLoadPrevTermList
            // 
            this.btnLoadPrevTermList.Location = new System.Drawing.Point(16, 141);
            this.btnLoadPrevTermList.Name = "btnLoadPrevTermList";
            this.btnLoadPrevTermList.Size = new System.Drawing.Size(185, 23);
            this.btnLoadPrevTermList.TabIndex = 6;
            this.btnLoadPrevTermList.Text = "Open PreviousTermination List";
            this.btnLoadPrevTermList.UseVisualStyleBackColor = true;
            this.btnLoadPrevTermList.Click += new System.EventHandler(this.btnLoadPrevTermList_Click);
            // 
            // btnCompareLists
            // 
            this.btnCompareLists.Location = new System.Drawing.Point(16, 271);
            this.btnCompareLists.Name = "btnCompareLists";
            this.btnCompareLists.Size = new System.Drawing.Size(185, 23);
            this.btnCompareLists.TabIndex = 5;
            this.btnCompareLists.Text = "Compare";
            this.btnCompareLists.UseVisualStyleBackColor = true;
            this.btnCompareLists.Visible = false;
            this.btnCompareLists.Click += new System.EventHandler(this.btnCompareLists_Click);
            // 
            // numColStart1
            // 
            this.numColStart1.Location = new System.Drawing.Point(322, 64);
            this.numColStart1.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numColStart1.Name = "numColStart1";
            this.numColStart1.Size = new System.Drawing.Size(40, 20);
            this.numColStart1.TabIndex = 4;
            this.numColStart1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numColStart1.Visible = false;
            // 
            // lblCellStart1
            // 
            this.lblCellStart1.AutoSize = true;
            this.lblCellStart1.Location = new System.Drawing.Point(146, 66);
            this.lblCellStart1.Name = "lblCellStart1";
            this.lblCellStart1.Size = new System.Drawing.Size(124, 13);
            this.lblCellStart1.TabIndex = 3;
            this.lblCellStart1.Text = "Data Starts at [Row][Col]";
            this.lblCellStart1.Visible = false;
            // 
            // numRowStart1
            // 
            this.numRowStart1.Location = new System.Drawing.Point(276, 64);
            this.numRowStart1.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numRowStart1.Name = "numRowStart1";
            this.numRowStart1.Size = new System.Drawing.Size(40, 20);
            this.numRowStart1.TabIndex = 2;
            this.numRowStart1.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
            this.numRowStart1.Visible = false;
            // 
            // selLatestWorksheet
            // 
            this.selLatestWorksheet.FormattingEnabled = true;
            this.selLatestWorksheet.Location = new System.Drawing.Point(19, 63);
            this.selLatestWorksheet.Name = "selLatestWorksheet";
            this.selLatestWorksheet.Size = new System.Drawing.Size(121, 21);
            this.selLatestWorksheet.TabIndex = 1;
            this.selLatestWorksheet.Text = "Select Worksheet";
            this.selLatestWorksheet.Visible = false;
            this.selLatestWorksheet.SelectedIndexChanged += new System.EventHandler(this.selLatestWorksheet1_SelectedIndexChanged);
            // 
            // btnLoadNewTermList
            // 
            this.btnLoadNewTermList.Location = new System.Drawing.Point(16, 34);
            this.btnLoadNewTermList.Name = "btnLoadNewTermList";
            this.btnLoadNewTermList.Size = new System.Drawing.Size(185, 23);
            this.btnLoadNewTermList.TabIndex = 0;
            this.btnLoadNewTermList.Text = "Open Latest Termination List";
            this.btnLoadNewTermList.UseVisualStyleBackColor = true;
            this.btnLoadNewTermList.Click += new System.EventHandler(this.btnLoadNewTermList_Click);
            // 
            // tabCompareIOList
            // 
            this.tabCompareIOList.Controls.Add(this.label5);
            this.tabCompareIOList.Controls.Add(this.label4);
            this.tabCompareIOList.Controls.Add(this.txtFatMarshallingPanels);
            this.tabCompareIOList.Controls.Add(this.btnFatAddInstToList);
            this.tabCompareIOList.Controls.Add(this.numFatInstColStart);
            this.tabCompareIOList.Controls.Add(this.lblFatInstStartCell);
            this.tabCompareIOList.Controls.Add(this.numFatInstRowStart);
            this.tabCompareIOList.Controls.Add(this.selFatInstWorksheets);
            this.tabCompareIOList.Controls.Add(this.button1);
            this.tabCompareIOList.Controls.Add(this.btnFatAddTermToList);
            this.tabCompareIOList.Controls.Add(this.numFatTermColStart);
            this.tabCompareIOList.Controls.Add(this.lblFatTermStartCell);
            this.tabCompareIOList.Controls.Add(this.numFatTermRowStart);
            this.tabCompareIOList.Controls.Add(this.selFatTermWorksheets);
            this.tabCompareIOList.Controls.Add(this.btnFatAddIOToList);
            this.tabCompareIOList.Controls.Add(this.numFatIOColStart);
            this.tabCompareIOList.Controls.Add(this.lblFatIOStartCell);
            this.tabCompareIOList.Controls.Add(this.numFatIORowStart);
            this.tabCompareIOList.Controls.Add(this.selFatIOWorksheets);
            this.tabCompareIOList.Controls.Add(this.btnFatAddInstrumentList);
            this.tabCompareIOList.Controls.Add(this.btnFatAddTermList);
            this.tabCompareIOList.Controls.Add(this.btnFatAddIOList);
            this.tabCompareIOList.Controls.Add(this.lstInstrumentFilesList);
            this.tabCompareIOList.Controls.Add(this.label3);
            this.tabCompareIOList.Controls.Add(this.lstTermFilesList);
            this.tabCompareIOList.Controls.Add(this.label2);
            this.tabCompareIOList.Controls.Add(this.label1);
            this.tabCompareIOList.Controls.Add(this.lstIOFilesList);
            this.tabCompareIOList.Controls.Add(this.btnCreateFatSheet);
            this.tabCompareIOList.Location = new System.Drawing.Point(4, 22);
            this.tabCompareIOList.Name = "tabCompareIOList";
            this.tabCompareIOList.Padding = new System.Windows.Forms.Padding(3);
            this.tabCompareIOList.Size = new System.Drawing.Size(630, 512);
            this.tabCompareIOList.TabIndex = 1;
            this.tabCompareIOList.Text = "Generate FAT Sheet";
            this.tabCompareIOList.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(336, 414);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(157, 39);
            this.label5.TabIndex = 28;
            this.label5.Text = "One on each line\r\nSame format as in spreadsheets\r\ni.e. 1317-MP-282";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 398);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 13);
            this.label4.TabIndex = 27;
            this.label4.Text = "Marshalling Panels";
            // 
            // txtFatMarshallingPanels
            // 
            this.txtFatMarshallingPanels.Location = new System.Drawing.Point(3, 414);
            this.txtFatMarshallingPanels.Multiline = true;
            this.txtFatMarshallingPanels.Name = "txtFatMarshallingPanels";
            this.txtFatMarshallingPanels.Size = new System.Drawing.Size(327, 92);
            this.txtFatMarshallingPanels.TabIndex = 26;
            // 
            // btnFatAddInstToList
            // 
            this.btnFatAddInstToList.Location = new System.Drawing.Point(336, 370);
            this.btnFatAddInstToList.Name = "btnFatAddInstToList";
            this.btnFatAddInstToList.Size = new System.Drawing.Size(75, 23);
            this.btnFatAddInstToList.TabIndex = 25;
            this.btnFatAddInstToList.Text = "<< Add To List";
            this.btnFatAddInstToList.UseVisualStyleBackColor = true;
            this.btnFatAddInstToList.Visible = false;
            this.btnFatAddInstToList.Click += new System.EventHandler(this.btnFatAddInstToList_Click);
            // 
            // numFatInstColStart
            // 
            this.numFatInstColStart.Location = new System.Drawing.Point(509, 345);
            this.numFatInstColStart.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numFatInstColStart.Name = "numFatInstColStart";
            this.numFatInstColStart.Size = new System.Drawing.Size(40, 20);
            this.numFatInstColStart.TabIndex = 24;
            this.numFatInstColStart.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numFatInstColStart.Visible = false;
            this.numFatInstColStart.ValueChanged += new System.EventHandler(this.numFatInstColStart_ValueChanged);
            // 
            // lblFatInstStartCell
            // 
            this.lblFatInstStartCell.AutoSize = true;
            this.lblFatInstStartCell.Location = new System.Drawing.Point(333, 347);
            this.lblFatInstStartCell.Name = "lblFatInstStartCell";
            this.lblFatInstStartCell.Size = new System.Drawing.Size(124, 13);
            this.lblFatInstStartCell.TabIndex = 23;
            this.lblFatInstStartCell.Text = "Data Starts at [Row][Col]";
            this.lblFatInstStartCell.Visible = false;
            // 
            // numFatInstRowStart
            // 
            this.numFatInstRowStart.Location = new System.Drawing.Point(463, 345);
            this.numFatInstRowStart.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numFatInstRowStart.Name = "numFatInstRowStart";
            this.numFatInstRowStart.Size = new System.Drawing.Size(40, 20);
            this.numFatInstRowStart.TabIndex = 22;
            this.numFatInstRowStart.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
            this.numFatInstRowStart.Visible = false;
            this.numFatInstRowStart.ValueChanged += new System.EventHandler(this.numFatInstRowStart_ValueChanged);
            // 
            // selFatInstWorksheets
            // 
            this.selFatInstWorksheets.FormattingEnabled = true;
            this.selFatInstWorksheets.Location = new System.Drawing.Point(336, 316);
            this.selFatInstWorksheets.Name = "selFatInstWorksheets";
            this.selFatInstWorksheets.Size = new System.Drawing.Size(121, 21);
            this.selFatInstWorksheets.TabIndex = 21;
            this.selFatInstWorksheets.Text = "Select Worksheet";
            this.selFatInstWorksheets.Visible = false;
            this.selFatInstWorksheets.SelectedIndexChanged += new System.EventHandler(this.selFatInstWorksheets_SelectedIndexChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(549, 7);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 20;
            this.button1.Text = "dbg";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnFatAddTermToList
            // 
            this.btnFatAddTermToList.Location = new System.Drawing.Point(336, 239);
            this.btnFatAddTermToList.Name = "btnFatAddTermToList";
            this.btnFatAddTermToList.Size = new System.Drawing.Size(75, 23);
            this.btnFatAddTermToList.TabIndex = 19;
            this.btnFatAddTermToList.Text = "<< Add To List";
            this.btnFatAddTermToList.UseVisualStyleBackColor = true;
            this.btnFatAddTermToList.Visible = false;
            this.btnFatAddTermToList.Click += new System.EventHandler(this.btnFatAddTermToList_Click);
            // 
            // numFatTermColStart
            // 
            this.numFatTermColStart.Location = new System.Drawing.Point(509, 214);
            this.numFatTermColStart.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numFatTermColStart.Name = "numFatTermColStart";
            this.numFatTermColStart.Size = new System.Drawing.Size(40, 20);
            this.numFatTermColStart.TabIndex = 18;
            this.numFatTermColStart.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numFatTermColStart.Visible = false;
            this.numFatTermColStart.ValueChanged += new System.EventHandler(this.numFatTermColStart_ValueChanged);
            // 
            // lblFatTermStartCell
            // 
            this.lblFatTermStartCell.AutoSize = true;
            this.lblFatTermStartCell.Location = new System.Drawing.Point(333, 216);
            this.lblFatTermStartCell.Name = "lblFatTermStartCell";
            this.lblFatTermStartCell.Size = new System.Drawing.Size(124, 13);
            this.lblFatTermStartCell.TabIndex = 17;
            this.lblFatTermStartCell.Text = "Data Starts at [Row][Col]";
            this.lblFatTermStartCell.Visible = false;
            // 
            // numFatTermRowStart
            // 
            this.numFatTermRowStart.Location = new System.Drawing.Point(463, 214);
            this.numFatTermRowStart.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numFatTermRowStart.Name = "numFatTermRowStart";
            this.numFatTermRowStart.Size = new System.Drawing.Size(40, 20);
            this.numFatTermRowStart.TabIndex = 16;
            this.numFatTermRowStart.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
            this.numFatTermRowStart.Visible = false;
            this.numFatTermRowStart.ValueChanged += new System.EventHandler(this.numFatTermRowStart_ValueChanged);
            // 
            // selFatTermWorksheets
            // 
            this.selFatTermWorksheets.FormattingEnabled = true;
            this.selFatTermWorksheets.Location = new System.Drawing.Point(336, 185);
            this.selFatTermWorksheets.Name = "selFatTermWorksheets";
            this.selFatTermWorksheets.Size = new System.Drawing.Size(121, 21);
            this.selFatTermWorksheets.TabIndex = 15;
            this.selFatTermWorksheets.Text = "Select Worksheet";
            this.selFatTermWorksheets.Visible = false;
            this.selFatTermWorksheets.SelectedIndexChanged += new System.EventHandler(this.selFatTermWorksheets_SelectedIndexChanged);
            // 
            // btnFatAddIOToList
            // 
            this.btnFatAddIOToList.Location = new System.Drawing.Point(336, 110);
            this.btnFatAddIOToList.Name = "btnFatAddIOToList";
            this.btnFatAddIOToList.Size = new System.Drawing.Size(75, 23);
            this.btnFatAddIOToList.TabIndex = 14;
            this.btnFatAddIOToList.Text = "<< Add To List";
            this.btnFatAddIOToList.UseVisualStyleBackColor = true;
            this.btnFatAddIOToList.Visible = false;
            this.btnFatAddIOToList.Click += new System.EventHandler(this.btnFatAddIOToList_Click);
            // 
            // numFatIOColStart
            // 
            this.numFatIOColStart.Location = new System.Drawing.Point(509, 85);
            this.numFatIOColStart.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numFatIOColStart.Name = "numFatIOColStart";
            this.numFatIOColStart.Size = new System.Drawing.Size(40, 20);
            this.numFatIOColStart.TabIndex = 13;
            this.numFatIOColStart.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numFatIOColStart.Visible = false;
            this.numFatIOColStart.ValueChanged += new System.EventHandler(this.numFatIOColStart_ValueChanged);
            // 
            // lblFatIOStartCell
            // 
            this.lblFatIOStartCell.AutoSize = true;
            this.lblFatIOStartCell.Location = new System.Drawing.Point(333, 87);
            this.lblFatIOStartCell.Name = "lblFatIOStartCell";
            this.lblFatIOStartCell.Size = new System.Drawing.Size(124, 13);
            this.lblFatIOStartCell.TabIndex = 12;
            this.lblFatIOStartCell.Text = "Data Starts at [Row][Col]";
            this.lblFatIOStartCell.Visible = false;
            // 
            // numFatIORowStart
            // 
            this.numFatIORowStart.Location = new System.Drawing.Point(463, 85);
            this.numFatIORowStart.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numFatIORowStart.Name = "numFatIORowStart";
            this.numFatIORowStart.Size = new System.Drawing.Size(40, 20);
            this.numFatIORowStart.TabIndex = 11;
            this.numFatIORowStart.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            this.numFatIORowStart.Visible = false;
            this.numFatIORowStart.ValueChanged += new System.EventHandler(this.numFatIORowStart_ValueChanged);
            // 
            // selFatIOWorksheets
            // 
            this.selFatIOWorksheets.FormattingEnabled = true;
            this.selFatIOWorksheets.Location = new System.Drawing.Point(336, 56);
            this.selFatIOWorksheets.Name = "selFatIOWorksheets";
            this.selFatIOWorksheets.Size = new System.Drawing.Size(121, 21);
            this.selFatIOWorksheets.TabIndex = 10;
            this.selFatIOWorksheets.Text = "Select Worksheet";
            this.selFatIOWorksheets.Visible = false;
            this.selFatIOWorksheets.SelectedIndexChanged += new System.EventHandler(this.selFatIOWorksheets_SelectedIndexChanged);
            // 
            // btnFatAddInstrumentList
            // 
            this.btnFatAddInstrumentList.Location = new System.Drawing.Point(336, 287);
            this.btnFatAddInstrumentList.Name = "btnFatAddInstrumentList";
            this.btnFatAddInstrumentList.Size = new System.Drawing.Size(124, 23);
            this.btnFatAddInstrumentList.TabIndex = 9;
            this.btnFatAddInstrumentList.Text = "Add Instrument List";
            this.btnFatAddInstrumentList.UseVisualStyleBackColor = true;
            this.btnFatAddInstrumentList.Click += new System.EventHandler(this.btnFatAddInstrumentList_Click);
            // 
            // btnFatAddTermList
            // 
            this.btnFatAddTermList.Location = new System.Drawing.Point(336, 156);
            this.btnFatAddTermList.Name = "btnFatAddTermList";
            this.btnFatAddTermList.Size = new System.Drawing.Size(124, 23);
            this.btnFatAddTermList.TabIndex = 8;
            this.btnFatAddTermList.Text = "Add Termination List";
            this.btnFatAddTermList.UseVisualStyleBackColor = true;
            this.btnFatAddTermList.Click += new System.EventHandler(this.btnAddTermList_Click);
            // 
            // btnFatAddIOList
            // 
            this.btnFatAddIOList.Location = new System.Drawing.Point(336, 23);
            this.btnFatAddIOList.Name = "btnFatAddIOList";
            this.btnFatAddIOList.Size = new System.Drawing.Size(124, 23);
            this.btnFatAddIOList.TabIndex = 7;
            this.btnFatAddIOList.Text = "Add IO List";
            this.btnFatAddIOList.UseVisualStyleBackColor = true;
            this.btnFatAddIOList.Click += new System.EventHandler(this.btnAddIOList_Click);
            // 
            // lstInstrumentFilesList
            // 
            this.lstInstrumentFilesList.FormattingEnabled = true;
            this.lstInstrumentFilesList.Location = new System.Drawing.Point(3, 287);
            this.lstInstrumentFilesList.Name = "lstInstrumentFilesList";
            this.lstInstrumentFilesList.Size = new System.Drawing.Size(327, 108);
            this.lstInstrumentFilesList.TabIndex = 6;
            this.lstInstrumentFilesList.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstInstrumentFilesList_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 269);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Instrument Lists";
            // 
            // lstTermFilesList
            // 
            this.lstTermFilesList.FormattingEnabled = true;
            this.lstTermFilesList.Location = new System.Drawing.Point(3, 156);
            this.lstTermFilesList.Name = "lstTermFilesList";
            this.lstTermFilesList.Size = new System.Drawing.Size(327, 108);
            this.lstTermFilesList.TabIndex = 4;
            this.lstTermFilesList.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstTermFilesList_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(0, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Termination Lists";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "IO Lists";
            // 
            // lstIOFilesList
            // 
            this.lstIOFilesList.FormattingEnabled = true;
            this.lstIOFilesList.Location = new System.Drawing.Point(3, 25);
            this.lstIOFilesList.Name = "lstIOFilesList";
            this.lstIOFilesList.Size = new System.Drawing.Size(327, 108);
            this.lstIOFilesList.TabIndex = 1;
            this.lstIOFilesList.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstIOFilesList_KeyDown);
            // 
            // btnCreateFatSheet
            // 
            this.btnCreateFatSheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateFatSheet.Location = new System.Drawing.Point(463, 460);
            this.btnCreateFatSheet.Name = "btnCreateFatSheet";
            this.btnCreateFatSheet.Size = new System.Drawing.Size(161, 46);
            this.btnCreateFatSheet.TabIndex = 0;
            this.btnCreateFatSheet.Text = "Create FAT Sheet";
            this.btnCreateFatSheet.UseVisualStyleBackColor = true;
            this.btnCreateFatSheet.Click += new System.EventHandler(this.btnCreateFatSheet_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(662, 562);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabControl1.ResumeLayout(false);
            this.tabCompareTermLists.ResumeLayout(false);
            this.tabCompareTermLists.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColStart2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowStart2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColStart1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowStart1)).EndInit();
            this.tabCompareIOList.ResumeLayout(false);
            this.tabCompareIOList.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numFatInstColStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatInstRowStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatTermColStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatTermRowStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatIOColStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numFatIORowStart)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabCompareTermLists;
        private System.Windows.Forms.Button btnCompareLists;
        private System.Windows.Forms.NumericUpDown numColStart1;
        private System.Windows.Forms.Label lblCellStart1;
        private System.Windows.Forms.NumericUpDown numRowStart1;
        private System.Windows.Forms.ComboBox selLatestWorksheet;
        private System.Windows.Forms.Button btnLoadNewTermList;
        private System.Windows.Forms.TabPage tabCompareIOList;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.NumericUpDown numColStart2;
        private System.Windows.Forms.Label lblCellStart2;
        private System.Windows.Forms.NumericUpDown numRowStart2;
        private System.Windows.Forms.ComboBox selPreviousWorksheet;
        private System.Windows.Forms.Button btnLoadPrevTermList;
        private System.Windows.Forms.Label lblPreviousFilename1;
        private System.Windows.Forms.Label lblLatestFilename1;
        private System.Windows.Forms.Label lblProgress1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnCreateFatSheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstIOFilesList;
        private System.Windows.Forms.ListBox lstInstrumentFilesList;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstTermFilesList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnFatAddInstrumentList;
        private System.Windows.Forms.Button btnFatAddTermList;
        private System.Windows.Forms.Button btnFatAddIOList;
        private System.Windows.Forms.Button btnFatAddIOToList;
        private System.Windows.Forms.NumericUpDown numFatIOColStart;
        private System.Windows.Forms.Label lblFatIOStartCell;
        private System.Windows.Forms.NumericUpDown numFatIORowStart;
        private System.Windows.Forms.ComboBox selFatIOWorksheets;
        private System.Windows.Forms.Button btnFatAddTermToList;
        private System.Windows.Forms.NumericUpDown numFatTermColStart;
        private System.Windows.Forms.Label lblFatTermStartCell;
        private System.Windows.Forms.NumericUpDown numFatTermRowStart;
        private System.Windows.Forms.ComboBox selFatTermWorksheets;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnFatAddInstToList;
        private System.Windows.Forms.NumericUpDown numFatInstColStart;
        private System.Windows.Forms.Label lblFatInstStartCell;
        private System.Windows.Forms.NumericUpDown numFatInstRowStart;
        private System.Windows.Forms.ComboBox selFatInstWorksheets;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtFatMarshallingPanels;
    }
}