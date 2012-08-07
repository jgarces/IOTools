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
            this.btnAddInstrumentList = new System.Windows.Forms.Button();
            this.btnAddTermList = new System.Windows.Forms.Button();
            this.btnAddIOList = new System.Windows.Forms.Button();
            this.lstInstrumentLists = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lstTermLists = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lstIOLists = new System.Windows.Forms.ListBox();
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
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabCompareTermLists);
            this.tabControl1.Controls.Add(this.tabCompareIOList);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(638, 477);
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
            this.tabCompareTermLists.Size = new System.Drawing.Size(630, 451);
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
            this.tabCompareIOList.Controls.Add(this.btnAddInstrumentList);
            this.tabCompareIOList.Controls.Add(this.btnAddTermList);
            this.tabCompareIOList.Controls.Add(this.btnAddIOList);
            this.tabCompareIOList.Controls.Add(this.lstInstrumentLists);
            this.tabCompareIOList.Controls.Add(this.label3);
            this.tabCompareIOList.Controls.Add(this.lstTermLists);
            this.tabCompareIOList.Controls.Add(this.label2);
            this.tabCompareIOList.Controls.Add(this.label1);
            this.tabCompareIOList.Controls.Add(this.lstIOLists);
            this.tabCompareIOList.Controls.Add(this.btnCreateFatSheet);
            this.tabCompareIOList.Location = new System.Drawing.Point(4, 22);
            this.tabCompareIOList.Name = "tabCompareIOList";
            this.tabCompareIOList.Padding = new System.Windows.Forms.Padding(3);
            this.tabCompareIOList.Size = new System.Drawing.Size(630, 451);
            this.tabCompareIOList.TabIndex = 1;
            this.tabCompareIOList.Text = "Compare IO Lists";
            this.tabCompareIOList.UseVisualStyleBackColor = true;
            // 
            // btnAddInstrumentList
            // 
            this.btnAddInstrumentList.Location = new System.Drawing.Point(497, 163);
            this.btnAddInstrumentList.Name = "btnAddInstrumentList";
            this.btnAddInstrumentList.Size = new System.Drawing.Size(124, 23);
            this.btnAddInstrumentList.TabIndex = 9;
            this.btnAddInstrumentList.Text = "Add Instrument List";
            this.btnAddInstrumentList.UseVisualStyleBackColor = true;
            // 
            // btnAddTermList
            // 
            this.btnAddTermList.Location = new System.Drawing.Point(290, 164);
            this.btnAddTermList.Name = "btnAddTermList";
            this.btnAddTermList.Size = new System.Drawing.Size(124, 23);
            this.btnAddTermList.TabIndex = 8;
            this.btnAddTermList.Text = "Add Termination List";
            this.btnAddTermList.UseVisualStyleBackColor = true;
            // 
            // btnAddIOList
            // 
            this.btnAddIOList.Location = new System.Drawing.Point(82, 164);
            this.btnAddIOList.Name = "btnAddIOList";
            this.btnAddIOList.Size = new System.Drawing.Size(124, 23);
            this.btnAddIOList.TabIndex = 7;
            this.btnAddIOList.Text = "Add IO List";
            this.btnAddIOList.UseVisualStyleBackColor = true;
            this.btnAddIOList.Click += new System.EventHandler(this.btnAddIOList_Click);
            // 
            // lstInstrumentLists
            // 
            this.lstInstrumentLists.FormattingEnabled = true;
            this.lstInstrumentLists.Location = new System.Drawing.Point(420, 23);
            this.lstInstrumentLists.Name = "lstInstrumentLists";
            this.lstInstrumentLists.Size = new System.Drawing.Size(201, 134);
            this.lstInstrumentLists.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(417, 7);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Instrument Lists";
            // 
            // lstTermLists
            // 
            this.lstTermLists.FormattingEnabled = true;
            this.lstTermLists.Location = new System.Drawing.Point(213, 23);
            this.lstTermLists.Name = "lstTermLists";
            this.lstTermLists.Size = new System.Drawing.Size(201, 134);
            this.lstTermLists.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(210, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Termination Lists";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "IO Lists";
            // 
            // lstIOLists
            // 
            this.lstIOLists.FormattingEnabled = true;
            this.lstIOLists.Location = new System.Drawing.Point(6, 23);
            this.lstIOLists.Name = "lstIOLists";
            this.lstIOLists.Size = new System.Drawing.Size(201, 134);
            this.lstIOLists.TabIndex = 1;
            // 
            // btnCreateFatSheet
            // 
            this.btnCreateFatSheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateFatSheet.Location = new System.Drawing.Point(503, 402);
            this.btnCreateFatSheet.Name = "btnCreateFatSheet";
            this.btnCreateFatSheet.Size = new System.Drawing.Size(121, 46);
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
            this.ClientSize = new System.Drawing.Size(662, 501);
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
        private System.Windows.Forms.ListBox lstIOLists;
        private System.Windows.Forms.ListBox lstInstrumentLists;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstTermLists;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnAddInstrumentList;
        private System.Windows.Forms.Button btnAddTermList;
        private System.Windows.Forms.Button btnAddIOList;
    }
}

