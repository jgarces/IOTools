﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using IRISSolutions.IOTools;

namespace IOToolsUI
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

            if (!Directory.Exists(@"C:\Temp")) {
                Directory.CreateDirectory(@"C:\Temp");
            }
        }

        /****************************************************************
         * Compare Termination Lists
         ****************************************************************/

        private struct ExcelDataFile
        {
            public string Filename { get; set; }
            public string Worksheet { get; set; }
            public int StartRow { get; set; }
            public int StartCol { get; set; }

            public ExcelDataFile(string filename)
                : this()
            {
                this.Filename = filename;
            }
            public ExcelDataFile(string filename, string ws)
                : this()
            {
                this.Filename = filename;
                this.Worksheet = ws;
            }
            public ExcelDataFile(string filename, string ws, int sR, int sC)
                : this()
            {
                this.Filename = filename;
                this.Worksheet = ws;
                this.StartRow = sR;
                this.StartCol = sC;
            }

            public override string ToString()
            {
                FileInfo fi = new FileInfo(Filename);
                return fi.Name + " :: [" + Worksheet + "] (" + StartRow + "," + StartCol + ")";
            }

            public int GetNumberOfRows()
            {
                FileInfo fi = new FileInfo(this.Filename);
                ExcelPackage ep = new ExcelPackage(fi);
                ExcelWorksheet ws = ep.Workbook.Worksheets[this.Worksheet];
                return ws.Dimension.End.Row;
            }

            public Dictionary<string, string> AsDictionary()
            {
                return new Dictionary<string, string>()
                {
                    { "filename", this.Filename },
                    { "worksheet", this.Worksheet },
                    { "startRow", this.StartRow.ToString() },
                    { "startCol", this.StartCol.ToString() },
                    { "numRows", this.GetNumberOfRows().ToString() },
                };
            }
        }

        private ExcelDataFile edfLatest;
        private ExcelDataFile edfPrevious;

        private void btnLoadNewTermList_Click(object sender, EventArgs e)
        {
            selLatestWorksheet.Items.Clear();
            lblLatestFilename1.Visible = false;
            lblCellStart1.Visible = false;
            numColStart1.Visible = false;
            numRowStart1.Visible = false;
            selLatestWorksheet.Visible = false;
            
            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            if (File.Exists(sFilePath))
            {
                FileInfo fi = new FileInfo(sFilePath);
                //FileInfo nfi = fi.CopyTo("data\\" + fi.Name, true);
                FileInfo nfi = fi.CopyTo(@"C:\Temp\" + fi.Name, true);

                ExcelPackage epLatest = new ExcelPackage(nfi);

                edfLatest = new ExcelDataFile();
                edfLatest.Filename = nfi.FullName;

                lblLatestFilename1.Text = nfi.Name;
                lblLatestFilename1.Visible = true;

                foreach (ExcelWorksheet ws in epLatest.Workbook.Worksheets)
                {
                    selLatestWorksheet.Items.Add(ws.Name);
                }

                selLatestWorksheet.Visible = true;
            }
        }

        private void selLatestWorksheet1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //edfLatest.Worksheet = selLatestWorksheet.SelectedItem.ToString();
            //CurDataFile.StartRow = Convert.ToInt32(numFatIORowStart.Value);
            //CurDataFile.StartCol = Convert.ToInt32(numFatIOColStart.Value);

            lblCellStart1.Visible = true;
            numColStart1.Visible = true;
            numRowStart1.Visible = true;

            if (selPreviousWorksheet.Visible && selLatestWorksheet.Visible)
            {
                btnCompareLists.Visible = true;
            }
        }

        private void btnLoadPrevTermList_Click(object sender, EventArgs e)
        {
            selPreviousWorksheet.Items.Clear();
            lblPreviousFilename1.Visible = false;
            lblCellStart2.Visible = false;
            numColStart2.Visible = false;
            numRowStart2.Visible = false;
            selPreviousWorksheet.Visible = false;

            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            if (File.Exists(sFilePath))
            {
                FileInfo fi = new FileInfo(sFilePath);
                //FileInfo nfi = fi.CopyTo("data\\" + fi.Name, true);
                FileInfo nfi = fi.CopyTo(@"C:\Temp\" + fi.Name, true);

                ExcelPackage epPrevious = new ExcelPackage(nfi);

                edfPrevious = new ExcelDataFile();
                edfPrevious.Filename = nfi.FullName;

                lblPreviousFilename1.Text = nfi.Name;
                lblPreviousFilename1.Visible = true;

                foreach (ExcelWorksheet ws in epPrevious.Workbook.Worksheets)
                {
                    selPreviousWorksheet.Items.Add(ws.Name);
                }

                selPreviousWorksheet.Visible = true;
            }
        }

        private void selPreviousWorksheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblCellStart2.Visible = true;
            numColStart2.Visible = true;
            numRowStart2.Visible = true;

            if (selPreviousWorksheet.Visible && selLatestWorksheet.Visible)
            {
                btnCompareLists.Visible = true;
            }
            
        }

        private void btnCompareLists_Click(object sender, EventArgs e)
        {

            lblProgress1.Text = "Comparing...";
            this.Refresh();

            edfLatest.Worksheet = selLatestWorksheet.SelectedItem.ToString();
            edfLatest.StartRow = Convert.ToInt32(numRowStart1.Value);
            edfLatest.StartCol = Convert.ToInt32(numColStart1.Value);

            edfPrevious.Worksheet = selPreviousWorksheet.SelectedItem.ToString();
            edfPrevious.StartRow = Convert.ToInt32(numRowStart2.Value);
            edfPrevious.StartCol = Convert.ToInt32(numColStart2.Value);

            string strDiffFilename = API.compareTerminationLists(edfLatest.AsDictionary(), edfPrevious.AsDictionary());

            if (strDiffFilename != "")
            {
                lblProgress1.Text = "Finished";

                saveFileDialog1.Filter = "Excel File (*.xlsm) | *.xlsm";
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = strDiffFilename;
                saveFileDialog1.Title = "Save file with differences as";
                saveFileDialog1.OverwritePrompt = false;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string strNewFilename = saveFileDialog1.FileName;

                    if (strNewFilename != strDiffFilename)
                    {
                        FileInfo fiDiffFile = new FileInfo(strDiffFilename);
                        fiDiffFile.CopyTo(strNewFilename);
                    }
                }

            }
            else
            {
                lblProgress1.Text = "There was an error";
            }
        }

        /****************************************************************
         * Create FAT Sheets
         ****************************************************************/

        private List<ExcelDataFile> FatIOLists = new List<ExcelDataFile>();
        private List<ExcelDataFile> FatTermLists = new List<ExcelDataFile>();
        private List<ExcelDataFile> FatInstLists = new List<ExcelDataFile>();

        private ExcelDataFile CurDataFile;

        /**
         * IO Lists
         */

        private void btnAddIOList_Click(object sender, EventArgs e)
        {
            // reset
            CurDataFile = new ExcelDataFile();
            selFatIOWorksheets.Items.Clear();
            selFatIOWorksheets.Text = "Select Worksheet";
            // disable other file open buttons
            btnFatAddTermList.Enabled = false;
            btnFatAddInstrumentList.Enabled = false;

            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            if (File.Exists(sFilePath) )
            {
                FileInfo fi = new FileInfo(sFilePath);
                //FileInfo nfi = fi.CopyTo("data\\" + fi.Name, true);
                FileInfo nfi = fi.CopyTo(@"C:\Temp\" + fi.Name, true);
                ExcelPackage ep = new ExcelPackage(nfi);

                CurDataFile.Filename = nfi.FullName;
                
                foreach (ExcelWorksheet ws in ep.Workbook.Worksheets)
                {
                    selFatIOWorksheets.Items.Add(ws.Name);
                }

                selFatIOWorksheets.Visible = true;
            }
        }

        private void selFatIOWorksheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblFatIOStartCell.Visible = true;
            numFatIORowStart.Visible = true;
            numFatIOColStart.Visible = true;
            btnFatAddIOToList.Visible = true;

            CurDataFile.Worksheet = selFatIOWorksheets.SelectedItem.ToString();
            CurDataFile.StartRow = Convert.ToInt32(numFatIORowStart.Value);
            CurDataFile.StartCol = Convert.ToInt32(numFatIOColStart.Value);
        }

        private void numFatIORowStart_ValueChanged(object sender, EventArgs e)
        {
            CurDataFile.StartRow = Convert.ToInt32(numFatIORowStart.Value);
        }

        private void numFatIOColStart_ValueChanged(object sender, EventArgs e)
        {
            CurDataFile.StartCol = Convert.ToInt32(numFatIOColStart.Value);
        }

        private void btnFatAddIOToList_Click(object sender, EventArgs e)
        {
            // Add the curDataFile to the List IO list
            FatIOLists.Add(CurDataFile);

            _repopulateList(FatIOLists, lstIOFilesList);

            // Hide and reset controls
            selFatIOWorksheets.Visible = false;
            selFatIOWorksheets.Items.Clear();
            lblFatIOStartCell.Visible = false;
            numFatIORowStart.Visible = false;
            numFatIOColStart.Visible = false;
            btnFatAddIOToList.Visible = false;
            btnFatAddTermList.Enabled = true;
            btnFatAddInstrumentList.Enabled = true;
        }

        private void lstIOFilesList_KeyDown(object sender, KeyEventArgs e)
        {
            _removeItemFromList(FatIOLists, lstIOFilesList, e);
        }


        /**
         * Term Lists
         */

        private void btnAddTermList_Click(object sender, EventArgs e)
        {
            // reset
            CurDataFile = new ExcelDataFile();
            selFatTermWorksheets.Items.Clear();
            selFatTermWorksheets.Text = "Select Worksheet";
            // disable other file open buttons
            btnFatAddIOList.Enabled = false;
            btnFatAddInstrumentList.Enabled = false;

            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            if (File.Exists(sFilePath))
            {
                FileInfo fi = new FileInfo(sFilePath);
                //FileInfo nfi = fi.CopyTo("data\\" + fi.Name, true);
                FileInfo nfi = fi.CopyTo(@"C:\Temp\" + fi.Name, true);
                ExcelPackage ep = new ExcelPackage(nfi);

                CurDataFile.Filename = nfi.FullName;

                foreach (ExcelWorksheet ws in ep.Workbook.Worksheets)
                {
                    selFatTermWorksheets.Items.Add(ws.Name);
                }

                selFatTermWorksheets.Visible = true;
            }
        }

        private void selFatTermWorksheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblFatTermStartCell.Visible = true;
            numFatTermRowStart.Visible = true;
            numFatTermColStart.Visible = true;
            btnFatAddTermToList.Visible = true;

            CurDataFile.Worksheet = selFatTermWorksheets.SelectedItem.ToString();
            CurDataFile.StartRow = Convert.ToInt32(numFatTermRowStart.Value);
            CurDataFile.StartCol = Convert.ToInt32(numFatTermColStart.Value);
        }

        private void numFatTermRowStart_ValueChanged(object sender, EventArgs e)
        {
            CurDataFile.StartRow = Convert.ToInt32(numFatTermRowStart.Value);
        }

        private void numFatTermColStart_ValueChanged(object sender, EventArgs e)
        {
            CurDataFile.StartCol = Convert.ToInt32(numFatTermColStart.Value);
        }

        private void btnFatAddTermToList_Click(object sender, EventArgs e)
        {
            // Add the curDataFile to the List IO list
            FatTermLists.Add(CurDataFile);

            _repopulateList(FatTermLists, lstTermFilesList);

            // Hide and reset controls
            selFatTermWorksheets.Visible = false;
            selFatTermWorksheets.Items.Clear();
            lblFatTermStartCell.Visible = false;
            numFatTermRowStart.Visible = false;
            numFatTermColStart.Visible = false;
            btnFatAddTermToList.Visible = false;
            btnFatAddIOList.Enabled = true;
            btnFatAddInstrumentList.Enabled = true;
        }

        private void lstTermFilesList_KeyDown(object sender, KeyEventArgs e)
        {
            _removeItemFromList(FatTermLists, lstTermFilesList, e);
        }


        /**
         * Instrument Lists
         */

        private void btnFatAddInstrumentList_Click(object sender, EventArgs e)
        {
            // reset
            CurDataFile = new ExcelDataFile();
            selFatInstWorksheets.Items.Clear();
            selFatInstWorksheets.Text = "Select Worksheet";
            // disable other file open buttons
            btnFatAddIOList.Enabled = false;
            btnFatAddTermList.Enabled = false;

            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            if (File.Exists(sFilePath))
            {
                FileInfo fi = new FileInfo(sFilePath);
                //FileInfo nfi = fi.CopyTo("data\\" + fi.Name, true);
                FileInfo nfi = fi.CopyTo(@"C:\Temp\" + fi.Name, true);
                ExcelPackage ep = new ExcelPackage(nfi);

                CurDataFile.Filename = nfi.FullName;

                foreach (ExcelWorksheet ws in ep.Workbook.Worksheets)
                {
                    selFatInstWorksheets.Items.Add(ws.Name);
                }

                selFatInstWorksheets.Visible = true;
            }
        }

        private void selFatInstWorksheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblFatInstStartCell.Visible = true;
            numFatInstRowStart.Visible = true;
            numFatInstColStart.Visible = true;
            btnFatAddInstToList.Visible = true;

            CurDataFile.Worksheet = selFatInstWorksheets.SelectedItem.ToString();
            CurDataFile.StartRow = Convert.ToInt32(numFatInstRowStart.Value);
            CurDataFile.StartCol = Convert.ToInt32(numFatInstColStart.Value);
        }

        private void numFatInstRowStart_ValueChanged(object sender, EventArgs e)
        {
            CurDataFile.StartRow = Convert.ToInt32(numFatInstRowStart.Value);
        }

        private void numFatInstColStart_ValueChanged(object sender, EventArgs e)
        {
            CurDataFile.StartCol = Convert.ToInt32(numFatInstColStart.Value);
        }

        private void btnFatAddInstToList_Click(object sender, EventArgs e)
        {
            // Add the curDataFile to the List IO list
            FatInstLists.Add(CurDataFile);

            _repopulateList(FatInstLists, lstInstrumentFilesList);

            // Hide and reset controls
            selFatInstWorksheets.Visible = false;
            selFatInstWorksheets.Items.Clear();
            lblFatInstStartCell.Visible = false;
            numFatInstRowStart.Visible = false;
            numFatInstColStart.Visible = false;
            btnFatAddInstToList.Visible = false;
            btnFatAddIOList.Enabled = true;
            btnFatAddTermList.Enabled = true;
        }

        private void lstInstrumentFilesList_KeyDown(object sender, KeyEventArgs e)
        {
            _removeItemFromList(FatInstLists, lstInstrumentFilesList, e);
        }

        private void _repopulateList(List<ExcelDataFile> list, ListBox lb)
        {
            // Clear the list display and re-add whats left in FatIOList
            lb.Items.Clear();
            foreach (ExcelDataFile edf in list)
            {
                FileInfo fi = new FileInfo(edf.Filename);
                lb.Items.Add(fi.Name + " [" + edf.Worksheet + "] " + " (" + edf.StartRow + "," + edf.StartCol + ") ");
            }
        }

        private void _removeItemFromList(List<ExcelDataFile> list, ListBox lb, KeyEventArgs e)
        {
            if (e.KeyValue == 46 && lb.Items.Count > 0)
            {
                // Remove from the FatIOList 
                list.RemoveAt(lb.SelectedIndex);

                _repopulateList(list, lb);
            }
        }

        private void btnCreateFatSheet_Click(object sender, EventArgs e)
        {
            btnCreateFatSheet.Text = "Generating...";
            this.Refresh();

            List<Dictionary<string, string>> ioSheets = new List<Dictionary<string, string>>();
            foreach (ExcelDataFile edf in FatIOLists)
            {
                ioSheets.Add(edf.AsDictionary());
            }

            List<Dictionary<string, string>> termSheets = new List<Dictionary<string, string>>();
            foreach (ExcelDataFile edf in FatTermLists)
            {
                termSheets.Add(edf.AsDictionary());
            }

            List<Dictionary<string, string>> instrumentSheets = new List<Dictionary<string, string>>();
            foreach (ExcelDataFile edf in FatInstLists)
            {
                instrumentSheets.Add(edf.AsDictionary());
            }

            List<string> marshallingPanels = new List<string>();
            string[] mps = txtFatMarshallingPanels.Text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            marshallingPanels.AddRange(mps);

            string fatSheetFilename = API.createMPFat(ioSheets.AsEnumerable(), termSheets.AsEnumerable(), instrumentSheets.AsEnumerable(), marshallingPanels);

            if (fatSheetFilename != "")
            {
                btnCreateFatSheet.Text = "Create FAT Sheet";
                this.Refresh();

                saveFileDialog1.Filter = "Excel File (*.xlsm) | *.xlsm";
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = "FAT-Sheet.xlsm";
                saveFileDialog1.Title = "Save file as";
                saveFileDialog1.OverwritePrompt = false;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string strNewFilename = saveFileDialog1.FileName;

                    if (strNewFilename != fatSheetFilename)
                    {
                        FileInfo fiFatFile = new FileInfo(fatSheetFilename);
                        fiFatFile.CopyTo(strNewFilename);
                        fiFatFile.Delete();
                    }
                }
            }
        }

    }
}