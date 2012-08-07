using System;
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

        public FileInfo fiLatest;
        public FileInfo fiPrevious;
        public ExcelPackage epLatest;
        public ExcelPackage epPrevious;

        public Form1()
        {
            InitializeComponent();
        }

        /****************************************************************
         * Compare Termination Lists
         ****************************************************************/

        private void btnLoadNewTermList_Click(object sender, EventArgs e)
        {
            selLatestWorksheet.Items.Clear();
            lblLatestFilename1.Visible = false;
            lblCellStart1.Visible = false;
            numColStart1.Visible = false;
            numRowStart1.Visible = false;
            selLatestWorksheet.Visible = false;
            
            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "All Files (*.*)|*.*";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            // make sure the file exists before adding
            // its path to the list of files to be
            // compressed
            if (File.Exists(sFilePath) == false)
            {
                return;
            }
            else
            {
                //TerminationList tList = new TerminationList(sFilePath);
                fiLatest = new FileInfo(sFilePath);

                epLatest = new ExcelPackage(fiLatest);

                lblLatestFilename1.Text = fiLatest.Name;
                lblLatestFilename1.Visible = true;

                foreach (ExcelWorksheet ws in epLatest.Workbook.Worksheets)
                {
                    Console.WriteLine(ws.Name);
                    selLatestWorksheet.Items.Add(ws.Name);
                }

                selLatestWorksheet.Visible = true;
            }
        }

        private void selLatestWorksheet1_SelectedIndexChanged(object sender, EventArgs e)
        {
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
            openFileDialog1.Filter = "All Files (*.*)|*.*";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            // make sure the file exists before adding
            // its path to the list of files to be
            // compressed
            if (File.Exists(sFilePath) == false)
            {
                return;
            }
            else
            {
                //TerminationList tList = new TerminationList(sFilePath);
                fiPrevious = new FileInfo(sFilePath);

                epPrevious = new ExcelPackage(fiPrevious);

                lblPreviousFilename1.Text = fiPrevious.Name;
                lblPreviousFilename1.Visible = true;

                foreach (ExcelWorksheet ws in epPrevious.Workbook.Worksheets)
                {
                    Console.WriteLine(ws.Name);
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

            string strLatestFilename = fiLatest.FullName;
            string strPrevFilename = fiPrevious.FullName;

            string strLatestWorksheet = selLatestWorksheet.SelectedItem.ToString();
            string strPrevWorksheet = selPreviousWorksheet.SelectedItem.ToString();

            int iLatestNumRows = epLatest.Workbook.Worksheets[strLatestWorksheet].Dimension.End.Row;
            int iPrevNumRows = epPrevious.Workbook.Worksheets[strPrevWorksheet].Dimension.End.Row;

            int iLatestRowStart = Convert.ToInt32(numRowStart1.Value);
            int iLatestColStart = Convert.ToInt32(numColStart1.Value);
            int iPrevRowStart = Convert.ToInt32(numRowStart2.Value);
            int iPrevColStart = Convert.ToInt32(numColStart2.Value);

            string strDiffFilename = API.compareTerminationLists(strLatestFilename, strLatestWorksheet, iLatestRowStart, iLatestColStart, iLatestNumRows,
                                                                strPrevFilename, strPrevWorksheet, iPrevRowStart, iPrevColStart, iPrevNumRows);

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

        private void btnAddIOList_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "All Files (*.xlsm)|*.xlsm";
            openFileDialog1.FileName = "";

            openFileDialog1.ShowDialog();

            string sFilePath;
            sFilePath = openFileDialog1.FileName;
            if (sFilePath == "")
                return;

            // make sure the file exists before adding
            // its path to the list of files to be
            // compressed
            if (File.Exists(sFilePath) == false)
            {
                return;
            }
            else
            {
                //TerminationList tList = new TerminationList(sFilePath);
                fiLatest = new FileInfo(sFilePath);

                epLatest = new ExcelPackage(fiLatest);

                lblLatestFilename1.Text = fiLatest.Name;
                lblLatestFilename1.Visible = true;

                foreach (ExcelWorksheet ws in epLatest.Workbook.Worksheets)
                {
                    Console.WriteLine(ws.Name);
                    selLatestWorksheet.Items.Add(ws.Name);
                }

                selLatestWorksheet.Visible = true;
            }
        }

        private void btnCreateFatSheet_Click(object sender, EventArgs e)
        {



            //List<Dictionary<string, string>> ioSheets = new List<Dictionary<string, string>>();
            //Dictionary<string, string> ioSheet1 = new Dictionary<string,string>()
            //{
            //    { "filename", @"C:\Users\rpattison\Documents\Karara\MP-280\1317-IN-LST-1001_3.001.120523.xlsm" },
            //    { "worksheet", "1317-SR-107" },
            //    { "startRow", "3" },
            //    { "startCol", "1" },
            //    { "numRows", "1128" },
            //};

            //ioSheets.Add(ioSheet1);

            //List<Dictionary<string, string>> termSheets = new List<Dictionary<string, string>>();
            //Dictionary<string, string> termSheet1 = new Dictionary<string, string>()
            //{
            //    { "filename", @"C:\Users\rpattison\Documents\Karara\MP-280\1317-IN-LST-1031_1.001.120523.xlsm" },
            //    { "worksheet", "1317-IN-LST-1031_1" },
            //    { "startRow", "6" },
            //    { "startCol", "1" },
            //    { "numRows", "921" },
            //};
            //Dictionary<string, string> termSheet2 = new Dictionary<string, string>()
            //{
            //    { "filename", @"C:\Users\rpattison\Documents\Karara\MP-280\1320-IN-LST-1031_D.001.120523.xlsm" },
            //    { "worksheet", "1320-IN-LST-1031_D" },
            //    { "startRow", "6" },
            //    { "startCol", "1" },
            //    { "numRows", "310" },
            //};

            //termSheets.Add(termSheet1);
            //termSheets.Add(termSheet2);

            //List<Dictionary<string, string>> instrumentSheets = new List<Dictionary<string, string>>();
            //Dictionary<string, string> instrSheet1 = new Dictionary<string, string>()
            //{
            //    { "filename", @"C:\Users\rpattison\Documents\Karara\MP-280\1300-IN-LST-1004_1_BMcL.xlsm" },
            //    { "worksheet", "1301 IN List bmcl revs" },
            //    { "startRow", "2" },
            //    { "startCol", "1" },
            //    { "numRows", "6767" },
            //};

            //instrumentSheets.Add(instrSheet1);

            //List<string> marshallingPanels = new List<string>(){ "1317-MP-280" };

            //API.createMPFat(ioSheets.AsEnumerable(), termSheets.AsEnumerable(), instrumentSheets.AsEnumerable(), marshallingPanels);
        }

        


    }
}
