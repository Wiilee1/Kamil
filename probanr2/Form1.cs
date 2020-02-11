using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace probanr2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        string sFileName;
        int iRow, iCol = 2;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            
            openFileDialog1.Title = "Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog1.FileName;

                if (sFileName.Trim() != "")
                {
                    readExcel(sFileName);
                }
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void readExcel(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);           // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];      // NAME OF THE SHEET.

            for (iRow = 1; iRow <= xlWorkSheet.Rows.Count; iRow++)  // START FROM THE SECOND ROW.
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {               // POPULATE COMBO BOX.
                    cmbEmp.Items.Add(xlWorkSheet.Cells[iRow, 1].value);
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();
        }
    }
}
