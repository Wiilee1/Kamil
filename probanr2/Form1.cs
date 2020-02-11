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
    public partial class Form1 : System.Windows.Forms.Form
    {
        List<Customer> customers = new List<Customer>();
        public Form1()
        {
            InitializeComponent();
        }

        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        string sFileName;

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

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void readExcel(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);           // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets[2];      // NAME OF THE SHEET.

            string name, address, email, county;
            double phone, reference;

            name = xlWorkSheet.Cells[11, 2].value;
            address = xlWorkSheet.Cells[11, 4].value;
            email = xlWorkSheet.Cells[11, 8].value;
            county = xlWorkSheet.Cells[39, 7].value;
            phone = xlWorkSheet.Cells[11, 6].value;
            string phonestring = System.Convert.ToString(phone);
            reference = xlWorkSheet.Cells[7, 3].value;
            string refstring = System.Convert.ToString(reference);

            ListViewItem item = new ListViewItem(xlWorkSheet.Cells[11, 2].value);
            item.SubItems.Add(address);
            item.SubItems.Add(email);
            item.SubItems.Add(county);
            item.SubItems.Add(phonestring);
            item.SubItems.Add(refstring);
            listView1.Items.Add(item);
            customers.Add(new Customer(name, address, email, county, phone, reference));

            xlWorkBook.Close();
            xlApp.Quit();
        }


    }
}
