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
            xlWorkSheet = xlWorkBook.Worksheets[2];      // NAME OR NUMBER OF THE SHEET.

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

        private void loadCustomers_button_Click(object sender, EventArgs e)  // LOAD FILE BUTTON
        {
            string str;
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    myStream = openFileDialog1.OpenFile();
                    StreamReader sr = new StreamReader(myStream);
                    string[] split;

                    while ((str = sr.ReadLine()) != null)
                    {

                        split = str.Split(new char[] { ',' });
                        string name = split[0];
                        string address = split[1];
                        string email = split[2];
                        string county = split[3];
                        double phone = System.Convert.ToDouble(split[4]);
                        double reference = System.Convert.ToDouble(split[5]);


                        customers.Add(new Customer(name, address, email, county, phone, reference));
                    }


                    sr.Close();
                    myStream.Close();

                    listView1.Items.Clear();

                    foreach (Customer c in customers)
                    {
                        string name = c.getName();
                        string address = c.getAddress();
                        string email = c.getEmail();
                        string county = c.getCounty();
                        double phone = c.getPhone();
                        double reference = c.getReference();
                        string phonestring = System.Convert.ToString(phone);
                        string referencestring = System.Convert.ToString(reference);

                        ListViewItem item = new ListViewItem(name);
                        item.SubItems.Add(address);
                        item.SubItems.Add(email);
                        item.SubItems.Add(county);
                        item.SubItems.Add(phonestring);
                        item.SubItems.Add(referencestring);
                        listView1.Items.Add(item);
                    }



                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void saveCustomer_button_Click(object sender, EventArgs e) // SAVE FILE BUTTON
        {

            Stream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "c:\\";
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            string[] output = new string[customers.Count];
            int j = 0;
            foreach (Customer c in customers)
            {
                output[j] = c.getName() + "," + c.getAddress() + "," + c.getEmail() + "," + c.getCounty() + "," +
                    c.getPhone() + "," + c.getReference();        
                j++;
            }



            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    myStream = saveFileDialog1.OpenFile();
                    StreamWriter sw = new StreamWriter(myStream);
                    for (int i = 0; i < customers.Count; i++)
                    {
                        sw.WriteLine(output[i]);
                    }

                    sw.Close();
                    myStream.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }
    }
}
