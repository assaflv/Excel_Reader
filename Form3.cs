using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {

        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;



        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }


        private void readExcel(string sFile)
        {
            string[] values_arr = new string[5];
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];          // NAME OF THE SHEET.
            int iRow, iCol = 2;
            int arr_loc = 0;



            if (xlWorkSheet.Cells[5, 9].value == "start_string")
            {

                // START FROM THE SECOND ROW.
                for (iRow = 5; iRow <= xlWorkSheet.Rows.Count; iRow++)
                {
                    if (xlWorkSheet.Cells[iRow, 10].value == null)
                    {
                        break;              // BREAK LOOP.
                    }
                    else
                    {                       // POPULATE COMBO BOX.
                                            //listBox1.Items.Add(xlWorkSheet.Cells[iRow, 9].value);
                        values_arr[arr_loc] = (xlWorkSheet.Cells[iRow, 10].value).ToString();
                        arr_loc = arr_loc + 1;
                    }
                }
                this.dataGridView1.Rows.Add(values_arr);

            }
            

            xlWorkBook.Close();
            xlApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string file_path = @"D:\C_Sharp_Excel\A123.xlsx";
            readExcel(file_path);
     
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //string[] folder = Directory.GetFiles(@"D:\C_Sharp_Excel\", " *.xlsx");
            //string[] folder = Directory.GetFiles(@"D:\C_Sharp_Excel\");

            string dic_path = @"D:\C_Sharp_Excel\";

            DirectoryInfo directory = new DirectoryInfo(dic_path);
            FileInfo[] files = directory.GetFiles();

            var filtered = files.Where(f => !f.Attributes.HasFlag(FileAttributes.Hidden));

            foreach (var f in filtered)
            {
                    try
                    {
                    //MessageBox.Show(dic_path + f.ToString());
                    readExcel(dic_path + f.ToString());
                }
                    catch (Exception)
                    {
                        MessageBox.Show(string.Format("File is corrupt : {0}", f));
                    }
                }
            

        }
    }
}
