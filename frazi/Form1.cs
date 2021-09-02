using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System;
    
    namespace frazi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string s = "dfdgfjhiknssshtygogmgbsss s hgtun sss";
            Properties.Settings.Default.s2 = s;
            Properties.Settings.Default.Save();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string imatab;
            string s1 = textBox1.Text;

            //Excel.Application excelApp = new Excel.Application();
            //Excel.Workbook workBook = null; // Создаём экземпляр рабочий книги Excel
           // Excel.Workbooks workbooks = null;
           // Excel.Worksheet workSheet;// Создаём экземпляр листа Excel
           // Excel.ListObject listObject;
          //  Excel.Application app = null;

         //   app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
           //       as Excel.Application;

           // workbooks = excelApp.Workbooks;
          //  workBook = workbooks.Open(imatab);
          //  workSheet = workBook.Worksheets[1];
          //  listObject = workSheet.ListObjects[1];
          //  string name = listObject.Name;
           
            string s3=Properties.Settings.Default.s2;
            
            int n=0;
            int l = s1.Length;
            int k = 0;
            while (n != -1)
            {
                n = s3.IndexOf(s1);
                if (n == -1) { break; };

                s3 = s3.Remove(n, l);
                k++;
                
            }
            Properties.Settings.Default.s2 = s3;
            Properties.Settings.Default.Save();

            MessageBox.Show(k.ToString());

        }
        }
}
