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

namespace ExcelLab
{
    public partial class Form1 : Form
    {
        Excel.Application excapp = null;
        Excel.Window excwindow = null;
        Excel.Workbook excwrkbook = null ;
        Excel.Sheets excshts = null;
        Excel.Worksheet excsht = null;
        Excel.Range exccells = null;

        public Form1()
        {
            InitializeComponent();
            if (r_button1.Checked) r_button2.Checked = false;
            else r_button1.Checked = false;
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            if (excapp != null)
            {
                excapp.Quit();
            }
            excapp = new Excel.Application();
            excapp.Visible = true;
            excapp.SheetsInNewWorkbook = 3;
            excapp.Workbooks.Add(Type.Missing);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string pish = "";
            if (excapp == null){
                excapp = new Excel.Application();
                excapp.Visible = true;
            }
            if(r_button1.Checked)  pish = "Z:/A1.xls";
            else pish = "Z:/A2.xls";
            excapp.Workbooks.Open(pish,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            excapp.Workbooks.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            excapp.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            string sStr = "";

            string pish = "";

            int viRez;
            float vfRez;

            if (excapp == null)
            {
                excapp = new Excel.Application();
                excapp.Visible = true;
            }
            if (r_button1.Checked) pish = "Z:/A1.xls";
            else pish = "Z:/A2.xls";
            excwrkbook=excapp.Workbooks.Open(pish,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing);
            excshts = excwrkbook.Worksheets;
            excsht = (Excel.Worksheet)excshts.get_Item(1);
            //Выбираем ячейку для вывода A1
            exccells= excsht.get_Range("A1", Type.Missing);
            sStr = Convert.ToString(exccells.Value2);
            Text = sStr + " ";
            exccells= excsht.get_Range("A2", Type.Missing);
            viRez = Convert.ToInt32(exccells.Value2);
            Text += Convert.ToString(viRez) + " ";
            exccells= excsht.get_Range("A3", Type.Missing);
            sStr = Convert.ToString(exccells.Value2);
            Text += sStr + " ";
            exccells= excsht.get_Range("A4", Type.Missing);
            vfRez = Convert.ToSingle(exccells.Value2);
            Text += Convert.ToString(vfRez) + " ";

        }

       
       
    }
}
