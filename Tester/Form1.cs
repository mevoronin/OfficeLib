using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeLibs.Excel;

namespace Tester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel();
            Workbook book = excel.Workbooks.Add();
            excel.Visible = true;
            Worksheet sheet = book.Worksheets.get_Worksheet(1);
            Range range = sheet.Range("A1");
            range.Value = "test";
            range.SetColor(ColorEnum.Green);
        }
    }
}
