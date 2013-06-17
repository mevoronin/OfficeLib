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
            sheet.Cell(1, 1).Value = "11";
            sheet.Cell(1, 2).Value = "12";
            sheet.Cell(1, 3).Value = "13";

            sheet.Cell(2, 1).Value = "21";
            sheet.Cell(2, 2).Value = "22";
            sheet.Cell(2, 3).Value = "23";

            sheet.Cell(3, 1).Value = "31";
            sheet.Cell(3, 2).Value = "32";
            sheet.Cell(3, 3).Value = "33";
            sheet.Range(sheet.Cell(1, 1), sheet.Cell(2, 3)).Delete(XlDeleteShiftDirection.xlShiftUp);
        }
    }
}
