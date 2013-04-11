using System;
using System.Reflection;
using System.Text;

namespace OfficeLibs.Excel
{
    /// <summary>
    /// Summary description for Worksheet
    /// </summary>
    public class Worksheet
    {
        private object worksheet;

        public Range Range(string range)
        {
            return new Range(worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { range }));
        }
        public Range Cell(int row, int column)
        {
            return new Range(worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { getRangeAdress(row, column) }));
        }

        public Worksheet(object _worksheet)
        {
            worksheet = _worksheet;
        }
        public void Delete()
        {
            worksheet.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, worksheet, null);
        }
        public string Name
        {
            get
            {
                return (string)worksheet.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, worksheet, null);
            }
            set
            {
                worksheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, worksheet, new object[] { value });
            }
        }
        public void Paste()
        {
            worksheet.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, worksheet, null);
        }
        public void Select()
        {
            worksheet.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, worksheet, null);
        }

        public string getRangeAdress(int row, int column)
        {
            string result = "";
            column = column - 1;
            while (column >= 26)
            {
                result = Convert.ToChar(65 + (column % 26)) + result;
                column = (column / 26) - 1;
            }
            return Convert.ToChar(65 + column) + result + row.ToString();
        }

    }
}
