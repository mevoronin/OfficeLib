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
        public string getRangeAdress(int row, int column)
        {
            char[] abc = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            StringBuilder sb = new StringBuilder();
            if (column > 0)
            {
                var multi = (int)Math.Floor(((decimal)column / 26));
                if (multi > 0)
                {
                    sb.Append(abc[multi - 1]);
                }
                sb.Append(abc[(column % 26) - 1]);
            }
            sb.Append(row);
            return sb.ToString(); ;
        }
    }
}
