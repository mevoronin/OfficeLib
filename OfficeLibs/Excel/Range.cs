using System.Reflection;
using System;

namespace OfficeLibs.Excel
{
    public enum ColorIndex
    {
        None = -4142,
        Auto = -4105,
        LightGray = 15,
        Gray = 16,
        Black = 1,
        Blue1 = 42,
        DarkBlue = 55,
        LightRed = 18,
        Red = 3,
        Yellow = 6,
        LightGreen = 43,
        Green = 14,
        Perlamutr = 33,
        Phiolet = 47
    }
    public enum XlCellType
    {
        xlCellTypeAllFormatConditions = -4172,
        xlCellTypeAllValidation = -4174,
        xlCellTypeBlanks = 4,
        xlCellTypeComments = -4144,
        xlCellTypeConstants = 2,
        xlCellTypeFormulas = -4123,
        xlCellTypeLastCell = 11,
        xlCellTypeSameFormatConditions = -4173,
        xlCellTypeSameValidation = -4175,
        xlCellTypeVisible = 12
    }
    /// <summary>
    /// Summary description for Range
    /// </summary>
    public class Range
    {
        private object range;

        /// <summary>
        /// Значение ячейки
        /// </summary>
        public object Value
        {
            get
            {
                return range.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, range, null);
            }
            set
            {
                range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, range, new object[] { value });
            }
        }
        /// <summary>
        /// Текст ячейки
        /// </summary>
        public object Text
        {
            get
            {
                return range.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, range, null);
            }
            set
            {
                range.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, range, new object[] { value });
            }
        }

        public object Formula
        {
            get
            {
                return range.GetType().InvokeMember("Formula", BindingFlags.GetProperty, null, range, null);
            }
            set
            {
                range.GetType().InvokeMember("Formula", BindingFlags.SetProperty, null, range, new object[] { value });
            }
        }

        public object NumberFormat
        {
            get
            {
                return range.GetType().InvokeMember("NumberFormat", BindingFlags.GetProperty, null, range, null);
            }
            set
            {
                range.GetType().InvokeMember("NumberFormat", BindingFlags.SetProperty, null, range, new object[] { value });
            }
        }
        /// <summary>
        /// Ширина колонки
        /// </summary>
        public object ColumnWidth
        {
            get
            {
                return range.GetType().InvokeMember("ColumnWidth", BindingFlags.GetProperty, null, range, null);
            }
            set
            {
                range.GetType().InvokeMember("ColumnWidth", BindingFlags.SetProperty, null, range, new object[] { value });
            }
        }
        public bool MergeCells
        {
            get
            {
                return (bool)range.GetType().InvokeMember("MergeCells", BindingFlags.GetProperty, null, range, null);
            }
            set
            {
                range.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, range, new object[] { value });
            }
        }

        public Range(object _range)
        {
            range = _range;
        }
        public ColorIndex ColorIndex
        {
            set
            {
                object interior = range.GetType().InvokeMember("Interior", BindingFlags.GetProperty, null, range, null);
                interior.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, interior, new object[] { (int)value });
            }
        }
        public bool Bold
        {
            get
            {
                object interior = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
                return (bool)interior.GetType().InvokeMember("Bold", BindingFlags.GetProperty, null, interior, null);
            }
            set
            {
                object interior = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
                interior.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, interior, new object[] { value });
            }
        }

        public void Insert()
        {
            object entireRow = range.GetType().InvokeMember("EntireRow", BindingFlags.GetProperty, null, range, null);
            entireRow.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, entireRow, null);
        }
        public void RowsAutoFit()
        {
            object entireRow = range.GetType().InvokeMember("EntireRow", BindingFlags.GetProperty, null, range, null);
            entireRow.GetType().InvokeMember("AutoFit", BindingFlags.InvokeMethod, null, entireRow, null);
        }
        public void ColumnsAutoFit()
        {
            object entireRow = range.GetType().InvokeMember("EntireColumn", BindingFlags.GetProperty, null, range, null);
            entireRow.GetType().InvokeMember("AutoFit", BindingFlags.InvokeMethod, null, entireRow, null);
        }
        public Range SpecialCells(XlCellType cellType)
        {
            return new Range(range.GetType().InvokeMember("SpecialCells", BindingFlags.InvokeMethod, null, range, new object[] { cellType }));
        }
        public int Row { get { return Convert.ToInt32(range.GetType().InvokeMember("Row", BindingFlags.GetProperty, null, range, null)); } }
        public int Column { get { return Convert.ToInt32(range.GetType().InvokeMember("Column", BindingFlags.GetProperty, null, range, null)); } }
    }
}
