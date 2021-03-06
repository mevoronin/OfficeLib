﻿using System.Reflection;
using System;

namespace OfficeLibs.Excel
{
    public enum ColorIndexEnum
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
    public enum ColorEnum
    {
        Black=0,
        Blue=16711680,
        Cyan = 16776960 ,
        Green = 65280 ,
        Magenta = 16711935 ,
        Red = 255 ,
        White = 16777215 ,
        Yellow = 65535 
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
    public enum XlDeleteShiftDirection
    {
        xlShiftToLeft = -4159, //Cells are shifted to the left. 
        xlShiftUp = -4162 //Cells are shifted up. 
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
        public object NumberFormatLocal
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
        public ColorIndexEnum ColorIndex
        {
            set
            {
                object interior = range.GetType().InvokeMember("Interior", BindingFlags.GetProperty, null, range, null);
                interior.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, interior, new object[] { (int)value });
            }
        }
        public ColorEnum Color
        {
            set
            {
                object font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
                font.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, font, new object[] { (int)value });
            }
        }
        public void SetColor(ColorEnum color)
        {
            object font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
            font.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, font, new object[] { (int)color });
        }
        public void SetColor(int color)
        {
            object font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
            font.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, font, new object[] { color });
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
        public void Delete()
        {
            range.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, range, null);
        }
        public void Delete(XlDeleteShiftDirection deleteDirection)
        {
            range.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, range, new object[] { deleteDirection });
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

        public void Copy()
        {
            range.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, range, null);
        }
        public void Select()
        {
            range.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, range, null);
        }
        public void PasteSpecial()
        {
            range.GetType().InvokeMember("PasteSpecial", BindingFlags.InvokeMethod, null, range, null);
        }

    }
}
