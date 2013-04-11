using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeLibs.Excel
{
    /// <summary>
    /// Summary description for Workbook
    /// </summary>
    public class Workbook
    {
        private object workbook;

        public void Save()
        {
            workbook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, workbook, null);
        }

        public void SaveAs(string fileName)
        {
            workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, workbook, new object[] { fileName });
        }
        public object Name
        {
            get
            {
                return workbook.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, workbook, null);
            }
        }
        public object FullName
        {
            get
            {
                return workbook.GetType().InvokeMember("FullName", BindingFlags.GetProperty, null, workbook, null);
            }
        }
        public object Path
        {
            get
            {
                return workbook.GetType().InvokeMember("Path", BindingFlags.GetProperty, null, workbook, null);
            }
        }

        public void Close(bool saveChanges)
        {
            workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { saveChanges });
            Marshal.ReleaseComObject(workbook);
            GC.GetTotalMemory(true);
        }
        public void Close()
        {
            Close(false);
        }

        //public string FullName
        //{
        //    get
        //    {
        //        return (string)workbook.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, workbook,null);
        //    }
        //    set
        //    {
        //        workbook.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, workbook, new object[] { value });
        //    }
        //}
        public Worksheets Worksheets
        {
            get
            {
                return new Worksheets(workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null));
            }
        }

        public Workbook(object _workbook)
        {
            workbook = _workbook;
        }
    }
}
