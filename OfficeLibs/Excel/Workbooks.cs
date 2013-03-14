using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeLibs.Excel
{
    /// <summary>
    /// Summary description for Workbooks
    /// </summary>
    public class Workbooks
    {
        private object workbooks;

        public Workbook Add()
        {
            return new Workbook(workbooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, workbooks, null));
        }

        public Workbook Open(string fileName)
        {
            return new Workbook(workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { fileName }));
        }
        public Workbook Item(int index)
        {
            return new Workbook(workbooks.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, workbooks, new object[] { index }));
        }

        public int Count
        {
            get
            {
                return (int)workbooks.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, workbooks, null);
            }
        }
        public void Close()
        {
            workbooks.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbooks, null);
            Marshal.ReleaseComObject(workbooks);
            GC.GetTotalMemory(true);
        }

        public Workbooks(object _workbooks)
        {
            workbooks = _workbooks;
        }
    }
}
