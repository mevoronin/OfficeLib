using System;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;


namespace OfficeLibs.Excel
{
    /// <summary>
    /// Summary description for Excel
    /// </summary>
    public class Excel : IDisposable
    {
        private object application;
        private bool isNew;

        public bool Visible
        {
            get
            {
                return Convert.ToBoolean(application.GetType().InvokeMember("Visible", BindingFlags.GetProperty, null, application, null));
            }
            set
            {
                application.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, application, new object[] { value });
            }
        }
        public bool CutCopyMode
        {
            get
            {
                return Convert.ToBoolean(application.GetType().InvokeMember("CutCopyMode", BindingFlags.GetProperty, null, application, null));
            }
            set
            {
                application.GetType().InvokeMember("CutCopyMode", BindingFlags.SetProperty, null, application, new object[] { value });
            }
        }

        public void Quit()
        {
            if (isNew) application.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, application, null);
            Marshal.ReleaseComObject(application);
            GC.GetTotalMemory(true);
        }
        public void Dispose()
        {
            Marshal.ReleaseComObject(application);
            GC.GetTotalMemory(true);
        }
        public Workbooks Workbooks
        {
            get
            {
                return new Workbooks(application.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, application, null));
            }
        }

        public object DecimalSeparator
        {
            get
            {
                return application.GetType().InvokeMember("DecimalSeparator", BindingFlags.GetProperty, null, application, null);
            }
            set
            {
                application.GetType().InvokeMember("DecimalSeparator", BindingFlags.SetProperty, null, application, new object[] { value });
            }
        }

        public object ThousandsSeparator
        {
            get
            {
                return application.GetType().InvokeMember("ThousandsSeparator", BindingFlags.GetProperty, null, application, null);
            }
            set
            {
                application.GetType().InvokeMember("ThousandsSeparator", BindingFlags.SetProperty, null, application, new object[] { value });
            }
        }

        public object UseSystemSeparators
        {
            get
            {
                return application.GetType().InvokeMember("UseSystemSeparators", BindingFlags.GetProperty, null, application, null);
            }
            set
            {
                application.GetType().InvokeMember("UseSystemSeparators", BindingFlags.SetProperty, null, application, new object[] { value });
            }
        }

        public Excel()
        {
            string sAppProgID = "Excel.Application";
            try
            {
                application = Marshal.GetActiveObject(sAppProgID);
                isNew = false;
            }
            catch (COMException ex)
            {
                Type tExcelObj = Type.GetTypeFromProgID(sAppProgID);
                application = Activator.CreateInstance(tExcelObj);
                isNew = true;
            }
        }
        public Excel(bool needNew)
        {

            string sAppProgID = "Excel.Application";
            if (needNew)
            {
                Type tExcelObj = Type.GetTypeFromProgID(sAppProgID);
                application = Activator.CreateInstance(tExcelObj);
                isNew = needNew;
            }
            else
            {
                try
                {
                    application = Marshal.GetActiveObject(sAppProgID);
                    isNew = false;
                }
                catch (COMException ex)
                {
                    Type tExcelObj = Type.GetTypeFromProgID(sAppProgID);
                    application = Activator.CreateInstance(tExcelObj);
                    isNew = true;
                }
            }
        }
        public int SheetsInNewWorkbook
        {
            get
            {
                var count = application.GetType().InvokeMember("SheetsInNewWorkbook", BindingFlags.GetProperty, null, application, null);
                return int.Parse(count.ToString());
            }
            set
            {
                application.GetType().InvokeMember("SheetsInNewWorkbook", BindingFlags.SetProperty, null, application, new object[] { value });
            }
        }

        public Range ActiveCell
        {
            get
            {
                object cell=application.GetType().InvokeMember("ActiveCell", BindingFlags.GetProperty, null, application, null);
                if (cell != null)
                    return new Range(cell);
                else
                    return null;
            }
        }
    }







}

