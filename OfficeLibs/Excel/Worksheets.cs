using System.Reflection;
namespace OfficeLibs.Excel
{
    /// <summary>
    /// Summary description for Worksheets
    /// </summary>
    public class Worksheets
    {
        private object worksheets;

        public Worksheet get_Worksheet(int index)
        {
            return new Worksheet(worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { index }));
        }
        public Worksheet get_Worksheet(string sheetname)
        {
            return new Worksheet(worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { sheetname }));
        }
        public Worksheet Add()
        {
            return new Worksheet(worksheets.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, worksheets, null));
        }
        public int Count
        {
            get
            {
                return (int)worksheets.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, worksheets, null);
            }
        }
        public void Delete(int index)
        {
            Worksheet sh = get_Worksheet(index);
            sh.Delete();
        }
        public Worksheets(object _worksheets)
        {
            worksheets = _worksheets;
        }
    }
}
