using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelMapper
{
    public class ExcelFile : IDisposable
    {
        private Application excel;
        private Workbook workbook;
        private Worksheet worksheet;
        private string filename;
        public ExcelFile(string filename)
        {
            this.filename = filename;
            var excel = new Application();
            if (excel == null)
            {
                Console.WriteLine("oops - no excel");
                Environment.Exit(1);
            }
            workbook = excel.Workbooks.Open(filename, true, true);
            var cellk = (Worksheet)workbook.Worksheets;
        }

        public void Process()
        {
            worksheet = (Worksheet)workbook.Sheets["Sheet1"];
            var cellValue = worksheet.Cells[1, 1];
            var x = cellValue.Value2;
        }

        public void Close()
        {
            if (workbook != null)
            {
                workbook.Close(false);
            }
            
            if (excel != null)
            {
                excel.Quit();
            }
        }

        public void Dispose()
        {
            if (worksheet != null)
            {
                releaseObject(worksheet);
            }

            if (workbook != null)
            {
                releaseObject(workbook);
            }
            
            if (excel != null)
            {
                releaseObject(excel);
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("oops - cleaning up");
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
