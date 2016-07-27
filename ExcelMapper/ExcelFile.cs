using System;
using Microsoft.Office.Interop.Excel;
using NLog;
using System.IO;

namespace ExcelMapper
{
    public class ExcelFile : IDisposable
    {
        private Logger logger = LogManager.GetCurrentClassLogger();
        private Application excel;
        private Workbook workbook;
        private Worksheet worksheet;
        public ExcelFile(string filename)
        {
            Filename = filename;
            var excel = new Application();
            if (excel == null)
            {
                logger.Error("Unable to run Microsoft Excel. Is it installed?");
                Environment.Exit(-1);
            }
            if (File.Exists(Filename))
            {
                workbook = excel.Workbooks.Open(Filename, true, true);
            }
            else
            {
                workbook = excel.Workbooks.Add(Filename);
            }
        }

        public string Filename { get; set; }

        public Object GetCell(string sheet, string cell)
        {
            worksheet = (Worksheet)workbook.Sheets[sheet];
            var cellValue = worksheet.Cells.Range[cell].Value2;
            return cellValue;
        }

        internal void SaveAs(string path)
        {
            workbook.SaveAs(path);
        }

        public void SetCell(string sheet, string cell, object value)
        {
            worksheet = (Worksheet)workbook.Sheets[sheet];
            worksheet.Cells.Range[cell].Value2 = value;
        }

        public void Close(bool save = false)
        {
            if (save) { workbook.Save(); }
            if (workbook != null) { workbook.Close(false); }
            if (excel != null) { excel.Quit(); }
        }

        public void Dispose()
        {
            if (worksheet != null) { releaseObject(worksheet); }
            if (workbook != null) { releaseObject(workbook); }
            if (excel != null) { releaseObject(excel); }
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
