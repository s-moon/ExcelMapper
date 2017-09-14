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
        public ExcelFile(string template, string filename)
        {
            Filename = filename;
            var excel = new Application();
            if (excel == null)
            {
                logger.Error("Unable to run Microsoft Excel. Is it installed?");
                Environment.Exit(-1);
            }

            excel.DisplayAlerts = false;

            if (File.Exists(Filename))
            {
                // update the links but don't open as read only
                workbook = excel.Workbooks.Open(Filename, true, false);
            }
            else
            {
                if (File.Exists(template))
                {
                    workbook = excel.Workbooks.Open(Filename, true, false);
                }
                else
                {
                    workbook = excel.Workbooks.Add(Filename);
                }
            }
        }

        public string Filename { get; set; }

        public Object GetCell(string sheet, string cell)
        {
            worksheet = (Worksheet)workbook.Sheets[sheet];
            var cellValue = worksheet.Cells.Range[cell].Value2;
            return cellValue;
        }

        public Object GetPictureCell(string sheet, int item)
        {
            worksheet = (Worksheet)workbook.Sheets[sheet];
            var picture = worksheet.Shapes.Item(item);
            return picture;
        }

        internal void SaveAs(string path)
        {
            workbook.SaveAs(path);
        }

        public void SetPictureCell(string sheet, string cell, object value)
        {
            worksheet = (Worksheet)workbook.Sheets[sheet];
            Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)worksheet.Range[cell];
            float Left = (float)((double)oRange.Left);
            float Top = (float)((double)oRange.Top);
            const float ImageSize = 32;
            //worksheet.Shapes.AddPicture("C:\\pic.JPG", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, ImageSize, ImageSize);
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
                logger.Error("Problems releasing object ( " + ex + ")");
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
