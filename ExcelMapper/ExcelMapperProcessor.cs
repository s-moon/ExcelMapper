using System;

namespace ExcelMapper
{
    internal class ExcelMapperProcessor
    {
        Settings settings;
        ExcelMapperConfiguration emp;

        public ExcelMapperProcessor()
        {
            settings = new Settings();
            emp = new ExcelMapperConfiguration(settings.MappingXMLPath);
        }

        public void ProcessFiles()
        {
            var filesToProcess = FilesToProcess();
            foreach (var excelFile in filesToProcess)
            {
                ProcessFile(excelFile, emp.Ranges.Cell);
            }

            settings.WriteSetting(Settings.LastRunSetting, DateTime.Now.ToString());
        }

        private void ProcessFile(string excelFile, ExcelMapRangesCell[] cell)
        {
            
            // open Excel file

            // copy all cells to destination
            // read settings
            //var x = emp.Ranges;

            //for (int i = 0; i < x.Cell.Length; i++)
            //{
            //    Console.WriteLine("From:" + x.Cell[i].SourceSheet + "!" + x.Cell[i].SourceCell); ;
            //    Console.WriteLine("To:" + x.Cell[i].TargetSheet + "!" + x.Cell[i].TargetCell);
            //}
        }

        private string[] FilesToProcess()
        {
            var fm = new FileMonitor(emp.SourceFolder, emp.SourceFile);
            return fm.AvailableFiles(settings.LastRun);
        }

        private ExcelMapRangesCell[] GenerateRangesToCopy()
        {
            return null;
        }
    }
}