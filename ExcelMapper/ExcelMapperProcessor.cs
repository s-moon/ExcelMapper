using NLog;
using System;

namespace ExcelMapper
{
    internal class ExcelMapperProcessor
    {
        private Logger logger = LogManager.GetCurrentClassLogger();
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
            if (filesToProcess == null || filesToProcess.Length == 0)
            {
                logger.Info("No files to process since last successful run of: " + settings.LastRun);
                return;
            }

            int row = emp.StartRow;
            foreach (var excelFile in filesToProcess)
            {
                ProcessFile(excelFile, emp.CellRanges, row++);
            }

            //settings.WriteSetting(Settings.LastRunSetting, DateTime.Now.ToString());
        }

        private void ProcessFile(string excelFile, ExcelMapRangesCell[] cell, int row)
        {
            logger.Info("Processing: {0}", excelFile);
            var xf = new ExcelFile(excelFile); 
            for (int i = 0; i < cell.Length; i++)
            {
                logger.Info(cell[i].SourceSheet + "!" + cell[i].SourceCell +
                    "(" + xf.GetCell(cell[i].SourceSheet, cell[i].SourceCell) + ")" +
                    " => " + cell[i].TargetSheet + "!" + string.Format(cell[i].TargetCell, row));
            }
            xf.Close(false);
        }

        private string[] FilesToProcess()
        {
            var fm = new FileMonitor(emp.SourceFolder, emp.SourceFile);
            return fm.AvailableFiles(settings.LastRun);
        }
    }
}