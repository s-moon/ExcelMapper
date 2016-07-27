using NLog;
using System;
using System.IO;

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
            if (!AreNewFilesToProcess())
            {
                logger.Info("No files to process since last successful run of: " + settings.LastRun);
                return;
            }

            var filesToProcess = AllFilesToProcess();

            string destinationTemplate = ExcelSummaryTemplatePath(emp.TargetFolder, emp.TargetFile);
            var destinationExcelFile = new ExcelFile(destinationTemplate);

            int row = emp.StartRow;
            foreach (var excelFile in filesToProcess)
            {
                ProcessFile(excelFile, emp.CellRanges, row++, destinationExcelFile);
            }

            string destinationFilename = NextValidExcelDestinationFile(emp.TargetFolder, emp.TargetFile);
            logger.Info("Saving destination file: {0}", destinationFilename);
            destinationExcelFile.SaveAs(destinationFilename);
            destinationExcelFile.Close();

            settings.WriteSetting(Settings.LastRunSetting, DateTime.Now.ToString());
        }

        private bool AreNewFilesToProcess()
        {
            var fm = new FileMonitor(emp.SourceFolder, emp.SourceFile);
            return fm.AvailableFiles(settings.LastRun).Length != 0;
        }

        private string ExcelSummaryTemplatePath(string targetFolder, string targetFile)
        {
            return String.Concat(targetFolder, targetFolder.EndsWith("\\") ? "" : "\\", String.Format(targetFile, String.Empty));
        }

        private string NextValidExcelDestinationFile(string targetFolder, string targetFile)
        {
            string fullPath;
            int count = 0;

            targetFolder = targetFolder.EndsWith("\\") ? targetFolder : targetFolder + "\\";
            do
            {
                count++;
                fullPath = String.Concat(targetFolder, 
                    String.Format(targetFile, 
                    String.Concat(DateTime.Now.ToString("yyyyMMdd"), 
                    "_",
                    count)));
            } while (File.Exists(fullPath));

            return fullPath;    
        }

        private void ProcessFile(string excelFile, ExcelMapRangesCell[] cell, int row, ExcelFile destinationExcelFile)
        {
            string destinationCell = String.Empty;
            logger.Info("Processing: {0}", excelFile);
            var xf = new ExcelFile(excelFile); 
            for (int i = 0; i < cell.Length; i++)
            {
                destinationCell = string.Format(cell[i].TargetCell, row);
                logger.Info(cell[i].SourceSheet + "!" + cell[i].SourceCell +
                    "(" + xf.GetCell(cell[i].SourceSheet, cell[i].SourceCell) + ")" +
                    " => " + cell[i].TargetSheet + "!" + destinationCell);
                var cellContents = xf.GetCell(cell[i].SourceSheet, cell[i].SourceCell);
                destinationExcelFile.SetCell(cell[i].TargetSheet, destinationCell,  cellContents);
            }
            xf.Close(false);
        }

        private string[] AllFilesToProcess()
        {
            var fm = new FileMonitor(emp.SourceFolder, emp.SourceFile);
            return fm.AllFiles();
        }
    }
}