using ExcelMapper.Model;
using Newtonsoft.Json;
using NLog;
using System;
using System.IO;

namespace ExcelMapper
{
    internal class ExcelMapperProcessor
    {
        private Logger logger = LogManager.GetCurrentClassLogger();
        private Settings settings;
        private MapModel model;

        public ExcelMapperProcessor()
        {
            settings = new Settings();
            model = JsonConvert.DeserializeObject<MapModel>(File.ReadAllText(settings.MappingJSONPath));
        }

        public void ProcessFiles()
        {
            ProcessOneToManyFiles(model.OneToMany);
            ProcessManyToOneFiles(model.ManyToOne);

            settings.WriteSetting(Settings.LastRunSetting, DateTime.Now.ToString());
        }

        private void ProcessOneToManyFiles(OneToMany[] entries)
        {
            foreach (var entry in entries)
            {
                if (!AreNewFilesToProcess(entry.SrcFolder, entry.SrcFilemask))
                {
                    logger.Info("No OneToMany files of type: {0} to process since last successful run of: {1}", 
                        entry.SrcFilemask, 
                        settings.LastRun);
                    return;
                }

                string[] filesToProcess = AllFilesToProcess(entry.SrcFolder, entry.SrcFilemask);

                int rowNum = Int32.Parse(entry.StartRow);
                foreach (string file in filesToProcess)
                {
                    string destinationTemplate = ExcelSummaryTemplatePath(entry.DstFolder, entry.DstFileMask);
                    var destinationExcelFile = new ExcelFile(destinationTemplate);

                    logger.Info("Processing: {0} in folder: {1} into: {2} in folder: {3}", 
                        file, 
                        entry.SrcFolder,
                        destinationTemplate,
                        entry.DstFolder);

                    ProcessFile(file, entry.Cells, rowNum++, destinationExcelFile);
                }
            }
            

            //var filesToProcess = AllFilesToProcess();

            //string destinationTemplate = ExcelSummaryTemplatePath(emp.TargetFolder, emp.TargetFile);
            //var destinationExcelFile = new ExcelFile(destinationTemplate);

            //int row = emp.StartRow;
            //foreach (var excelFile in filesToProcess)
            //{
            //    ProcessFile(excelFile, emp.CellRanges, row++, destinationExcelFile);
            //}

            //string destinationFilename = NextValidExcelDestinationFile(emp.TargetFolder, emp.TargetFile);
            //logger.Info("Saving destination file: {0}", destinationFilename);
            //destinationExcelFile.SaveAs(destinationFilename);
            //destinationExcelFile.Close();
        }

        private void ProcessManyToOneFiles(ManyToOne[] entries)
        {
            return;
        }

        private string ExcelSummaryTemplatePath(string targetFolder, string targetFile)
        {
            return String.Concat(targetFolder, targetFolder.EndsWith("\\") ? "" : "\\", String.Format(targetFile, String.Empty));
        }

        private string NextValidExcelDestinationFile(string targetFolder, string targetFile)
        {
            //string fullPath;
            //int count = 0;

            //targetFolder = targetFolder.EndsWith("\\") ? targetFolder : targetFolder + "\\";
            //do
            //{
            //    count++;
            //    fullPath = String.Concat(targetFolder, 
            //        String.Format(targetFile, 
            //        String.Concat(DateTime.Now.ToString("yyyyMMdd"), 
            //        "_",
            //        count)));
            //} while (File.Exists(fullPath));

            //return fullPath;    
            return string.Empty;
        }

        private void ProcessFile(string excelFile, Cell[] cells, int row, ExcelFile destinationExcelFile)
        {
            string destinationCell = String.Empty;
            logger.Info("Processing: {0}", excelFile);
            var xf = new ExcelFile(excelFile);
            for (int i = 0; i < cells.Length; i++)
            {
                destinationCell = string.Format(cells[i].DstCell, row);
                logger.Info(cells[i].SrcSheet + "!" + cells[i].SrcCell +
                    "(" + xf.GetCell(cells[i].SrcSheet, cells[i].SrcCell) + ")" +
                    " => " + cells[i].DstSheet + "!" + destinationCell);
                var cellContents = xf.GetCell(cells[i].SrcSheet, cells[i].SrcCell);
                destinationExcelFile.SetCell(cells[i].DstSheet, destinationCell, cellContents);
            }
            xf.Close(false);
        }

        private bool AreNewFilesToProcess(string folder, string file)
        {
            var fm = new FileMonitor(folder, file);
            return fm.AvailableFiles(settings.LastRun).Length != 0;
        }

        private string[] AllFilesToProcess(string folder, string file)
        {
            var fm = new FileMonitor(folder, file);
            return fm.AvailableFiles(settings.LastRun);
        }
    }
}