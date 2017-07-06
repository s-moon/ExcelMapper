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
                foreach (string file in filesToProcess)
                {
                    ProcessSpecificOneToManyFile(entry, file);
                }
            }
        }

        private void ProcessSpecificOneToManyFile(OneToMany entry, string file)
        {
            bool finishedProcessing = false;
            string regNumber = string.Empty;
            int startRowNum;
            int rowNum;

            startRowNum = Int32.Parse(entry.StartRow);
            rowNum = startRowNum;
            using (var sourceExcelFile = new ExcelFile(file))
            {
                while (!finishedProcessing)
                {
                    var cellValue = sourceExcelFile.GetCell(entry.DstFileMaskSheet, string.Format(entry.DstFileMaskValue, rowNum));
                    regNumber = (cellValue ?? string.Empty).ToString();
                    string destinationTemplate = ExcelSummaryTemplatePath(entry.DstFolder, entry.DstFileMask);
                    string destinationFilename = destinationTemplate + regNumber.Replace("\n", "") + ".xlsx";

                    string destinationFileToOpen = destinationTemplate;
                    if (File.Exists(destinationFilename))
                    {
                        destinationFileToOpen = destinationFilename;
                    }
                    using (var destinationExcelFile = new ExcelFile(destinationFileToOpen))
                    {
                        if (!(finishedProcessing = ProcessFile(
                            sourceExcelFile,
                            entry.Cells,
                            rowNum,
                            destinationExcelFile,
                            entry.DstFileMaskSheet,
                            entry.DstFileMaskValue
                            )))
                        {
                            destinationExcelFile.SaveAs(destinationFilename);
                            rowNum++;

                            logger.Info("Processed: {0} in folder: {1} into: {2} in folder: {3}",
                            file,
                            entry.SrcFolder,
                            destinationFilename,
                            entry.DstFolder);
                        }
                        destinationExcelFile.Close(true);
                    }
                }
                sourceExcelFile.Close(false);
            }
        }

        private void ProcessManyToOneFiles(ManyToOne[] entries)
        {
            return;
        }

        private string ExcelSummaryTemplatePath(string targetFolder, string targetFile)
        {
            return String.Concat(targetFolder, targetFolder.EndsWith("\\") ? "" : "\\", String.Format(targetFile, String.Empty));
        }

        private bool ProcessFile(ExcelFile xf, Cell[] cells, int row, ExcelFile destinationExcelFile, string markerSheet, string markerCell)
        {
            string sourceCell = String.Empty;
            string destinationCell = String.Empty;

            var rowCellMarker = xf.GetCell(markerSheet, string.Format(markerCell, row));
            if (string.IsNullOrEmpty((rowCellMarker ?? string.Empty).ToString()) || 
                string.IsNullOrWhiteSpace((rowCellMarker ?? string.Empty).ToString()))
            {
                return true;
            }

            for (int i = 0; i < cells.Length; i++)
            {
                sourceCell = string.Format(cells[i].SrcCell, row);
                destinationCell = string.Format(cells[i].DstCell, row);

                logger.Info(cells[i].SrcSheet + "!" + sourceCell +
                    "(" + xf.GetCell(cells[i].SrcSheet, sourceCell) + ")" +
                    " => " + cells[i].DstSheet + "!" + destinationCell);

                var cellContents = xf.GetCell(cells[i].SrcSheet, sourceCell);
                var destCellContents = xf.GetCell(cells[i].DstSheet, destinationCell);

                //if (destCellContents != null)
                //{
                //    logger.Info("Skipping cell " + destCellContents + " of sheet " + cells[i].DstSheet + " because it is non empty (" + destCellContents + ")");
                //}
                //else
                //{
                    destinationExcelFile.SetCell(cells[i].DstSheet, destinationCell, cellContents);
                //}
            }
            return false;
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