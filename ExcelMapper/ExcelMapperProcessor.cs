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
            //ProcessOneToManyFiles(model.OneToMany);
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
                    regNumber = (cellValue ?? string.Empty).ToString().Replace("\n", "");
                    string destinationTemplate = string.Format(entry.Template, string.Empty); //ExcelSummaryTemplatePath(entry.DstFolder, entry.DstFileMask);
                    string destinationFilename = entry.DstFolder + entry.DstFileMask + regNumber + ".xlsx";

                    string destinationFileToOpen = destinationTemplate;
                    if (File.Exists(destinationFilename))
                    {
                        destinationFileToOpen = destinationFilename;
                    }

                    using (var destinationExcelFile = new ExcelFile(destinationFileToOpen))
                    {
                        if (!(finishedProcessing = ProcessManyFile(
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
            foreach (var entry in entries)
            {
                if (!AreNewFilesToProcess(entry.SrcFolder, entry.SrcFileMask))
                {
                    logger.Info("No ManyToOne files of type: {0} to process since last successful run of: {1}",
                        entry.SrcFileMask,
                        settings.LastRun);
                    return;
                }

                string[] filesToProcess = AllFilesToProcess(entry.SrcFolder, entry.SrcFileMask);
                foreach (string file in filesToProcess)
                {
                    ProcessSpecificManyToOneFile(entry, file);
                }
            }
        }

        private void ProcessSpecificManyToOneFile(ManyToOne entry, string file)
        {
            string loanNumber = string.Empty;
            string regNumber = string.Empty;
            int startRowNum;

            startRowNum = Int32.Parse(entry.StartRow);
            using (var sourceExcelFile = new ExcelFile(file))
            {
                loanNumber = (sourceExcelFile.GetCell(entry.SrcSheet, entry.SrcLoanNumberCell) ?? string.Empty).ToString().Replace("\n", "");
                regNumber = (sourceExcelFile.GetCell(entry.SrcSheet, entry.SrcRegNumberCell) ?? string.Empty).ToString().Replace("\n", "");

                string destinationFilename = entry.DstFolder + string.Format(entry.DstFileMask, loanNumber) + ".xls";

                if (!File.Exists(destinationFilename))
                {
                    logger.Warn("Destination Object List file ({0})does not exist!)", destinationFilename);
                }
                else
                {
                    using (var destinationExcelFile = new ExcelFile(destinationFilename))
                    {
                        ProcessOneFile(
                            regNumber,
                            sourceExcelFile,
                            entry,
                            startRowNum,
                            destinationExcelFile);
                        destinationExcelFile.SaveAs(destinationFilename);
                        destinationExcelFile.Close(true);
                        logger.Info("Processed: {0} in folder: {1} into: {2} in folder: {3}",
                            file,
                            entry.SrcFolder,
                            destinationFilename,
                            entry.DstFolder);
                    }
                    sourceExcelFile.Close(false);
                }
            }
        }

        private void ProcessOneFile(string regNumber, ExcelFile srcExcelFile, ManyToOne entry, int startRowNum, ExcelFile dstExcelFile)
        {
            string srcCell = String.Empty;
            string dstCell = String.Empty;
            int destRowNum = FindRegNumberRow(regNumber, startRowNum, entry.SrcSheet, entry.LoanNumberCellMask, dstExcelFile);

            if (-1 == destRowNum)
            {
                logger.Warn("Warning! Unable to find a row matching: {0} in {1}",
                    regNumber,
                    dstExcelFile.Filename);
                return;
            }
            foreach (Cell cell in entry.Cells)
            {
                srcCell = string.Format(cell.SrcCell, destRowNum);
                dstCell = string.Format(cell.DstCell, destRowNum);

                var srcCellContents = (srcExcelFile.GetCell(cell.SrcSheet, srcCell)).ToString().Trim();
                var dstCellContents = (dstExcelFile.GetCell(cell.DstSheet, dstCell)).ToString().Trim();

                if (string.Empty == dstCellContents)
                {
                    dstExcelFile.SetCell(cell.DstSheet, dstCell, srcExcelFile.GetCell(cell.SrcSheet, srcCell));
                    logger.Info(cell.SrcSheet + "!" + srcCell +
                        "(" + srcCellContents + ")" +
                        " => " + cell.DstSheet + "!" + dstCell);
                }
                else
                {
                    logger.Info("Skipping: " + cell.SrcSheet + "!" + srcCell +
                        "(" + srcCellContents + ")" +
                        " => " + cell.DstSheet + "!" + dstCell + " - destination not empty.");
                }
            }
        }

        private int FindRegNumberRow(string regNumber, int startRowNum, string sheet, string cellMask, ExcelFile destinationExcelFile)
        {
            bool found = false;
            int currentRowNum = startRowNum;

            while (false == found)
            {
                var cell = destinationExcelFile.GetCell(sheet, string.Format(cellMask, currentRowNum.ToString()));
                string cellValue = (cell ?? string.Empty).ToString();
                if (string.IsNullOrEmpty(cellValue) || string.IsNullOrWhiteSpace(cellValue))
                {
                    return -1;
                }

                if (cellValue == regNumber)
                {
                    found = true;
                }
                else
                {
                    currentRowNum++;
                }
            }
            return currentRowNum;
        }

        private string ExcelSummaryTemplatePath(string targetFolder, string targetFile)
        {
            return String.Concat(targetFolder, targetFolder.EndsWith("\\") ? "" : "\\", String.Format(targetFile, String.Empty));
        }

        private bool ProcessManyFile(ExcelFile xf, Cell[] cells, int row, ExcelFile destinationExcelFile, string markerSheet, string markerCell)
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