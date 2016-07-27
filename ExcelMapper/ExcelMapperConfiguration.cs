using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ExcelMapper
{
    class ExcelMapperConfiguration
    {
        private Logger logger = LogManager.GetCurrentClassLogger();
        private ExcelMap excelMapData;
        public ExcelMapperConfiguration(string xmlPath)
        {
            logger.Info("Reading XML mapping file: " + xmlPath);
            var deserializer = new XmlSerializer(typeof(ExcelMap));
            try
            {
                using (TextReader reader = new StreamReader(xmlPath))
                {
                    excelMapData = (ExcelMap)deserializer.Deserialize(reader);
                }
            }
            catch (Exception e)
            {
                logger.Error("Unable to process: " + xmlPath);
                logger.Error("Exiting application. " + e);
                Environment.Exit(-1);
            }
        }

        public string SourceFolder
        {
            get 
            {
                return ((ExcelMapSource)excelMapData.Items[0]).Folder;
            }
        }

        public string SourceFile
        {
            get
            {
                return ((ExcelMapSource)excelMapData.Items[0]).Name;
            }
        }

        public string TargetFolder
        {
            get
            {
                return (((ExcelMapTarget)excelMapData.Items[1])).Folder;
            }
        }

        public string TargetFile
        {
            get
            {
                return ((ExcelMapTarget)excelMapData.Items[1]).Name;
            }
        }

        public ExcelMapRangesCell[] CellRanges
        {
            get
            {
                ExcelMapRanges emr = (ExcelMapRanges)excelMapData.Items[2];
                return emr.Cell;
            }   
        }

        public int StartRow
        {
            get
            {
                var element = (ExcelMapTarget)excelMapData.Items[1];
                return Int32.Parse(element.StartRow);
            }
        }
    }
}
