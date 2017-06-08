using System.IO;

using System.Xml.Serialization;
using System;
using NLog;

namespace ExcelMapper
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            logger.Info("Excel Mapper starting.");

            var processor = new ExcelMapperProcessor();
            processor.ProcessFiles();

            logger.Info("Excel Mapper ending.");
            logger.Info("");
        }
    }
}
