using System.IO;

using System.Xml.Serialization;
using System;

namespace ExcelMapper
{
    class Program
    {
        static void Main(string[] args)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(ExcelMap));
            TextReader reader = new StreamReader(@"ExcelMapping.xml");
            var excelMapData = (ExcelMap)deserializer.Deserialize(reader);
            using (var xf = new ExcelFile(@"F:\Projects\Open Source\ExcelMapper\ExcelMapper\bin\Debug\excel.xlsx"))
            {
                xf.Process();
                xf.Close();
            }
            reader.Close();
        }
        private static void releaseObject(object obj)
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
