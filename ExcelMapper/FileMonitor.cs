using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper
{
    public class FileMonitor
    {
        public FileMonitor(string folder, string fileMask)
        {
            Folder = folder;
            FileMask = fileMask;
        }

        public string Folder { get; set; }

        public string FileMask { get; set; }

        public string[] AvailableFiles(DateTime newerThanDate)
        {
            DirectoryInfo d = new DirectoryInfo(Folder);
            FileInfo[] files = d.GetFiles(FileMask); 

            var results = files
                .Where(i => i.LastWriteTime >= newerThanDate)
                .Select(i => i.FullName).ToArray();
            
            return results;
        }
    }
}
