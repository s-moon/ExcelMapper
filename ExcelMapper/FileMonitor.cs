using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper
{
    public class FileMonitor
    {
        private Logger logger = LogManager.GetCurrentClassLogger();

        public FileMonitor(string folder, string fileMask)
        {
            Folder = folder;
            FileMask = fileMask;
        }

        public string Folder { get; set; }
        public string FileMask { get; set; }

        public string[] AvailableFiles(DateTime newerThanDate)
        {
            var results = new string[0];
            try
            {
                DirectoryInfo d = new DirectoryInfo(Folder);
                FileInfo[] files = d.GetFiles(FileMask);

                results = files
                    //.Where(i => i.LastWriteTime >= newerThanDate)
                    .Select(i => i.FullName).ToArray();
            }
            catch (ArgumentNullException e)
            {
                logger.Error("Is the file or folder name missing? " + e);
                Environment.Exit(-1);
            }
            catch (DirectoryNotFoundException e)
            {
                logger.Error("Unable to access folder: " + Folder + " : "  + e);
                Environment.Exit(-1);
            }
            catch (SecurityException e)
            {
                logger.Error("Oops. This looks like a permissions problem with folder: " + Folder + " : " + e);
                Environment.Exit(-1);
            }
            catch(Exception e)
            {
                logger.Error("Something terrible happened that I couldn't work around. " + e);
                Environment.Exit(-1);
            }
           
            return results;
        }
    }
}
