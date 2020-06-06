using System;
using System.IO;

namespace ConsoleApp
{
    public class FileDirectoryHelper
    {
        public static String GetDirPath(String path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                return path;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static void CreateFile(String content, String fullPathFile, Boolean append = false)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(fullPathFile, append))
                {
                    sw.Write(content);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
