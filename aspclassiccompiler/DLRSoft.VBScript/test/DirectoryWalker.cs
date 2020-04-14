using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;

namespace Dlrsoft.VBScriptTest
{
    public class DirectoryWalker
    {
        static public string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        static public void Walk(string path, Action<string> fileHandler)
        {
            FileAttributes fa = File.GetAttributes(path);
            if ((fa & FileAttributes.Directory) == FileAttributes.Directory)
            {
                Walk(new DirectoryInfo(path), fileHandler);
            }
            else
            {
                Walk(new FileInfo(path), fileHandler);
            }
        }

        static private void Walk(DirectoryInfo di, Action<string> fileHandler)
        {
            foreach (FileInfo fi in di.GetFiles())
            {
                Walk(fi, fileHandler);
            }

            foreach (DirectoryInfo cdi in di.GetDirectories())
            {
                Walk(cdi, fileHandler);
            }
        }

        static private void Walk(FileInfo fi, Action<string> fileHandler)
        {
            fileHandler(fi.FullName);
        }
    }
}
