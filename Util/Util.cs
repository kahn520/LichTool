using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UtilLib
{
    public static class Util
    {
        public static string SelectFolderDialog()
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() != DialogResult.OK)
                return "";
            return folderBrowser.SelectedPath;
        }

        public static string[] SelectFolderDialog(string strFilder)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() != DialogResult.OK)
                return new string[] { };
            var files = Directory.GetFiles(folderBrowser.SelectedPath).Where(f => f.Contains(strFilder) && !f.Contains("~$")).ToArray();
            return files;
        }
    }

}
