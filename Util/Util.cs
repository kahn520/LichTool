using System;
using System.Collections.Generic;
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
    }

}
