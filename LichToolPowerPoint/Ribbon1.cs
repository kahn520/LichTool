using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using UtilLib;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace LichToolPowerPoint
{
    public partial class Ribbon1
    {
        private Application _application;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _application = Globals.ThisAddIn.Application;
        }

        private void btnSaveAsPptx_Click(object sender, RibbonControlEventArgs e)
        {
            string strFolder = Util.SelectFolderDialog();
            if (strFolder != "")
                SaveAsPptx(strFolder);
        }

        private void SaveAsPptx(string strFolder)
        {
            var files = Directory.GetFiles(strFolder).Where(f => f.EndsWith(".doc") && !f.Contains("~$")).ToArray();
            foreach (string file in files)
            {
                Presentation pres = _application.Presentations.Open(file);
                pres.SaveAs(file.Replace(".ppt", ".pptx"));
                pres.Close();
            }
        }
    }
}
