using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using UtilLib;
using Application = Microsoft.Office.Interop.Word.Application;

namespace LichToolWord
{
    public partial class Ribbon1
    {
        private Application _application;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _application = Globals.ThisAddIn.Application;
        }

        private void btnReplaceTitle_Click(object sender, RibbonControlEventArgs e)
        {
            string strFolder = Util.SelectFolderDialog();
            if (strFolder != "")
                ReplaceTitle(strFolder);
        }

        private void ReplaceTitle(string strFolder)
        {
            Regex regex = new Regex(@"\(\d+\)");
            var files = Directory.GetFiles(strFolder).Where(f => f.Contains(".doc") && !f.Contains("~$")).ToArray();
            string[] listEmpty = new[] { "\r", "\n", "\a", "\v" };
            foreach (string file in files)
            {
                Document doc = _application.Documents.Open(file);
                int paracnt = doc.Paragraphs.Count;
                for (int i = 1; i <= paracnt; i++)
                {
                    Paragraph para = doc.Paragraphs[i];
                    string text = para.Range.Text;
                    if (string.IsNullOrEmpty(text))
                        continue;

                    foreach (var empty in listEmpty)
                    {
                        text = text.Replace(empty, "");
                    }

                    if (string.IsNullOrWhiteSpace(text) || string.IsNullOrEmpty(text))
                        continue;

                    int cntTable = para.Range.Tables.Count;
                    if (cntTable > 0)
                        continue;

                    string title = Path.GetFileNameWithoutExtension(file);
                    title = regex.Replace(title, "");
                    title = title.Trim();
                    title += "\r\a";

                    para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    para.Range.Text = title;

                    break;
                }
                doc.Save();
                doc.Close();
            }
        }

        private void btnSaveAsDocx_Click(object sender, RibbonControlEventArgs e)
        {
            string strFolder = Util.SelectFolderDialog();
            if (strFolder != "")
                SaveAsDocx(strFolder);
        }

        private void SaveAsDocx(string strFolder)
        {
            var files = Directory.GetFiles(strFolder).Where(f => f.EndsWith(".doc") && !f.Contains("~$")).ToArray();
            foreach (string file in files)
            {
                Document doc = _application.Documents.Open(file);
                doc.SaveAs(file.Replace(".doc", ".docx"), WdSaveFormat.wdFormatXMLDocument);
                doc.Close();
            }
        }
    }
}
