using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
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
                    para.Format.FirstLineIndent = 0;

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

        private List<string> listChineseNum = new List<string>() {"一", "二", "三", "四", "五", "六", "七", "八", "九", "十"};
        Regex regexNumPart = new Regex(@"^\d+\．");
        Regex regexNumPart1 = new Regex(@"^\d+\.");

        private void btnFixNum_Click(object sender, RibbonControlEventArgs e)
        {
            var files = Util.SelectFolderDialog(".doc");
            foreach (var file in files)
            {
                int num = 1;
                Document doc = _application.Documents.Open(file);
                foreach (Paragraph paragraph in doc.Paragraphs)
                {
                    if (IsPagePart(paragraph))
                        num = 1;
                    Match match = IsNumPart(paragraph);
                    if (match != null && match.Success)
                    {
                        int start = paragraph.Range.Start;
                        int end = start + match.Length;
                        doc.Range(start, end).Text = $"{num}{match.Value.Last()}";

                        num++;
                    }
                }
                doc.Save();
                doc.Close();
            }
        }

        private bool IsPagePart(Paragraph paragraph)
        {
            if (paragraph.Range.Font.Bold != (int) MsoTriState.msoTrue)
                return false;

            if (paragraph.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                return false;

            if (paragraph.Range.Text.Contains("试卷"))
                return true;

            return false;
        }

        private bool IsBigPart(Paragraph paragraph)
        {
            if (paragraph.Range.Font.Bold != (int) MsoTriState.msoTrue)
                return false;
            foreach (char t in paragraph.Range.Text)
            {
                if (listChineseNum.Contains(t.ToString()))
                    return true;
            }

            return false;
        }

        private Match IsNumPart(Paragraph paragraph)
        {
            if (paragraph.Range.Font.Bold == (int)MsoTriState.msoTrue)
                return null;

            if (paragraph.CharacterUnitLeftIndent > 0)
                return null;

            object isTable = paragraph.Range.Information[WdInformation.wdWithInTable];
            if (isTable != null && (bool) isTable == true)
                return null;

            string text = paragraph.Range.Text;
            Match match = regexNumPart.Match(text);
            if (match.Success)
                return match;

            match = regexNumPart1.Match(text);
            if (match.Success)
                return match;
            return null;
        }
    }
}
