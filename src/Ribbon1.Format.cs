using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using MSWord = Microsoft.Office.Interop.Word;

namespace 公文排版
{
    public partial class Ribbon1
    {
        /// <summary>查找替换</summary>
        private void WordReplace(bool wildcard, string findString, string replaceString, bool range = false)
        {
            try
            {
                WordApp.Selection.Find.ClearFormatting();
                WordApp.Selection.Find.Replacement.ClearFormatting();
                if (range)
                {
                    WordApp.Selection.Find.Execute(findString, false, false, wildcard, false, false, true,
                        MSWord.WdFindWrap.wdFindStop, false, replaceString, MSWord.WdReplace.wdReplaceAll);
                }
                else
                {
                    WordApp.Selection.Find.Execute(findString, false, false, wildcard, false, false, true,
                        MSWord.WdFindWrap.wdFindContinue, false, replaceString, MSWord.WdReplace.wdReplaceAll);
                }
                WordApp.Selection.Find.ClearFormatting();
                WordApp.Selection.Find.Replacement.ClearFormatting();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "排版替换功能出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private bool HasSelectionOrConfirm(string title)
        {
            if (WordApp.Selection.Characters.Count > 1) return true;
            var dr = MessageBox.Show("您未选定范围,将执行全局操作!", title,
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            return dr != DialogResult.Cancel;
        }

        private void btnFormatFullLine_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("断句重排");
            if (!s) return;
            WordApp.ScreenUpdating = false;
            WordReplace(true, "([!。：！？])^13@", "\\1", s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatDelEmptyLine_Click(object sender, RibbonControlEventArgs e)
        {
            int n = 0;
            try
            {
                MSWord.Paragraphs p = WordApp.Selection.Paragraphs;
                if (WordApp.Selection.Characters.Count <= 1)
                {
                    var dr = MessageBox.Show("您未选定范围,将执行全局操作!", "删除空行",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dr == DialogResult.Cancel) return;
                    p = WordApp.ActiveDocument.Paragraphs;
                }
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除空行");
                WordApp.ScreenUpdating = false;
                foreach (MSWord.Paragraph i in p)
                {
                    if (i.Range.Text.Trim().Length <= 0 && i.Range.ShapeRange.Count == 0)
                    {
                        i.Range.Delete();
                        n++;
                    }
                }
                undo.EndCustomRecord();
                WordApp.StatusBar = "已删除 " + n + " 个空行!";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "删除空行时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally { WordApp.ScreenUpdating = true; }
        }

        private void btnFormatDelSelectEmptyLine_Click(object sender, RibbonControlEventArgs e)
        {
            int n = 0;
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除空行");
                MSWord.Paragraphs p = WordApp.Selection.Paragraphs;
                WordApp.ScreenUpdating = false;
                foreach (MSWord.Paragraph i in p)
                {
                    if (i.Range.Text.Trim().Length <= 0 && i.Range.ShapeRange.Count == 0)
                    {
                        i.Range.Delete();
                        n++;
                    }
                }
                undo.EndCustomRecord();
                WordApp.StatusBar = "已删除 " + n + " 个空行!";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "删除空行时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally { WordApp.ScreenUpdating = true; }
        }

        private void btnFormatDelSpace_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("删除空白");
            if (!s) return;
            WordApp.ScreenUpdating = false;
            WordReplace(false, "^w", "", s);
            WordReplace(false, "　", "", s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatR2N_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("换行符转回车");
            if (!s) return;
            WordApp.ScreenUpdating = false;
            WordReplace(false, "^l", "^p", s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatDelGraphic_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("删除图形");
            if (!s) return;
            WordApp.ScreenUpdating = false;
            WordReplace(false, "^g", "", s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatRemoveURL_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                bool s = HasSelectionOrConfirm("移走链接");
                if (!s) return;
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("移走链接");
                WordApp.ScreenUpdating = false;
                if (s) WordApp.Selection.Fields.Unlink();
                else WordApp.ActiveDocument.Fields.Unlink();
                WordApp.ScreenUpdating = true;
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnFormatDelSelectSpace_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("删除半角空格");
            if (!s) return;
            WordApp.ScreenUpdating = false;
            WordReplace(false, " ", "", s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatAddSpace_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("中英文之间加空格");
            if (!s) return;
            WordApp.ScreenUpdating = false;
            WordReplace(true, "([A-Za-z])([㐀-﨩])", "\\1 \\2", s);
            WordReplace(true, "([㐀-﨩])([A-Za-z])", "\\1 \\2", s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatE2C_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = HasSelectionOrConfirm("标点符号(英转中)");
            if (!s) return;
            WordApp.ScreenUpdating = false;

            string strFind = "([!0-9])\\.([!0-9])";
            string strRep = "\\1。\\2";
            WordReplace(true, strFind, strRep, s);

            string[] eArray = { ",", ";", ":", "?", "!", "-", "~", "(", ")", "[", "]" };
            string[] cArray = { "，", "；", "：", "？", "！", "—", "～", "（", "）", "［", "］" };
            for (int i = 0; i < eArray.Length; i++)
                WordReplace(false, eArray[i], cArray[i], s);

            if (!IsWps)
            {
                strFind = "\"(*)\"";
                strRep = "“\\1”";
                WordReplace(true, strFind, strRep, s);
            }
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatC2E_Click(object sender, RibbonControlEventArgs e)
        {
            if (WordApp.Options.AutoFormatAsYouTypeReplaceQuotes)
            {
                string text = "Word自动更正选项中,已开启输入直引号自动更正为弯引号." +
                    "\r\n此功能将影响中文引号的转换.要关闭此选项并开始转换吗?" +
                    "\r\n\r\n点击\"是\"将关闭该自动更正选项并转换.\r\n点击\"否\"将直接转换,引号可能被Word自动更正." +
                    "\r\n点击\"取消\"退出该操作.";
                var dr = MessageBox.Show(text, "中文标点符号转为英文", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes) WordApp.Options.AutoFormatAsYouTypeReplaceQuotes = false;
                else if (dr == DialogResult.Cancel) return;
            }

            bool s = HasSelectionOrConfirm("标点符号(中转英)");
            if (!s) return;
            WordApp.ScreenUpdating = false;

            string[] cArray = { "。", "，", "；", "：", "？", "！", "—", "～", "（", "）", "［", "］", "“", "”", "‘", "’" };
            string[] eArray = { ".", ",", ";", ":", "?", "!", "-", "~", "(", ")", "[", "]", "\"", "\"", "'", "'" };
            for (int i = 0; i < cArray.Length; i++)
                WordReplace(false, cArray[i], eArray[i], s);
            WordApp.ScreenUpdating = true;
        }

        private void btnFormatFrame2txt_Click(object sender, RibbonControlEventArgs e) => Frame2txt();

        private void Frame2txt()
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("文本框转文字");
            try
            {
                if (WordApp.Selection.ShapeRange.Count > 0)
                {
                    MSWord.Shape sh = WordApp.Selection.ShapeRange[1];
                    string t = sh.TextFrame.TextRange.Text;
                    sh.Delete();
                    WordApp.Selection.TypeText(t);
                }
            }
            finally { undo.EndCustomRecord(); }
        }

        private void btnFormatParaPage_Click(object sender, RibbonControlEventArgs e)
        {
            try { WordApp.Run("ParaPageBreakBefore"); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "所选段落另起一页",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}