using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        /// <summary>显示/切换页面设置对话框</summary>
        private void btnPaper_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Word 2007/2010 走模态对话框，避免 Ribbon 状态与窗体生命周期冲突
                if (wps.Contains("2010") || wps.Contains("2007"))
                {
                    using (var dlg = new frmPaperSetup())
                        dlg.ShowDialog();
                    btnPaper.Checked = false;
                    return;
                }

                if (fp == null || fp.IsDisposed)
                {
                    fp = new frmPaperSetup();
                    var clsWinWrap = new ClsWinWrap(WordApp.ActiveWindow.Hwnd);
                    clsWinWrap.Show(fp);
                    btnPaper.Checked = true;
                }
                else
                {
                    WordApp.StatusBar = "页面设置对话框已打开";
                    fp.Activate();
                    btnPaper.Checked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "打开页面设置时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                btnPaper.Checked = false;
            }
        }

        /// <summary>中文缩进：首行缩进 2 字符</summary>
        private void btnInsert2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("首行缩进2字符");
                WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "中文缩进时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>段落顶格：清除首行缩进、左缩进</summary>
        private void btnLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("段落顶格");
                WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                WordApp.Selection.ParagraphFormat.FirstLineIndent = 0;
                WordApp.Selection.ParagraphFormat.LeftIndent = 0;
                WordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = 0;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "段落顶格时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>删除当前页页眉横线</summary>
        private void btnHeadNoLine_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除页眉横线");
                WordApp.ScreenUpdating = false;

                if (WordApp.ActiveWindow.ActivePane.View.SeekView != MSWord.WdSeekView.wdSeekCurrentPageHeader)
                    WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;

                WordApp.Selection.WholeStory();
                WordApp.Selection.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleNone;

                WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "删除页眉横线失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                WordApp.ScreenUpdating = true;
            }
        }

        /// <summary>编号转文字（全文转换）</summary>
        private void BtnNum2Txt_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var dr = MessageBox.Show(
                    "编号转换为文字后,现有的项目将无法自动编号.\r\n你确定要继续吗?",
                    "编号转为文字", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr != DialogResult.OK) return;

                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("编号转为文字");
                WordApp.ActiveDocument.Content.ListFormat.ConvertNumbersToText();
                undo.EndCustomRecord();
                WordApp.StatusBar = "编号已全部转为文字形式!";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "编号转文字时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}