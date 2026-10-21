using System;
using System.Diagnostics;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        private void selectCurrentPage()
        {
            try
            {
                WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                WordApp.ActiveDocument.Bookmarks["\\page"].Range.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "选择整页时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SelectShape()
        {
            try
            {
                if (WordApp.ActiveDocument != null)
                    WordApp.ActiveDocument.Shapes.SelectAll();
            }
            catch (Exception ex) { WordApp.StatusBar = ex.Message; }
        }

        private void ConverSharp(int type = 1)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("转换图片环绕方式");
                int i = 0;
                foreach (MSWord.InlineShape shape in WordApp.ActiveDocument.InlineShapes)
                {
                    switch (type)
                    {
                        case 1: shape.ConvertToShape(); i++; break;
                        case 2: shape.ConvertToShape().WrapFormat.Type = MSWord.WdWrapType.wdWrapSquare; i++; break;
                        case 3: shape.ConvertToShape().WrapFormat.Type = MSWord.WdWrapType.wdWrapTopBottom; i++; break;
                    }
                }
                undo.EndCustomRecord();
                WordApp.StatusBar = "已转换 " + i + "个图形!";
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void ConverUnderShape()
        {
            int i = 0;
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("转换衬于文字下图片");
                foreach (MSWord.Shape shape in WordApp.ActiveDocument.Shapes)
                {
                    if (shape.WrapFormat.Type == MSWord.WdWrapType.wdWrapBehind)
                    {
                        shape.WrapFormat.Type = MSWord.WdWrapType.wdWrapFront;
                        i++;
                    }
                }
            }
            catch (Exception ex) { WordApp.StatusBar = ex.Message; }
            finally
            {
                MessageBox.Show("共转换 " + i + "个衬于文字下方的图形为浮于文字上方!", "小恐龙公文排版助手");
            }
        }

        private void delPageHeader()
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除页眉");
                WordApp.ScreenUpdating = false;
                if (WordApp.ActiveWindow.ActivePane.View.SeekView != MSWord.WdSeekView.wdSeekCurrentPageHeader)
                    WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
                WordApp.Selection.WholeStory();
                WordApp.Selection.Delete();
                WordApp.Selection.ClearFormatting();

                var style = WordApp.ActiveDocument.Styles.Add("公文:页眉", MSWord.WdStyleType.wdStyleTypeParagraph);
                style.set_BaseStyle("");
                WordApp.Selection.set_Style(style);

                WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "删除页面横线失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally { WordApp.ScreenUpdating = true; }
        }

        private void splitDocBySection()
        {
            MSWord.Document olddoc = WordApp.ActiveDocument;
            if (olddoc.Sections.Count <= 1)
            {
                MessageBox.Show("拆分文档前请插入分节符!", "小恐龙公文排版助手:拆分文档",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                WordApp.CommandBars["Office Clipboard"].Visible = true;
                WordApp.CommandBars["Office Clipboard"].Visible = false;
            }
            catch
            {
                MessageBox.Show("Office 功能不完整(比如精简版、绿色版）,拆分文档将无法继续!",
                    "拆分文档遇到错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int j = olddoc.Sections.Count + 1;
            for (int i = 1; i < j; i++)
            {
                olddoc.Sections[i].Range.Copy();
                MSWord.Document newdoc = WordApp.Documents.Add();
                newdoc.Content.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                newdoc.PageSetup = olddoc.Sections[i].PageSetup;
                WordReplace(false, "^b", "");
            }
        }

        private void SaveSelect()
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.AddExtension = true;
                sfd.Title = "将所选内容保存为可编辑的新文档";
                string fname = "新文档";
                try
                {
                    fname = WordApp.Selection.Range.Paragraphs[1].Range.Text;
                    fname = System.Text.RegularExpressions.Regex.Replace(fname, @"[^a-zA-Z0-9\u4e00-\u9fa5\s]", "");
                    fname = fname.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "")
                        .Replace("\\", "").Replace("*", "").Replace("?", "").Replace(">", "")
                        .Replace("/", "").Replace("<", "").Replace("|", "").Replace("\"", "").Replace(":", "");
                    if (fname.Length > 20) fname = fname.Substring(0, 20);
                }
                finally { sfd.FileName = fname; }

                sfd.DefaultExt = ".docx";
                sfd.Filter = "Word文档(*.docx)|*.docx|标准 HTML 文档(*.htm)|*.htm|" +
                    "单一 HTML 文档(*.mht)|*.mht|Word XML 文档(*.xml)|*.xml|文本文档(*.txt)|*.txt";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                WordApp.StatusBar = "正在将所选内容存为可编辑的新文档... ";
                int wf;
                switch (sfd.FilterIndex)
                {
                    case 1: wf = 16; break;
                    case 2: wf = 8; break;
                    case 3: wf = 9; break;
                    case 4: wf = 12; break;
                    case 5: wf = 2; break;
                    default: wf = 16; break;
                }
                try
                {
                    WordApp.Selection.Range.ExportFragment(sfd.FileName, (MSWord.WdSaveFormat)wf);
                    WordApp.StatusBar = "所选内容已保存!";
                    WordApp.Documents.Open(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\r\n" + ex.Source);
                    WordApp.StatusBar = "";
                }
            }
        }

        private void SaveAsLongPic()
        {
            // 保留实验方法,不建议使用
            MessageBox.Show("该功能不稳定,已停用.", "小恐龙公文排版助手",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnSelectSharp_Click(object sender, RibbonControlEventArgs e)
        {
            SelectShape();
            Globals.Ribbons.Ribbon1.Tabs[0].RibbonUI.ActivateTabMso("TabPictureToolsFormat");
        }

        private void BtnSelectTables_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (WordApp.ActiveDocument.ProtectionType >= 0)
                {
                    WordApp.StatusBar = "文档受到保护,无法选择!";
                    return;
                }
                if (WordApp.ActiveDocument != null)
                {
                    WordApp.Selection.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                    foreach (MSWord.Table table in WordApp.ActiveDocument.Tables)
                        table.Range.Editors.Add(-1);
                    WordApp.ActiveDocument.SelectAllEditableRanges(-1);
                    WordApp.ActiveDocument.DeleteAllEditableRanges(-1);
                    WordApp.ActiveWindow.View.ShadeEditableRanges = 0;
                    Globals.Ribbons.Ribbon1.Tabs[0].RibbonUI.ActivateTabMso("TabTableToolsDesign");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.Source, "选择所有表格时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSelectCurrentPage_Click(object sender, RibbonControlEventArgs e) => selectCurrentPage();
        private void BtnHeadDelete_Click(object sender, RibbonControlEventArgs e) => delPageHeader();
        private void btnSplit_Click(object sender, RibbonControlEventArgs e) => splitDocBySection();
        private void btnSharp2Top_Click(object sender, RibbonControlEventArgs e) => ConverUnderShape();
        private void btnSharp2OnText_Click(object sender, RibbonControlEventArgs e) => ConverSharp();
        private void btnSharp2Wrapping_Click(object sender, RibbonControlEventArgs e) => ConverSharp(2);
        private void btnSharp2TopBonttom_Click(object sender, RibbonControlEventArgs e) => ConverSharp(3);

        private void groupTest_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!WordApp.ActiveDocument.Saved) return;
                string docPath = WordApp.ActiveDocument.FullName;
                var uriFile = new Uri(docPath);
                if (uriFile.IsFile)
                {
                    Process.Start(@"explorer.exe", "/select,\"" + docPath + "\"");
                }
                else
                {
                    string shortName = docPath;
                    for (int i = 0; i < 4; i++)
                    {
                        int index = shortName.IndexOf('/');
                        if (index >= 0)
                        {
                            shortName = shortName.Substring(index + 1);
                            docPath = Environment.GetEnvironmentVariable("OneDrive") + "/" + shortName;
                            docPath = docPath.Replace("/", "\\");
                        }
                    }
                    Process.Start(@"explorer.exe", "/select,\"" + docPath + "\"");
                }
            }
            catch (UriFormatException ex)
            {
                WordApp.StatusBar = ex.Message + "\r\n\r\n" + ex.Source;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "打开文档所在文件夹时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSelectObject_Click(object sender, RibbonControlEventArgs e)
        {
            try { WordApp.Run("SelectDrawingObjects"); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "选择对象",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}