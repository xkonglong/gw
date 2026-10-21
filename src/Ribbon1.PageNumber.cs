using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using MSWord = Microsoft.Office.Interop.Word;

namespace 公文排版
{
    public partial class Ribbon1
    {
        /// <summary>设置页码位置: 0 居中, 1 外侧, 2 居右</summary>
        private void PageNumberAligment(int position)
        {
            try
            {
                if (WordApp.ActiveWindow.ActivePane.View.Type == MSWord.WdViewType.wdNormalView ||
                    WordApp.ActiveWindow.ActivePane.View.Type == MSWord.WdViewType.wdOutlineView)
                {
                    WordApp.ActiveWindow.ActivePane.View.Type = MSWord.WdViewType.wdPrintView;
                }
                WordApp.ScreenUpdating = false;

                switch (position)
                {
                    case 0:
                        foreach (MSWord.Section s in WordApp.ActiveDocument.Sections)
                        {
                            s.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageFooter;
                        InsertNumber();
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                        break;

                    case 1:
                        if (WordApp.ActiveWindow.Panes[1].Pages.Count <= 1)
                        {
                            MessageBox.Show("文档仅为1页时,无法设置外侧页码.", "插入页码",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (IsWps)
                        {
                            WordApp.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = -1;
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekPrimaryFooter;
                            WordApp.Selection.ParagraphFormat.CharacterUnitRightIndent = 1;
                            WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                            InsertNumber();
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekEvenPagesFooter;
                            WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 1;
                            WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                            InsertNumber();
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                        }
                        else
                        {
                            foreach (MSWord.Section s in WordApp.ActiveDocument.Sections)
                            {
                                s.PageSetup.OddAndEvenPagesHeaderFooter = -1;
                                MSWord.Range r;

                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                r.Font.Size = (float)Math.Round(iniPageNumber.SizeInPoints, 1);
                                r.Font.Name = iniPageNumber.Name;
                                if (iniPageNumber.Bold) r.Font.Bold = 1;
                                if (iniPageNumber.Italic) r.Font.Italic = 1;
                                r.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                                r.Text = "— ";
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Fields.Add(r, MSWord.WdFieldType.wdFieldEmpty, "PAGE \\* Arabic");
                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Text = " —";
                                r.ParagraphFormat.CharacterUnitRightIndent = 1;

                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                                r.Font.Size = (float)Math.Round(iniPageNumber.SizeInPoints, 1);
                                r.Font.Name = iniPageNumber.Name;
                                if (iniPageNumber.Bold) r.Font.Bold = 1;
                                if (iniPageNumber.Italic) r.Font.Italic = 1;
                                r.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                                r.Text = "— ";
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Fields.Add(r, MSWord.WdFieldType.wdFieldEmpty, "PAGE \\* Arabic");
                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Text = " —";
                                r.ParagraphFormat.CharacterUnitFirstLineIndent = 1;
                            }
                        }
                        break;

                    case 2:
                        foreach (MSWord.Section s in WordApp.ActiveDocument.Sections)
                        {
                            s.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                        }
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageFooter;
                        InsertNumber();
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "PageNumberAligment 方法出错");
            }
            finally { WordApp.ScreenUpdating = true; }
        }

        /// <summary>插入页码,格式为 — 1 —</summary>
        private void InsertNumber()
        {
            try
            {
                WordApp.Selection.Font.Size = (float)Math.Round(iniPageNumber.SizeInPoints, 1);
                WordApp.Selection.Font.Name = iniPageNumber.Name;
                if (iniPageNumber.Bold) WordApp.Selection.Font.Bold = 1;
                if (iniPageNumber.Italic) WordApp.Selection.Font.Italic = 1;
                WordApp.Selection.TypeText("— ");
                WordApp.Selection.Fields.Add(WordApp.Selection.Range, MSWord.WdFieldType.wdFieldEmpty, "PAGE \\* Arabic");
                WordApp.Selection.TypeText(" —");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "InsertNumber 方法出错");
            }
        }

        private void btnNumberCenter_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入居中页码");
            PageNumberAligment(0);
            undo.EndCustomRecord();
        }

        private void BtnNumberRight_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入右侧页码");
            PageNumberAligment(2);
            undo.EndCustomRecord();
        }

        private void btnNumberOutside_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入外侧页码");
            PageNumberAligment(1);
            undo.EndCustomRecord();
        }

        private void BtnNumberOnLocation_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("光标处插入页码");
            InsertNumber();
            undo.EndCustomRecord();
        }
    }
}