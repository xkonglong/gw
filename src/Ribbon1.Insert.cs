using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using MSWord = Microsoft.Office.Interop.Word;
using MSOffice = Microsoft.Office.Core;

namespace 公文排版
{
    public partial class Ribbon1
    {
        /// <summary>插入版记线</summary>
        private void InsertLine(int LineNumber = 1)
        {
            if (LineNumber < 1) LineNumber = 1;
            MSWord.Document WordDoc = WordApp.ActiveDocument;
            MSWord.Table newTable = WordDoc.Tables.Add(WordApp.Selection.Range, LineNumber, 1);

            newTable.Rows.RelativeVerticalPosition = MSWord.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
            newTable.Rows.VerticalPosition = (float)MSWord.WdFramePosition.wdFrameBottom;
            newTable.PreferredWidth = 443;
            newTable.Rows.Height = 28.5f;
            newTable.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowLeft;

            newTable.Borders[MSWord.WdBorderType.wdBorderTop].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            newTable.Borders[MSWord.WdBorderType.wdBorderTop].LineWidth = MSWord.WdLineWidth.wdLineWidth100pt;
            newTable.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            newTable.Borders[MSWord.WdBorderType.wdBorderBottom].LineWidth = MSWord.WdLineWidth.wdLineWidth100pt;
            if (LineNumber > 1)
            {
                newTable.Borders[MSWord.WdBorderType.wdBorderHorizontal].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                newTable.Borders[MSWord.WdBorderType.wdBorderHorizontal].LineWidth = MSWord.WdLineWidth.wdLineWidth075pt;
            }

            newTable.Select();
            WordApp.Selection.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            SetFont(iniFoot);
        }

        /// <summary>插入版头</summary>
        private void InsertHead(string autoText)
        {
            string startPath = GetDllPath();
            string dotFilename = Path.Combine(startPath ?? string.Empty, "Resources", "公文部件.dotx");
            if (!File.Exists(dotFilename))
            {
                MessageBox.Show("未找到公文部件,请重新安装/修复本插件!", "载入公文部件时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                WordApp.AddIns.Add(dotFilename, Install: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "加载公文部件时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            try
            {
                WordApp.Selection.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                MSWord.Template template = WordApp.Templates[dotFilename];
                if (template.BuildingBlockEntries.Count < 1) return;
                template.BuildingBlockEntries.Item(autoText).Insert(Where: WordApp.Selection.Range);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "插入版头时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnFootSingle_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入单行版记");
            InsertLine();
            undo.EndCustomRecord();
        }

        private void btnFootDouble_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入双行版记");
            InsertLine(2);
            undo.EndCustomRecord();
        }

        private void btnDate_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入日期");
            SetFont(iniNormal);
            try
            {
                WordApp.Selection.InsertDateTime("YYYY年M月D日", false);
                WordApp.Selection.InsertParagraphAfter();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { undo.EndCustomRecord(); }
        }

        private void btnLine1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                WordApp.ScreenUpdating = false;
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("插入间隔线");
                WordApp.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                WordApp.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineWidth = MSWord.WdLineWidth.wdLineWidth225pt;
                undo.EndCustomRecord();

                undo.StartCustomRecord("间隔线设为红色");
                WordApp.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].Color = MSWord.WdColor.wdColorRed;
                undo.EndCustomRecord();
                WordApp.ScreenUpdating = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnInsertPaper_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("插入横页");
                WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                WordApp.Selection.InsertBreak(MSWord.WdBreakType.wdSectionBreakNextPage);
                WordApp.Selection.PageSetup.Orientation = MSWord.WdOrientation.wdOrientLandscape;
                undo.EndCustomRecord();

                undo.StartCustomRecord("插入竖页");
                WordApp.Selection.InsertBreak(MSWord.WdBreakType.wdSectionBreakNextPage);
                WordApp.Selection.PageSetup.Orientation = MSWord.WdOrientation.wdOrientPortrait;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "插入横页失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void galSymbol_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string f = glySymbol.SelectedItem.Label;
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("插入符号:" + f);
                WordApp.Selection.TypeText(f);
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnSymbol_zkh_Click(object sender, RibbonControlEventArgs e)
        {
            WordReplace(true, @"\[([0-9]{1,4})\]", @"〔\1〕");
            WordReplace(true, "［([0-9]{1,4})］", @"〔\1〕");
        }

        private void btnSymbol_hkh_Click(object sender, RibbonControlEventArgs e)
        {
            WordReplace(true, "【([0-9]{1,4})】", @"〔\1〕");
        }

        private void BtnSymbol_index_Click(object sender, RibbonControlEventArgs e)
        {
            int at = 1;
            string number = WordApp.Selection.Text;

            if (!int.TryParse(number, out at))
            {
                MessageBox.Show("请选中一个 1-20 的数字", "数字转换为带圈序号",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (at > 20)
            {
                MessageBox.Show("带圈序号最大只能到20", "数字转换为带圈序号",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("带圈数字序列");
            MSWord.ListTemplate myTemplate2 = WordApp.ActiveDocument.ListTemplates.Add();
            myTemplate2.ListLevels[1].NumberFormat = "%1";
            myTemplate2.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleNumberInCircle;
            myTemplate2.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
            myTemplate2.ListLevels[1].TextPosition = 0;
            myTemplate2.ListLevels[1].NumberPosition = 0;
            if (at <= 0) at = 1;
            myTemplate2.ListLevels[1].StartAt = at;
            WordApp.Selection.Range.ListFormat.ApplyListTemplate(myTemplate2);
            undo.EndCustomRecord();
        }

        // ============ 公章 ============
        private void btnInsertSignPicLink_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("以链接方式插入印章");
                btnInsertSignPic.Tag = "Link";
                btnInsertSignPic.SuperTip = btnInsertSignPicLink.SuperTip;
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Title = "请选择以链接方式插入的印章";
                    ofd.Filter = "支持透明度的图片(*.PNG,*.gif,*.TIFF)|*.png;*.gif;*.tiff|所有文件(*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                        InsertSignPic(ofd.FileName);
                }
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "插入印章时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnInsertSignPicSave_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("以嵌入方式插入印章");
                btnInsertSignPic.Tag = "";
                btnInsertSignPic.SuperTip = btnInsertSignPicSave.SuperTip;
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Title = "请选择以保存方式插入的印章";
                    ofd.Filter = "支持透明度的图片(*.PNG,*.gif,*.TIFF)|*.png;*.gif;*.tiff|所有文件(*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                        InsertSignPic(ofd.FileName, false);
                }
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "插入印章时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnInsertSignPic_Click(object sender, RibbonControlEventArgs e)
        {
            if ((string)btnInsertSignPic.Tag == "Link") btnInsertSignPicLink_Click(sender, e);
            else btnInsertSignPicSave_Click(sender, e);
        }

        private void InsertSignPic(string FileName, bool LinkToFile = true)
        {
            MSWord.InlineShape pic = WordApp.Selection.Range.InlineShapes.AddPicture(FileName, LinkToFile, !LinkToFile);
            MSWord.Shape sPic = pic.ConvertToShape();
            sPic.ZOrder(MSOffice.MsoZOrderCmd.msoBringInFrontOfText);
        }

        // ============ 版头快捷按钮 ============
        private void btnHeadSingle_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:单行发文机关");
            InsertHead("版头:单行发文机关");
            undo.EndCustomRecord();
        }

        private void btnHeadMulti_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:联合发文机关");
            InsertHead("版头:联合发文机关");
            undo.EndCustomRecord();
        }

        private void BtnLineH1_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:函线(上沿)");
            InsertHead("版头:函线(上沿)");
            undo.EndCustomRecord();
        }

        private void BtnLineH2_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:函线(下沿)");
            InsertHead("版头:函线(下沿)");
            undo.EndCustomRecord();
        }

        private void BtnBindLine_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入装订线");
            InsertHead("装订线");
            undo.EndCustomRecord();
        }
    }
}