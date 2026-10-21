using System;
using System.Drawing;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using MSOffice = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        /// <summary>设置 Word 所选文字字体(名称、大小、粗体、斜体)</summary>
        internal void SetFont(Font myFont)
        {
            if (myFont == null) return;
            try
            {
                WordApp.Selection.Font.Name = myFont.Name;
                WordApp.Selection.Font.Size = myFont.SizeInPoints;
                if (myFont.Bold) WordApp.Selection.Font.Bold = 1;
                if (myFont.Italic) WordApp.Selection.Font.Italic = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "设置字体时出错");
            }
        }

        /// <summary>设置标题编号</summary>
        private void SetTitle(object sender, int index = 1)
        {
            try
            {
                MSWord.ListTemplate myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
                myTemplate.ListLevels[1].TextPosition = 0;
                myTemplate.ListLevels[1].NumberPosition = 0;

                int at = Convert.ToInt32(((RibbonButton)sender).Tag);
                if (at <= 0) at = 1;
                myTemplate.ListLevels[1].StartAt = at;

                switch (index)
                {
                    case 1:
                        myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                        myTemplate.ListLevels[1].NumberFormat = "%1、";
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel1;
                        break;
                    case 2:
                        myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                        myTemplate.ListLevels[1].NumberFormat = "（%1）";
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel2;
                        break;
                    case 3:
                        myTemplate.ListLevels[1].NumberFormat = "%1．";
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel3;
                        break;
                    case 4:
                        myTemplate.ListLevels[1].NumberFormat = "（%1）";
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel4;
                        break;
                }
                WordApp.Selection.Range.ListFormat.ApplyListTemplate(myTemplate);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "设置标题格式时出错!",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void MnuTitle4_item_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为四级标题");
            SetFont(iniTitle4);
            SetTitle(sender, 4);
            undo.EndCustomRecord();
        }

        private void MnuTitle3_item_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为三级标题");
            SetFont(iniTitle3);
            SetTitle(sender, 3);
            undo.EndCustomRecord();
        }

        private void MnuTitle2_item_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为二级标题");
            SetFont(iniTitle2);
            SetTitle(sender, 2);
            undo.EndCustomRecord();
        }

        private void MnuTitle1_item_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为一级标题");
            SetFont(iniTitle1);
            SetTitle(sender, 1);
            undo.EndCustomRecord();
        }

        private void BtnNormal_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为公文:正文");
            SetFont(iniNormal);
            undo.EndCustomRecord();
            try
            {
                undo.StartCustomRecord("所选文字取消编号和大纲级别");
                WordApp.Selection.Range.ListFormat.RemoveNumbers(MSWord.WdNumberType.wdNumberAllNumbers);
                WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevelBodyText;
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnTitleHei_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为公文:标题");
            SetFont(iniTitle);
            undo.EndCustomRecord();
            try
            {
                undo.StartCustomRecord("设置标题居中");
                WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void cobFonts_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为二号" + cobFonts.Text);
            WordApp.Selection.Font.Name = cobFonts.Text;
            WordApp.Selection.Font.Size = 22;
            undo.EndCustomRecord();
            try
            {
                undo.StartCustomRecord("设置标题居中");
                WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnHead_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("设置所选字体");
            SetFont(iniHead);
            undo.EndCustomRecord();
        }

        private void btnName_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("设置所选字体");
            SetFont(iniName);
            undo.EndCustomRecord();
        }

        private void btnFoot_Click(object sender, RibbonControlEventArgs e)
        {
            var undo = WordApp.UndoRecord;
            undo.StartCustomRecord("设置所选字体");
            SetFont(iniFoot);
            undo.EndCustomRecord();
        }

        private void btnUnderDot_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("着重号");
                if (WordApp.Selection.Font.EmphasisMark != MSWord.WdEmphasisMark.wdEmphasisMarkUnderSolidCircle)
                    WordApp.Selection.Font.EmphasisMark = MSWord.WdEmphasisMark.wdEmphasisMarkUnderSolidCircle;
                else
                    WordApp.Selection.Font.EmphasisMark = MSWord.WdEmphasisMark.wdEmphasisMarkNone;
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnHideFont_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("隐藏");
                int hide = WordApp.Selection.Font.Hidden;
                WordApp.Selection.Font.Hidden = hide == -1 ? 0 : -1;
                undo.EndCustomRecord();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void BtnFontChangeDialog_Click(object sender, RibbonControlEventArgs e)
        {
            using (var f = new FrmFontChangeDialog())
                f.ShowDialog();
        }

        private void btnFontFSgb2312_Click(object sender, RibbonControlEventArgs e) => FontSwitch("仿宋_GB2312", "仿宋");
        private void btnFontKTgb2312_Click(object sender, RibbonControlEventArgs e) => FontSwitch("楷体_GB2312", "楷体");
        private void btnFontKT_Click(object sender, RibbonControlEventArgs e) => FontSwitch("楷体", "楷体_GB2312");
        private void btnFontFS_Click(object sender, RibbonControlEventArgs e) => FontSwitch("仿宋", "仿宋_GB2312");

        private void btnFontBS2HT_Click(object sender, RibbonControlEventArgs e)
        {
            FontSwitch("方正小标宋_GBK", "黑体");
            FontSwitch("方正小标宋简体", "黑体");
        }

        private void btnFontHT2BS_Click(object sender, RibbonControlEventArgs e) => FontSwitch("黑体", "方正小标宋_GBK");

        private void FontSwitch(string oldFont, string newFont)
        {
            try
            {
                WordApp.Selection.Find.ClearFormatting();
                WordApp.Selection.Find.Font.NameFarEast = oldFont;
                WordApp.Selection.Find.Replacement.ClearFormatting();
                WordApp.Selection.Find.Replacement.Font.NameFarEast = newFont;
                WordApp.Selection.Find.Execute("", false, false, false, false, false, true,
                    MSWord.WdFindWrap.wdFindContinue, true, "", MSWord.WdReplace.wdReplaceAll);
                WordApp.StatusBar = String.Format("字体{0}已替换为{1}.", oldFont, newFont);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "替换字体时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}