using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        private void clearCustomStyles()
        {
            try
            {
                foreach (MSWord.Style sy in WordApp.ActiveDocument.Styles)
                    if (!sy.BuiltIn) sy.Delete();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "清理自定义样式时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void creatCustomStyle()
        {
            try
            {
                bool t4, t3, t2, t1, t, th, ts;
                t4 = t3 = t2 = t1 = t = th = ts = false;
                foreach (MSWord.Style name in WordApp.ActiveDocument.Styles)
                {
                    switch (name.NameLocal)
                    {
                        case "公文:4级标题": t4 = true; break;
                        case "公文:3级标题": t3 = true; break;
                        case "公文:2级标题": t2 = true; break;
                        case "公文:1级标题": t1 = true; break;
                        case "公文:标题": th = true; break;
                        case "公文:正文": t = true; break;
                        case "公文:正文(缩进)": ts = true; break;
                    }
                }
                if ((t4 & t3 & t2 & t1 & t & th & ts) == true)
                {
                    System.Diagnostics.Debug.WriteLine("公文样式已存在!");
                    return;
                }

                MSWord.ListTemplate myTemplate;
                MSWord.Style syT4, syT3, syT2, syT1, sy, syTh, sys;

                if (!t4)
                {
                    syT4 = WordApp.ActiveDocument.Styles.Add("公文:4级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT4.set_BaseStyle("");
                    syT4.Font.NameFarEast = iniTitle4.Name;
                    syT4.Font.Size = iniTitle4.SizeInPoints;
                    if (iniTitle4.Bold) syT4.Font.Bold = 1;
                    if (iniTitle4.Italic) syT4.Font.Italic = 1;
                    syT4.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel4;

                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberFormat = "（%1）";
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
                    myTemplate.ListLevels[1].TextPosition = 0;
                    syT4.LinkToListTemplate(myTemplate);
                    myTemplate = null;
                }

                if (!t3)
                {
                    syT3 = WordApp.ActiveDocument.Styles.Add("公文:3级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT3.set_BaseStyle("");
                    syT3.Font.NameFarEast = iniTitle3.Name;
                    syT3.Font.Size = iniTitle3.SizeInPoints;
                    if (iniTitle3.Bold) syT3.Font.Bold = 1;
                    if (iniTitle3.Italic) syT3.Font.Italic = 1;
                    syT3.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel3;

                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberFormat = "%1．";
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
                    myTemplate.ListLevels[1].TextPosition = 0;
                    syT3.LinkToListTemplate(myTemplate);
                    myTemplate = null;
                }

                if (!t2)
                {
                    syT2 = WordApp.ActiveDocument.Styles.Add("公文:2级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT2.set_BaseStyle("");
                    syT2.Font.NameFarEast = iniTitle2.Name;
                    syT2.Font.Size = iniTitle2.SizeInPoints;
                    if (iniTitle2.Bold) syT2.Font.Bold = 1;
                    if (iniTitle2.Italic) syT2.Font.Italic = 1;
                    syT2.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel2;

                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                    myTemplate.ListLevels[1].NumberFormat = "（%1）";
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
                    myTemplate.ListLevels[1].TextPosition = 0;
                    syT2.LinkToListTemplate(myTemplate);
                    myTemplate = null;
                }

                if (!t1)
                {
                    syT1 = WordApp.ActiveDocument.Styles.Add("公文:1级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT1.set_BaseStyle("");
                    syT1.Font.NameFarEast = iniTitle1.Name;
                    syT1.Font.Size = iniTitle1.SizeInPoints;
                    if (iniTitle1.Bold) syT1.Font.Bold = 1;
                    if (iniTitle1.Italic) syT1.Font.Italic = 1;
                    syT1.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel1;

                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                    myTemplate.ListLevels[1].NumberFormat = "%1、";
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
                    myTemplate.ListLevels[1].TextPosition = 0;
                    syT1.LinkToListTemplate(myTemplate);
                    myTemplate = null;
                }

                if (!t)
                {
                    sy = WordApp.ActiveDocument.Styles.Add("公文:正文", MSWord.WdStyleType.wdStyleTypeParagraph);
                    sy.set_BaseStyle("");
                    sy.Font.NameFarEast = iniNormal.Name;
                    sy.Font.Size = iniNormal.SizeInPoints;
                    if (iniNormal.Bold) sy.Font.Bold = 1;
                    if (iniNormal.Italic) sy.Font.Italic = 1;
                    sy.set_BaseStyle(MSWord.WdBuiltinStyle.wdStyleNormal);

                    int key = WordApp.BuildKeyCode(MSWord.WdKey.wdKeyAlt, MSWord.WdKey.wdKeyShift, MSWord.WdKey.wdKeyZ);
                    WordApp.KeyBindings.Add(MSWord.WdKeyCategory.wdKeyCategoryStyle, "公文:正文", key);
                }

                if (!ts)
                {
                    sys = WordApp.ActiveDocument.Styles.Add("公文:正文(缩进)", MSWord.WdStyleType.wdStyleTypeParagraph);
                    sys.set_BaseStyle("");
                    sys.Font.NameFarEast = iniNormal.Name;
                    sys.Font.Size = iniNormal.SizeInPoints;
                    if (iniNormal.Bold) sys.Font.Bold = 1;
                    if (iniNormal.Italic) sys.Font.Italic = 1;
                    sys.set_BaseStyle(MSWord.WdBuiltinStyle.wdStyleNormal);
                    sys.ParagraphFormat.CharacterUnitFirstLineIndent = 2;

                    int key = WordApp.BuildKeyCode(MSWord.WdKey.wdKeyAlt, MSWord.WdKey.wdKeyShift, MSWord.WdKey.wdKeyS);
                    WordApp.KeyBindings.Add(MSWord.WdKeyCategory.wdKeyCategoryStyle, "公文:正文(缩进)", key);
                }

                if (!th)
                {
                    syTh = WordApp.ActiveDocument.Styles.Add("公文:标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syTh.set_BaseStyle("");
                    syTh.Font.NameFarEast = iniTitle.Name;
                    syTh.Font.Size = iniTitle.SizeInPoints;
                    if (iniTitle.Bold) syTh.Font.Bold = 1;
                    if (iniTitle.Italic) syTh.Font.Italic = 1;
                    syTh.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    syTh.set_BaseStyle(MSWord.WdBuiltinStyle.wdStyleNormal);
                }
            }
            catch (Exception ex)
            {
                string err = ex.Message + "比如指定了不存在的字体或者在英文环境使用了!请到关于-高级设置里先指定公文字体。\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "新建公文样式时发生错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                creatCustomStyle();
                WordApp.TaskPanes[MSWord.WdTaskPanes.wdTaskPaneFormatting].Visible = true;
                WordApp.ShowStylePreviews = true;
            }
            finally { Cursor.Current = Cursors.Default; }
        }

        private void btnStyleR_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                WordApp.Selection.ClearFormatting();
                clearCustomStyles();
                creatCustomStyle();
                WordApp.TaskPanes[MSWord.WdTaskPanes.wdTaskPaneFormatting].Visible = true;
                WordApp.ShowStylePreviews = true;
            }
            finally { Cursor.Current = Cursors.Default; }
        }
    }
}