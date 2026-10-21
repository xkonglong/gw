using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using MSWord = Microsoft.Office.Interop.Word;

namespace 公文排版
{
    public partial class Ribbon1
    {
        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            clearVBA();
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp.Selection.TypeText("===========所有内置对话框 ===========\r\n");
            ShwoAllDialog();
            WordApp.Selection.TypeText("===========所有任务面板 ===========\r\n");
            ShowAllPanes();
        }

        private void ShowAllNumber()
        {
            MSWord.ListTemplate myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
            myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;
            myTemplate.ListLevels[1].TextPosition = 0;
            myTemplate.ListLevels[1].NumberPosition = 0;
            myTemplate.ListLevels[1].StartAt = 1;

            foreach (int myCode in Enum.GetValues(typeof(MSWord.WdListNumberStyle)))
            {
                string strName = Enum.GetName(typeof(MSWord.WdListNumberStyle), myCode);
                WordApp.Selection.TypeText(strName + "   " + myCode + "\r\n");
                try
                {
                    myTemplate.ListLevels[1].NumberStyle = (MSWord.WdListNumberStyle)myCode;
                    WordApp.Selection.Range.ListFormat.ApplyListTemplate(myTemplate);
                }
                catch { }
                WordApp.ActiveDocument.Content.ListFormat.ConvertNumbersToText();
                WordApp.Selection.TypeText("\r\n");
            }
        }

        private void ShwoAllDialog()
        {
            foreach (int myCode in Enum.GetValues(typeof(MSWord.WdWordDialog)))
            {
                string strName = Enum.GetName(typeof(MSWord.WdWordDialog), myCode);
                WordApp.Selection.TypeText(strName + "   " + myCode + "\r\n");
            }
        }

        private void ShowAllPanes()
        {
            foreach (int myCode in Enum.GetValues(typeof(MSWord.WdTaskPanes)))
            {
                string strName = Enum.GetName(typeof(MSWord.WdTaskPanes), myCode);
                WordApp.Selection.TypeText(strName + "   " + myCode + "\r\n");
            }
        }

        private void group5_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (IsWps) WordApp.ListCommands(false);
                else WordApp.Dialogs[MSWord.WdWordDialog.wdDialogListCommands].Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "打开Word内置对话框时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void group6_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try { WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFileSummaryInfo].Show(); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "打开Word内置对话框时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}