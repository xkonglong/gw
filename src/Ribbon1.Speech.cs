using System;
using System.Diagnostics;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        private void Sp_SpeakCompleted(object sender, System.Speech.Synthesis.SpeakCompletedEventArgs e)
        {
            try
            {
                this.btnSpeak.OfficeImageId = "SpeakCells";
                this.btnSpeak.Tag = "";
                this.btnSpeak.Label = "朗读";
                WordApp.StatusBar = "朗读完成!";
                this.btnSpeak.Checked = false;
                this.btnSpeak.SuperTip = "单击朗读,再次单击停止.";
            }
            catch (Exception ex) { Debug.WriteLine(ex.Message); }
        }

        private void BtnSpeakOption_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var clsWinWrap = new ClsWinWrap(WordApp.ActiveWindow.Hwnd);
                clsWinWrap.Show(new frmSpeakOption());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "打开朗读选项时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSpeak_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (sp == null) return;
                if (this.btnSpeak.Tag == null) this.btnSpeak.Tag = "";

                if (this.btnSpeak.Tag.ToString().Length == 0)
                {
                    this.btnSpeak.Tag = "playing";
                    this.btnSpeak.Label = "停止";
                    this.btnSpeak.Checked = true;
                    this.btnSpeak.OfficeImageId = "SpeakStop";
                    this.btnSpeak.SuperTip = "正在朗读中,单击停止.";

                    if (WordApp.Selection.Text.Length > 1)
                    {
                        sp.SpeakAsync(WordApp.Selection.Text);
                    }
                    else
                    {
                        WordApp.Selection.MoveEnd(MSWord.WdUnits.wdParagraph);
                        sp.SpeakAsync(WordApp.Selection.Text);
                    }
                    return;
                }

                sp.SpeakAsyncCancelAll();
                this.btnSpeak.Tag = "";
                this.btnSpeak.Label = "朗读";
                this.btnSpeak.Checked = false;
                this.btnSpeak.OfficeImageId = "SpeakCells";
                this.btnSpeak.SuperTip = "单击朗读,再次单击停止.";
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
    }
}