using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace 公文排版
{
    public partial class frmPaperSetup : Form
    {
        private readonly Microsoft.Office.Interop.Word.Application WordApp = Globals.ThisAddIn.Application;

        private Microsoft.Office.Interop.Word.Document ActiveDoc
        {
            get { return WordApp.ActiveDocument; }
        }

        #region 函数区

        private void ReadSetting()
        {
            nudTop.Value = (decimal)Properties.Settings.Default.上边距;
            nudBottom.Value = (decimal)Properties.Settings.Default.下边距;
            nudLeft.Value = (decimal)Properties.Settings.Default.左边距;
            nudRight.Value = (decimal)Properties.Settings.Default.右边距;
            nudFoot.Value = (decimal)Properties.Settings.Default.页脚边距;

            nudLine.Value = Properties.Settings.Default.行数;
            nudLetters.Value = Properties.Settings.Default.字数;

            chkGW.Checked = Properties.Settings.Default.公文版式;
            chkLineSpacing.Checked = Properties.Settings.Default.允许行间距;
            nudLineSpacing.Value = (decimal)Properties.Settings.Default.行间距;
        }

        private void WriteSetting()
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {
                Properties.Settings.Default.上边距 = (float)nudTop.Value;
                Properties.Settings.Default.下边距 = (float)nudBottom.Value;
                Properties.Settings.Default.左边距 = (float)nudLeft.Value;
                Properties.Settings.Default.右边距 = (float)nudRight.Value;
                Properties.Settings.Default.行数 = (int)nudLine.Value;
                Properties.Settings.Default.字数 = (int)nudLetters.Value;
                Properties.Settings.Default.页脚边距 = (float)nudFoot.Value;
                Properties.Settings.Default.行间距 = (float)nudLineSpacing.Value;
                Properties.Settings.Default.允许行间距 = chkLineSpacing.Checked;
                Properties.Settings.Default.公文版式 = chkGW.Checked;
                Properties.Settings.Default.Save();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        void SetGWPage()
        {
            var doc = ActiveDoc;
            if (doc == null) throw new InvalidOperationException("没有打开的文档");

            float top = WordApp.CentimetersToPoints((float)nudTop.Value / 10);
            float bottom = WordApp.CentimetersToPoints((float)nudBottom.Value / 10);
            float left = WordApp.CentimetersToPoints((float)nudLeft.Value / 10);
            float right = WordApp.CentimetersToPoints((float)nudRight.Value / 10);
            float foot = WordApp.CentimetersToPoints((float)nudFoot.Value / 10);
            float letter = (float)nudLetters.Value;
            float line = (float)nudLine.Value;
            float spacing = 0;

            try
            {
                doc.Styles[-1].Font.NameFarEast = Properties.Settings.Default.正文.Name;
                string letterFont = Properties.Settings.Default.英文字体;
                if (!string.IsNullOrEmpty(letterFont))
                    doc.Styles[-1].Font.Name = letterFont;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("默认字体设置失败: " + ex.Message);
            }

            if (chkGW.Checked)
                doc.Styles[-1].Font.Size = 16;

            if (chkLineSpacing.Checked)
                spacing = (float)nudLineSpacing.Value;

            foreach (Section s in doc.Sections)
            {
                s.PageSetup.TopMargin = top;
                s.PageSetup.BottomMargin = bottom;
                s.PageSetup.LeftMargin = left;
                s.PageSetup.RightMargin = right;
                s.PageSetup.FooterDistance = foot;

                if (s.PageSetup.Orientation == WdOrientation.wdOrientPortrait)
                {
                    s.PageSetup.PageWidth = WordApp.CentimetersToPoints(21);
                    s.PageSetup.PageHeight = WordApp.CentimetersToPoints(29.7f);

                    if (chkGW.Checked)
                    {
                        s.PageSetup.LayoutMode = WdLayoutMode.wdLayoutModeGrid;
                        s.PageSetup.CharsLine = letter;
                        s.PageSetup.LinesPage = line;
                    }
                    if (chkLineSpacing.Checked)
                    {
                        s.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                        s.Range.ParagraphFormat.LineSpacing = spacing;
                    }
                }
                else
                {
                    s.PageSetup.PageWidth = WordApp.CentimetersToPoints(29.7f);
                    s.PageSetup.PageHeight = WordApp.CentimetersToPoints(21);
                }
            }
        }

        private void IncludFonts(bool embed)
        {
            try
            {
                var doc = ActiveDoc;
                if (doc == null) return;
                doc.EmbedTrueTypeFonts = embed;
                doc.SaveSubsetFonts = embed;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "设置保存嵌入字体时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void ShowHideFont(bool show)
        {
            try
            {
                WordApp.ActiveWindow.View.ShowHiddenText = show;
                WordApp.Options.PrintHiddenText = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "设置显示隐藏字体时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        #endregion

        #region 事件区

        public frmPaperSetup()
        {
            InitializeComponent();
        }

        private void FrmPaperSetup_Load(object sender, EventArgs e)
        {
            ReadSetting();
            nudLineSpacing.Enabled = chkLineSpacing.Checked;
            nudLetters.Enabled = chkGW.Checked;
            nudLine.Enabled = chkGW.Checked;
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定恢复默认页面设置吗？", "确认",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                Properties.Settings.Default.Reset();
                Properties.Settings.Default.Reload();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "重置设置时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            ReadSetting();
        }

        private bool ApplySettings()
        {
            try
            {
                Microsoft.Office.Interop.Word.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("公文页面设置");
                try { SetGWPage(); }
                finally { undo.EndCustomRecord(); }
                return true;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(ex.Message, "设置页面时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "设置页面时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (!ApplySettings()) return;
            IncludFonts(chkIncludeFonts.Checked);
            try
            {
                WriteSetting();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n请重置 Word 默认模板后再尝试!",
                    "保存设置时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void LinkLabel1_Click(object sender, EventArgs e)
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogFileDocumentLayout].Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "打开页面设置对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void FrmPaperSetup_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Escape) this.Close();
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                // ApplySettings 内部已完成 WriteSetting
                if (ApplySettings())
                    this.Close();
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void ChkShowHideFont_CheckedChanged(object sender, EventArgs e)
        {
            ShowHideFont(chkShowHideFont.Checked);
        }

        private void ChkLineSpacing_CheckedChanged(object sender, EventArgs e)
        {
            nudLineSpacing.Enabled = chkLineSpacing.Checked;
        }

        private void FrmPaperSetup_HelpButtonClicked(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                Process.Start(Properties.Settings.Default.帮助_页面设置);
            }
            catch (System.ComponentModel.Win32Exception w)
            {
                MessageBox.Show(w.Message, "打开在线帮助时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "打开在线帮助时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void FrmPaperSetup_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Globals.Ribbons.Ribbon1 != null)
                Globals.Ribbons.Ribbon1.btnPaper.Checked = false;
            this.Dispose();
        }

        private void ChkGW_CheckedChanged(object sender, EventArgs e)
        {
            nudLine.Enabled = chkGW.Checked;
            nudLetters.Enabled = chkGW.Checked;
        }

        private void ShowHiddenText(object sender, EventArgs e)
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogToolsOptionsView].Show();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ShowHiddenText: " + ex.Message);
                WordApp.StatusBar = ex.Message;
            }
        }

        private void SaveFontInDoc(object sender, EventArgs e)
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogToolsOptionsSave].Show();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("SaveFontInDoc: " + ex.Message);
            }
        }

        private void ShowWordGW(object sender, EventArgs e)
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogFilePageSetup].DefaultTab =
                    WdWordDialogTab.wdDialogFilePageSetupTabLayout;
                WordApp.Dialogs[WdWordDialog.wdDialogFilePageSetup].Show();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ShowWordGW: " + ex.Message);
            }
        }

        #endregion
    }
}