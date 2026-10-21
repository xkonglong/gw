using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace 公文排版
{
    public partial class FrmFontChangeDialog : Form
    {
        Microsoft.Office.Interop.Word.Application WordApp;
        public FrmFontChangeDialog()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 载入字体
        /// </summary>
        /// <param name="sender">指定的comboBox控件</param>
        private void LoadFont(object sender)
        {
            try
            {
                //载入字体
                ((ComboBox)sender).Items.Clear();
                System.Drawing.Text.InstalledFontCollection ifc = new InstalledFontCollection();
                FontFamily[] ffs = ifc.Families;
                //Font f;
                ((ComboBox)sender).Items.Clear();
                foreach (FontFamily ff in ffs)
                {
                                           
                        ((ComboBox)sender).Items.Add(ff.Name);
                    
                }
            }
            catch (Exception ex)
            {
               // Globals.ThisAddIn.Logging(ex.Source, ex.Message);
                Debug.WriteLine(ex.Message);
            }
        }
        /// <summary>
        /// 更改字母和数字字体的方法
        /// </summary>
        /// <param name="numLetter">字母还是数字,^#为任意数字,^$为任意字母</param>
        /// <param name="fontName">字体名称</param>
        /// <returns></returns>
        private bool ChangeFont(string numLetter , string fontName)
        {

            WordApp.Selection.Find.ClearFormatting();//清除查找格式
            WordApp.Selection.Find.Replacement.ClearFormatting(); //清除替换格式
            

            //开始替换
            WordApp.Selection.Find.Text = numLetter;

            WordApp.Selection.Find.Forward = true;
            WordApp.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            WordApp.Selection.Find.Format = true;
            WordApp.Selection.Find.MatchByte = true;
            WordApp.Selection.Find.MatchCase = false;
            WordApp.Selection.Find.MatchWholeWord = false;
            WordApp.Selection.Find.CorrectHangulEndings = false;
            WordApp.Selection.Find.HanjaPhoneticHangul = false;
            WordApp.Selection.Find.MatchWildcards = false;
            WordApp.Selection.Find.MatchSoundsLike = false;
            WordApp.Selection.Find.MatchAllWordForms = false;

            WordApp.Selection.Find.Replacement.Text = "^&"; //替换内容为查找内容
            WordApp.Selection.Find.Replacement.Font.Name = fontName;   //设置新的西文字体

            bool win=WordApp.Selection.Find.Execute(Replace:WdReplace.wdReplaceAll);


            WordApp.Selection.Find.ClearFormatting();//清除查找格式
            WordApp.Selection.Find.Replacement.ClearFormatting(); //清除替换格式
            return win;


            
        }







        //字体更改窗体载入时
        private void FrmFontChangeDialog_Load(object sender, EventArgs e)
        {
            WordApp = Globals.ThisAddIn.Application;
            LoadFont(cobLetter);
            LoadFont(cobNumber);
            LoadFont(cobFont2);
            LoadFont(cobFont1);

            string letterFont = Properties.Settings.Default.英文字体;
            if (letterFont.Length>0)
            {

                cobLetter.Text = letterFont;
            }
            string numberFont = Properties.Settings.Default.数字字体;
            if(numberFont.Length>0)
            {
                cobNumber.Text = numberFont;
            }
        }

        //勾选数字               
        private void ChkNumber_CheckedChanged(object sender, EventArgs e)
        {
            if (chkboxFont.Checked || chkNumber.Checked || chkLetter.Checked)
            {
                btnOK.Enabled = true;

            }
            else
            {
                btnOK.Enabled = false;
            }
            cobNumber.Enabled = chkNumber.Checked;
        }
        //勾选字母
        private void ChkLetter_CheckedChanged(object sender, EventArgs e)
        {
            if (chkboxFont.Checked || chkNumber.Checked || chkLetter.Checked)
            {
                btnOK.Enabled = true;

            }
            else
            {
                btnOK.Enabled = false;
            }
            cobLetter.Enabled = chkLetter.Checked;
        }

        //确定按钮
        private void BtnOK_Click(object sender, EventArgs e)
        {
            //  UndoRecord  不支持查找替换功能的记录, 否则可能导致word崩溃
            try
            {
                string w = "";
                if(chkboxFont.Checked)
                {
                    if (cobFont1.Text.Length == 0)
                    {
                        MessageBox.Show("请选择或者输入被替换字体的名称!", this.Text);
                        return;
                    }
                    else if (cobFont2.Text.Length==0)
                    {
                        MessageBox.Show("请选择或者输入要替换字体的名称!", this.Text);
                        return;
                    }
                    else
                    {
                        WordApp.Selection.Find.ClearFormatting();//清除查找格式
                        WordApp.Selection.Find.Replacement.ClearFormatting(); //清除替换格式

                        WordApp.Selection.Find.Text = "";
                        WordApp.Selection.Find.Replacement.Text = "";
                        //开始替换
                        WordApp.Selection.Find.Forward = true;
                        WordApp.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
                        WordApp.Selection.Find.Format = true;
                        WordApp.Selection.Find.MatchByte = true;
                        WordApp.Selection.Find.MatchCase = false;
                        WordApp.Selection.Find.MatchWholeWord = false;
                        WordApp.Selection.Find.CorrectHangulEndings = false;
                        WordApp.Selection.Find.HanjaPhoneticHangul = false;
                        WordApp.Selection.Find.MatchWildcards = false;
                        WordApp.Selection.Find.MatchSoundsLike = false;
                        WordApp.Selection.Find.MatchAllWordForms = false;

                        WordApp.Selection.Find.Font.Name = cobFont1.Text;  //查找的旧字体
                        WordApp.Selection.Find.Replacement.Font.Name = cobFont2.SelectedItem.ToString();   //替换的新字体

                        WordApp.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);

                        WordApp.Selection.Find.ClearFormatting();//清除查找格式
                        WordApp.Selection.Find.Replacement.ClearFormatting(); //清除替换格式
                        w = cobFont1.Text + "已改为:" + cobFont2.SelectedItem.ToString();
                    }
                }


                if(chkNumber.Checked)
                {
                    if ( cobNumber.Text == "")
                    {
                        MessageBox.Show("未指定数字字体!", this.Text);
                        return;
                    }
                    else
                    {
                        string f = cobNumber.SelectedItem.ToString();
                        ChangeFont("^#",f );   //更改数字字体
                        Properties.Settings.Default.数字字体 = f;
                        w += " 数字字体改为:" + f;
                    }
                }

                if (chkLetter.Checked)
                {
                    if (  cobLetter.Text == "")
                    {
                        MessageBox.Show("未指定字母字体!", this.Text);
                        return;
                    }
                    else
                    {
                        string f = cobLetter.SelectedItem.ToString();
                        ChangeFont("^$",f);   //更改字母字体
                        Properties.Settings.Default.英文字体 = f;
                        w += " 字母字体改为:" + f;
                    }

                }
                WordApp.StatusBar = w;
                Properties.Settings.Default.Save();
                this.Close();

            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "设置字体时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }



        }



        private void ChkboxFont_CheckedChanged(object sender, EventArgs e)
        {
            if (chkboxFont.Checked || chkNumber.Checked ||chkLetter.Checked)
            {
                btnOK.Enabled = true;
                
            }
            else
            {
                btnOK.Enabled = false;
            }
            cobFont1.Enabled = chkboxFont.Checked;
            cobFont2.Enabled = chkboxFont.Checked;
        }

        private void CobFont1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup(cobFont1, "选择要更改的字体,如果本机不含该字体,也可以直接输入该字体正确的名称.", hlpevent.MousePos);
        }



        private void FrmFontChangeDialog_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            Help.ShowPopup(this, "点击帮助按钮后,再点击不懂的地方或控件,可以获得帮助.", MousePosition);
        }

        private void CobFont2_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup((Control)sender, "在这里选择要替换的字体,你也可以输入字体的部分名称,助手会帮你找到类似字体.", hlpevent.MousePos);
        }

        private void CobNumber_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup((Control)sender, "在这里指定数字要使用的字体,你可以输入字体的部分名称,助手会帮你找到类似字体", hlpevent.MousePos);
        }

        private void CobLetter_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup((Control)sender, "在这里指定字母要使用的字体,你可以输入字体的部分名称,助手会帮你找到类似字体", hlpevent.MousePos);
        }

        private void ChkLetter_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup((Control)sender, "要更改字母字体,请勾选此处", hlpevent.MousePos);
        }

        private void ChkNumber_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup((Control)sender, "要更改数字字体,请勾选此处", hlpevent.MousePos);
        }

        private void ChkboxFont_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowPopup((Control)sender, "要替换指定的字体,请勾选此处", hlpevent.MousePos);
        }
    }


    }

