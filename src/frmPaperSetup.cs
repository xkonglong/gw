using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;

namespace 公文排版
{
    public partial class frmPaperSetup : Form
    {
        readonly Microsoft.Office.Interop.Word.Document wordDoc = Globals.ThisAddIn.Application.ActiveDocument;
        readonly Microsoft.Office.Interop.Word.Application WordApp = Globals.ThisAddIn.Application;

        /////////////////////////////////////////////////////////////////////////
        /// 
        ///           函数功能区
        /// 
        ////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// 读取设置
        /// </summary>
        private void ReadSetting()
        {
            nudTop.Value = (decimal)Properties.Settings.Default.上边距; //v3.6 起改为支持小数
            nudBottom.Value = (decimal)Properties.Settings.Default.下边距;
            nudLeft.Value = (decimal)Properties.Settings.Default.左边距;
            nudRight.Value = (decimal)Properties.Settings.Default.右边距;
            nudFoot.Value = (decimal)Properties.Settings.Default.页脚边距;  //24.7 有小数

            nudLine.Value = Properties.Settings.Default.行数;
            nudLetters.Value = Properties.Settings.Default.字数;


            chkGW.Checked = Properties.Settings.Default.公文版式; //是否开启公文版式选项

            chkLineSpacing.Checked = Properties.Settings.Default.允许行间距;//是否允许设置行间距
            nudLineSpacing.Value = (decimal)Properties.Settings.Default.行间距; //小数
        }

        /// <summary>
        /// 写入设置
        /// </summary>
        private void WriteSetting()
        {
            //v3.6 起页边距改为支持小数
            this.Cursor = Cursors.WaitCursor;
            Properties.Settings.Default.上边距 = (float)nudTop.Value;
            Properties.Settings.Default.下边距 = (float)nudBottom.Value;
            Properties.Settings.Default.左边距 = (float)nudLeft.Value;
            Properties.Settings.Default.右边距 = (float)nudRight.Value;
            Properties.Settings.Default.行数 = (int)nudLine.Value;
            Properties.Settings.Default.字数 = (int)nudLetters.Value;
            Properties.Settings.Default.页脚边距 = (float)nudFoot.Value;
            Properties.Settings.Default.行间距 = (float)nudLineSpacing.Value;
            Properties.Settings.Default.允许行间距 = chkLineSpacing.Checked;  //是否允许为固定行间距
            Properties.Settings.Default.公文版式 = chkGW.Checked;  //是否开启公文板式
            Properties.Settings.Default.Save();
            this.Cursor = Cursors.Default;
        }


        /// <summary>
        /// /公文页面设置方法
        /// </summary>
        void SetGWPage()
        {

            //获取页边距数值
            float top = WordApp.CentimetersToPoints((float)nudTop.Value / 10);
            float bottom = WordApp.CentimetersToPoints((float)nudBottom.Value / 10);
            float left = WordApp.CentimetersToPoints((float)nudLeft.Value / 10);
            float right = WordApp.CentimetersToPoints((float)nudRight.Value / 10);
            float foot = WordApp.CentimetersToPoints((float)nudFoot.Value / 10);
            //获取行数和字数值
            float letter = (float)nudLetters.Value;
            float line = (float)nudLine.Value;
            float spacing = 0;



            //设置中文和西文默认字体
            try
            {
                wordDoc.Styles[-1].Font.NameFarEast = Properties.Settings.Default.正文.Name;

                string letterFont = Properties.Settings.Default.英文字体;
                //wordDoc.Styles[-1].Font.NameAscii = letterFont;    //这个只能指定字母
                wordDoc.Styles[-1].Font.Name= letterFont;   //这个可以指定西文字体(字母+数字)
            }
            catch (Exception ex)
            {
                wordDoc.Styles[-1].Font.NameFarEast = "+中文正文";
                wordDoc.Styles[-1].Font.Name = "++西文正文";
                Debug.WriteLine(ex.Message);
            }





            if (chkGW.Checked) //如果允许公文精准版式
            {

                //设置行数和字数前必须设置中文字体为16     //wdStyleNormal	-1	 正常。
                //wordDoc.Styles[-1].Font.NameFarEast = "+中文正文";
                wordDoc.Styles[-1].Font.Size = 16;
                   
            }
            if (chkLineSpacing.Checked)    //如果允许固定行间距
            {
                spacing = (float)nudLineSpacing.Value;  //获取行间距的值
            }


            ////设置上下左右边距
            //wordDoc.PageSetup.TopMargin = top;
            //wordDoc.PageSetup.BottomMargin = bottom;
            //wordDoc.PageSetup.LeftMargin = left;
            //wordDoc.PageSetup.RightMargin = right;
            //wordDoc.PageSetup.FooterDistance = foot;
            ////wordDoc.PageSetup.Gutter = 0 //额外页边距

            foreach (Section s in wordDoc.Sections)
            {
                s.PageSetup.TopMargin = top;
                s.PageSetup.BottomMargin = bottom;
                s.PageSetup.LeftMargin = left;
                s.PageSetup.RightMargin = right;
                s.PageSetup.FooterDistance = foot;

                if (s.PageSetup.Orientation == WdOrientation.wdOrientPortrait)//如果页面竖向
                {
                    //设置竖向A4页面
                    s.PageSetup.PageWidth = WordApp.CentimetersToPoints(21);
                    s.PageSetup.PageHeight = WordApp.CentimetersToPoints(29.7f);
                   

                    if (chkGW.Checked) //如果允许公文精准版式
                    {
                        s.PageSetup.LayoutMode = WdLayoutMode.wdLayoutModeGrid;
                        s.PageSetup.CharsLine = letter;
                        s.PageSetup.LinesPage = line;
                            
                    }
                    if (chkLineSpacing.Checked)
                    {
                        s.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;  //固定行间距
                        s.Range.ParagraphFormat.LineSpacing = spacing; //设置行间距
                    }
                }
                else
                {
                    //设置横向A4, 通常为表格 行间距,行数字数不做调整
                    s.PageSetup.PageWidth = WordApp.CentimetersToPoints(29.7f);
                    s.PageSetup.PageHeight = WordApp.CentimetersToPoints(21);
                }

            }
  
        }


        /// <summary>
        /// 仅嵌入文档中所用的字符. 比如使用了小标宋标题,只保存这个标题用到的字符.减少文档尺寸,又保证了兼容性.
        /// </summary>
        /// <param name="i">bool值,为真嵌入字体</param>
        private void IncludFonts(bool i)
        {
            try
            {
                WordApp.ActiveDocument.EmbedTrueTypeFonts = i;
                WordApp.ActiveDocument.SaveSubsetFonts = i;
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "设置保存嵌入字体时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        /// <summary>
        /// 显示隐藏文字,但是不打印.
        /// </summary>
        /// <param name="i">bool值,是否显示隐藏文字,粗暴设为不打印.</param>
        private void ShowHideFont(bool i)
        {
            try
            {
                WordApp.ActiveWindow.View.ShowHiddenText = i;
                WordApp.Options.PrintHiddenText = false;
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "设置显示隐藏字体时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }





        /////////////////////////////////////////////////////////////////////////
        /// 
        ///           事件区
        /// 
        ////////////////////////////////////////////////////////////////////////




        public frmPaperSetup()
        {
            InitializeComponent();
        }



        private void FrmPaperSetup_Load(object sender, EventArgs e)// 设置窗口载入时读取配置文件
        {
            ReadSetting();
            nudLineSpacing.Enabled = chkLineSpacing.Checked;
            nudLetters.Enabled = chkGW.Checked;
            nudLine.Enabled = chkGW.Checked;
        }

        private void BtnReset_Click(object sender, EventArgs e)//重置按钮
        {
            try
            {
                Properties.Settings.Default.Reset();
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "重置设置时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            ReadSetting();

        }


        private void BtnApply_Click(object sender, EventArgs e) //应用页面设置
        {
            string tip = "";
            try
            {
                Microsoft.Office.Interop.Word.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("公文页面设置");
                SetGWPage();
                undo.EndCustomRecord();

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(ex.Message, "设置页面时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //是否嵌入字体字符
            IncludFonts(chkIncludeFonts.Checked);
            try
            {
                //保存配置
                WriteSetting();
            }
            catch (Exception ex)
            {
                tip += ex.Message + "\r\n" +"请重置 Word 默认模板后再尝试!";
                MessageBox.Show(tip, "保存设置时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }


        }

        private void LinkLabel1_Click(object sender, EventArgs e) //打开页面设置对话框
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogFileDocumentLayout].Show();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开页面设置对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        
        private void FrmPaperSetup_KeyPress(object sender, KeyPressEventArgs e)//页面设置对话框,按下Esc关闭.
        {
            if (e.KeyChar ==(char) Keys.Escape)
            {
                this.Close();
            }
        }
       
        private void BtnOK_Click(object sender, EventArgs e) //页面设置对话框,确定按钮
        {
            
            Cursor.Current = Cursors.WaitCursor;
            BtnApply_Click(sender,e); //应用设置


            WriteSetting(); //保存配置

            Cursor.Current = Cursors.Default;
            this.Close();  //关闭窗体

        }
        
        private void ChkShowHideFont_CheckedChanged(object sender, EventArgs e)//checkbox切换显示隐藏文字.
        {
            ShowHideFont(chkShowHideFont.Checked);
        }

        private void ChkLineSpacing_CheckedChanged(object sender, EventArgs e) //切换是否允许调整行间距
        {
            nudLineSpacing.Enabled = chkLineSpacing.Checked;
        }

        private void FrmPaperSetup_HelpButtonClicked(object sender, System.ComponentModel.CancelEventArgs e)//标题栏帮助按钮
        {
            try
            {
                Process.Start(Properties.Settings.Default.帮助_页面设置);
            }
            catch (System.ComponentModel.Win32Exception w)
            {

                MessageBox.Show(w.Message, "打开在线帮助时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开在线帮助时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //Help.ShowPopup((Control)sender, "点击帮助按钮后,再点击不懂的地方或控件,可以获得帮助.", MousePosition);


        }

        
        private void FrmPaperSetup_FormClosed(object sender, FormClosedEventArgs e)   //窗体关闭后,页面设置按钮恢复为未按下状态
        {
            
            Globals.Ribbons.Ribbon1.btnPaper.Checked = false;
            this.Dispose();
        }

        private void ChkGW_CheckedChanged(object sender, EventArgs e)
        {
            nudLine.Enabled = chkGW.Checked;
            nudLetters.Enabled = chkGW.Checked;
        }

 

        private void ShowHiddenText(object sender, EventArgs e)  //调用显示隐藏文字选项
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogToolsOptionsView].Show();
            }
            catch(Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
                throw;

            }
        }
        private void SaveFontInDoc(object sender, EventArgs e) //调用嵌入字体选项
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogToolsOptionsSave].Show();
            }
            catch(Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
            }
        }
        private void ShowWordGW(object sender, EventArgs e) //调用页面布局选项对话框
        {
            try
            {
                WordApp.Dialogs[WdWordDialog.wdDialogFilePageSetup].DefaultTab = WdWordDialogTab.wdDialogFilePageSetupTabLayout;
                WordApp.Dialogs[WdWordDialog.wdDialogFilePageSetup].Show();
            }
            catch(Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
            }
        }
 






    }
}
