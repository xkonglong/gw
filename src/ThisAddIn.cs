using System;
using MSWord = Microsoft.Office.Interop.Word;
using MSOffice = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using System.Windows.Forms;
using System.Diagnostics;

namespace 公文排版
{
    
    public partial class ThisAddIn
    {
        Microsoft.Office.Interop.Word.Application WordApp;




        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;  //在活动窗口中的所选内容更改时发生。
        }

        private void Application_WindowSelectionChange(MSWord.Selection Sel)  //在活动窗口中的所选内容更改时发生。
        {
            //throw new NotImplementedException();
            try
            {
                //如果文档是保护状态,控件变灰
                if (Sel.Document.ProtectionType != MSWord.WdProtectionType.wdNoProtection)
                {
                    ContorlEnable(false);
                    Globals.Ribbons.Ribbon1.btnSpeak.Enabled = true;
                    Globals.Ribbons.Ribbon1.glyBak.Enabled = true;
                }
                else
                {
                    ContorlEnable();
                }
            }
            catch (Exception ex)
            {
                this.Application.StatusBar = ex.Message;
                throw;
            }
        }

        //private void Pane_VisibleChanged(object sender, EventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            WordApp = Globals.ThisAddIn.Application;
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

            this.Application.DocumentOpen += Application_DocumentOpen;

            ((MSWord.ApplicationEvents4_Event)this.Application).NewDocument += ThisAddIn_NewDocument;
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.ProtectedViewWindowActivate += Application_ProtectedViewWindowActivate;
            //this.Application.WindowBeforeRightClick += Application_WindowBeforeRightClick;

        }


        private void Application_ProtectedViewWindowActivate(MSWord.ProtectedViewWindow PvWindow)//文档保护视图激活时,按钮禁用
        {
            ContorlEnable(false);
            Globals.Ribbons.Ribbon1.btnSpeak.Enabled = true;
            Globals.Ribbons.Ribbon1.glyBak.Enabled = true;

        }

        private void Application_DocumentBeforeClose(MSWord.Document Doc, ref bool Cancel)//文档关闭时,按钮禁用
        {
            ContorlEnable(false);
        }

        private void Application_WindowActivate(MSWord.Document Doc, MSWord.Window Wn)//文档激活时,判断按钮是否禁用
        {
            try
            { 
                //如果文档是保护状态,控件变灰
                if (Doc.ProtectionType != MSWord.WdProtectionType.wdNoProtection)
                {
                    ContorlEnable(false);
                    Globals.Ribbons.Ribbon1.btnSpeak.Enabled = true;
                    Globals.Ribbons.Ribbon1.glyBak.Enabled = true;

                }
                else
                {
                    ContorlEnable();
                }
            }
            catch(Exception ex)
            {
                this.Application.StatusBar = "无法获取文档保护状态,请检查Office完整性.";
                throw;
            }
        }

 
        /// <summary>
        /// 功能区控件是否可用.比如当文档无法编辑时,控件将变灰.
        /// </summary>
        /// <param name="enable">bool值,默认为可用true.</param>
        internal void ContorlEnable(bool enable=true)
        {
            Ribbon1 rb = Globals.Ribbons.Ribbon1;
            rb.btnFoot.Enabled = enable;
            rb.btnDate.Enabled = enable;
            rb.btnHead.Enabled = enable;
            rb.btnHeadNoLine.Enabled = enable;
            rb.btnInsert2.Enabled = enable;   //首行缩进
            rb.btnInsertPaper.Enabled = enable;   //插入横页
            rb.btnLeft.Enabled = enable;
            rb.btnName.Enabled = enable;

            rb.glyBak.Enabled = enable;
            rb.glyFormat.Enabled = enable;
            rb.glyFoot.Enabled = enable;
            rb.glyFont.Enabled = enable;
            rb.glyPassword.Enabled = enable;
            rb.glySymbol.Enabled = enable;

            rb.btnNormal.Enabled = enable;
            rb.btnNum2Txt.Enabled = enable;
            rb.btnPaper.Enabled = enable;
            //rb.btnProtectDoc.Enabled = enable;
            rb.btnSpeak.Enabled = enable;
            rb.btnStyle.Enabled = enable;

            rb.mnuTitle1.Enabled = enable;
            rb.mnuTitle2.Enabled = enable;
            rb.mnuTitle3.Enabled = enable;
            rb.mnuTitle4.Enabled = enable;
            
            rb.btnTitleHei.Enabled = enable;
            rb.btnUnderDot.Enabled = enable;
            rb.buttonGroup1.Enabled = enable;
            rb.cobFonts.Enabled = enable;
            rb.btnHideFont.Enabled = enable;
            rb.btnInsertSignPic.Enabled = enable;
            rb.btnInsertSignPicLink.Enabled = enable;
            rb.btnInsertSignPicSave.Enabled = enable;

            rb.mnuInsertGWObject.Enabled = enable; //公文部件

            //实验室功能
            rb.mnuSelect.Enabled = enable;
            rb.mnuConvert.Enabled = enable;
            rb.btnAI.Enabled = enable;

        }

        private void ThisAddIn_NewDocument(MSWord.Document Doc)//新建文档时,激活按钮
        {
            ContorlEnable();
        }


        private void Application_DocumentOpen(MSWord.Document Doc)//文档打开时,判断按钮是否禁用
        {
            //如果文档是保护状态,控件变灰
            if (Doc.ProtectionType != MSWord.WdProtectionType.wdNoProtection )
            {
                ContorlEnable(false);
                Globals.Ribbons.Ribbon1.btnSpeak.Enabled = true;
                Globals.Ribbons.Ribbon1.glyBak.Enabled = true;
            }
            else
            {
                ContorlEnable();
            }
        }
        #endregion
    }
}
