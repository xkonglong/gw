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
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        /// <summary>
        /// 主处理器：只负责根据文档保护状态启用/禁用功能区控件。
        /// Ribbon 自身状态（着重号、隐藏文字、选定内容按钮）由 Ribbon1.WordApp_WindowSelectionChange 负责。
        /// </summary>
        private void Application_WindowSelectionChange(MSWord.Selection Sel)
        {
            try
            {
                if (Sel == null) return;
                bool editable = Sel.Document.ProtectionType == MSWord.WdProtectionType.wdNoProtection;
                ControlEnable(editable);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Application_WindowSelectionChange: " + ex.Message);
                try { this.Application.StatusBar = ex.Message; } catch { }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
                this.Application.DocumentOpen -= Application_DocumentOpen;
                this.Application.WindowActivate -= Application_WindowActivate;
                this.Application.DocumentBeforeClose -= Application_DocumentBeforeClose;
                this.Application.ProtectedViewWindowActivate -= Application_ProtectedViewWindowActivate;
                ((MSWord.ApplicationEvents4_Event)this.Application).NewDocument -= ThisAddIn_NewDocument;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Shutdown: " + ex.Message);
            }
        }

        #region VSTO 生成的代码

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

            this.Application.DocumentOpen += Application_DocumentOpen;
            ((MSWord.ApplicationEvents4_Event)this.Application).NewDocument += ThisAddIn_NewDocument;
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.ProtectedViewWindowActivate += Application_ProtectedViewWindowActivate;
        }

        private void Application_ProtectedViewWindowActivate(MSWord.ProtectedViewWindow PvWindow)
        {
            try
            {
                ControlEnable(false);
                if (Globals.Ribbons.Ribbon1 != null)
                {
                    Globals.Ribbons.Ribbon1.btnSpeak.Enabled = true;
                    Globals.Ribbons.Ribbon1.glyBak.Enabled = true;
                }
            }
            catch (Exception ex) { Debug.WriteLine(ex.Message); }
        }

        private void Application_DocumentBeforeClose(MSWord.Document Doc, ref bool Cancel)
        {
            try
            {
                // 关闭时不必强制禁用；如果有其他文档激活，WindowActivate 会重新判断
                if (this.Application.Documents.Count <= 1)
                    ControlEnable(false);
            }
            catch (Exception ex) { Debug.WriteLine(ex.Message); }
        }

        private void Application_WindowActivate(MSWord.Document Doc, MSWord.Window Wn)
        {
            try
            {
                if (Doc == null) return;
                bool editable = Doc.ProtectionType == MSWord.WdProtectionType.wdNoProtection;
                ControlEnable(editable);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("WindowActivate: " + ex.Message);
                try { this.Application.StatusBar = "无法获取文档保护状态,请检查Office完整性."; } catch { }
            }
        }

        /// <summary>
        /// 功能区控件是否可用.
        /// </summary>
        internal void ControlEnable(bool enable = true)
        {
            Ribbon1 rb = Globals.Ribbons.Ribbon1;
            if (rb == null) return;

            rb.btnFoot.Enabled = enable;
            rb.btnDate.Enabled = enable;
            rb.btnHead.Enabled = enable;
            rb.btnHeadNoLine.Enabled = enable;
            rb.btnInsert2.Enabled = enable;
            rb.btnInsertPaper.Enabled = enable;
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

            rb.mnuInsertGWObject.Enabled = enable;

            rb.mnuSelect.Enabled = enable;
            rb.mnuConvert.Enabled = enable;
            rb.btnAI.Enabled = enable;
        }

        private void ThisAddIn_NewDocument(MSWord.Document Doc)
        {
            try { ControlEnable(true); }
            catch (Exception ex) { Debug.WriteLine(ex.Message); }
        }

        private void Application_DocumentOpen(MSWord.Document Doc)
        {
            try
            {
                if (Doc == null) return;
                bool editable = Doc.ProtectionType == MSWord.WdProtectionType.wdNoProtection;
                ControlEnable(editable);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DocumentOpen: " + ex.Message);
            }
        }

        #endregion
    }
}