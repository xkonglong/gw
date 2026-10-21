using System;
using System.Drawing;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using MSOffice = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.IO;

namespace 公文排版
{
    public partial class Ribbon1
    {
        // ============ 共享字段 ============
        internal MSWord.Application WordApp;                 // Word Application
        internal System.Speech.Synthesis.SpeechSynthesizer sp; // 语音模块
        internal frmPaperSetup fp;                           // 页面设置窗体
        internal string wps = "";                            // 宿主程序名

        // 字体变量
        internal Font iniTitle;
        internal Font iniTitle1;
        internal Font iniTitle2;
        internal Font iniTitle3;
        internal Font iniTitle4;
        internal Font iniNormal;
        internal Font iniHead;
        internal Font iniName;
        internal Font iniFoot;
        internal Font iniPageNumber;

        internal bool IsWps => wps.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) >= 0;

        // ============ 功能区载入/关闭 ============
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
#if DEBUG
            button2.Visible = true;
            button3.Visible = true;
            Debug.WriteLine("运行在调试模式");
#else
            button2.Visible = false;
            button3.Visible = false;
#endif
            try
            {
                btnAbout.Description = "关于小恐龙公文排版助手";
                wps = System.Windows.Forms.Application.ProductName;

                if (wps.Contains("2007"))
                {
                    MessageBox.Show("公文排版助手不再支持 MS Office 2007 \r\n\t 请安装旧版插件(v1.0.8.3 版本),或者升级您的 Office 版本! ",
                        "小恐龙公文排版助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                if (IsWps)
                {
                    ApplyWpsImages();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ribbon1_Load 外观设置: " + ex.Message);
            }

            // 语音初始化
            try
            {
                sp = new System.Speech.Synthesis.SpeechSynthesizer();
                sp.SpeakCompleted += Sp_SpeakCompleted;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("语音模块初始化失败: " + ex.Message);
                btnSpeak.Enabled = false;
                btnSpeak.ScreenTip = "操作系统语音模块加载失败，该功能无法使用！";
            }

            // 中文字体下拉
            try
            {
                cobFonts.Items.Clear();
                using (var ifc = new System.Drawing.Text.InstalledFontCollection())
                {
                    foreach (FontFamily ff in ifc.Families)
                    {
                        if (ff.Name.Contains("宋") || ff.Name.Contains("黑") || ff.Name.Contains("楷"))
                        {
                            var rItem = Factory.CreateRibbonDropDownItem();
                            rItem.Label = ff.Name;
                            cobFonts.Items.Add(rItem);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "读取字体列表时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            // 初始化标题序号菜单和 Word 事件
            string[] index = { "0", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
                               "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十" };
            try
            {
                WordApp = Globals.ThisAddIn.Application;
                WordApp.WindowSelectionChange += WordApp_WindowSelectionChange;

                mnuTitle1.Items.Clear();
                mnuTitle2.Items.Clear();
                mnuTitle3.Items.Clear();
                mnuTitle4.Items.Clear();

                for (int i = 1; i < 21; i++)
                {
                    var mnuTitle1_item = this.Factory.CreateRibbonButton();
                    var mnuTitle2_item = this.Factory.CreateRibbonButton();
                    var mnuTitle3_item = this.Factory.CreateRibbonButton();
                    var mnuTitle4_item = this.Factory.CreateRibbonButton();

                    mnuTitle1_item.Label = "标题" + index[i] + "、";
                    mnuTitle2_item.Label = "标题（" + index[i] + "）";
                    mnuTitle3_item.Label = "标题" + i + ".";
                    mnuTitle4_item.Label = "标题（" + i + "）";

                    mnuTitle1_item.ShowImage = false;
                    mnuTitle2_item.ShowImage = false;
                    mnuTitle3_item.ShowImage = false;
                    mnuTitle4_item.ShowImage = false;

                    mnuTitle1_item.Tag = i;
                    mnuTitle2_item.Tag = i;
                    mnuTitle3_item.Tag = i;
                    mnuTitle4_item.Tag = i;

                    mnuTitle1_item.Click += MnuTitle1_item_Click;
                    mnuTitle2_item.Click += MnuTitle2_item_Click;
                    mnuTitle3_item.Click += MnuTitle3_item_Click;
                    mnuTitle4_item.Click += MnuTitle4_item_Click;

                    mnuTitle1.Items.Add(mnuTitle1_item);
                    mnuTitle2.Items.Add(mnuTitle2_item);
                    mnuTitle3.Items.Add(mnuTitle3_item);
                    mnuTitle4.Items.Add(mnuTitle4_item);
                }

                btnInsertSignPic.SuperTip = btnInsertSignPicLink.SuperTip;
                btnInsertSignPic.Tag = "Link";

                if (WordApp.Documents.Count < 1) Globals.ThisAddIn.ControlEnable(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "小恐龙公文排版助手载入时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            ReadSetting();
            LoadVBA();
        }

        private void Ribbon1_Close(object sender, EventArgs e)
        {
            try
            {
                try { if (WordApp != null) WordApp.WindowSelectionChange -= WordApp_WindowSelectionChange; } catch { }

                if (sp != null) { sp.Dispose(); sp = null; }
                if (fp != null) { if (!fp.IsDisposed) fp.Dispose(); fp = null; }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ribbon1_Close: " + ex.Message);
            }
        }

        private void WordApp_WindowSelectionChange(MSWord.Selection Sel)
        {
            try
            {
                if (Sel == null) return;

                btnUnderDot.Checked = Sel.Range.Font.EmphasisMark == MSWord.WdEmphasisMark.wdEmphasisMarkUnderSolidCircle;
                btnHideFont.Checked = Sel.Range.Font.Hidden != 0;

                bool hasSelection = Sel.Type == MSWord.WdSelectionType.wdSelectionNormal && Sel.Text.Length > 1;
                btnFormatDelSelectSpace.Enabled = hasSelection;
                btnFormatDelSelectEmptyLine.Enabled = hasSelection;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("WordApp_WindowSelectionChange: " + ex.Message);
            }
        }

        // ============ 字体配置 ============
        private void ReadSetting()
        {
            iniTitle = Properties.Settings.Default.公文标题;
            iniTitle1 = Properties.Settings.Default.一级标题;
            iniTitle2 = Properties.Settings.Default.二级标题;
            iniTitle3 = Properties.Settings.Default.三级标题;
            iniTitle4 = Properties.Settings.Default.四级标题;
            iniNormal = Properties.Settings.Default.正文;
            iniHead = Properties.Settings.Default.版头;
            iniName = Properties.Settings.Default.签发人姓名;
            iniFoot = Properties.Settings.Default.版记;
            iniPageNumber = Properties.Settings.Default.页码;

            iniTitle = FormatFontSize(iniTitle);
            iniTitle1 = FormatFontSize(iniTitle1);
            iniTitle2 = FormatFontSize(iniTitle2);
            iniTitle3 = FormatFontSize(iniTitle3);
            iniTitle4 = FormatFontSize(iniTitle4);
            iniNormal = FormatFontSize(iniNormal);
            iniHead = FormatFontSize(iniHead);
            iniName = FormatFontSize(iniName);
            iniFoot = FormatFontSize(iniFoot);
            iniPageNumber = FormatFontSize(iniPageNumber);

            btnTitleHei.SuperTip = "设置公文大标题,规范为二号,小标宋.\r\n\r\n当前字体为:" + (iniTitle?.Name ?? "未设置");
            btnNormal.SuperTip = "设置所选文字为正文格式,规范为三号,仿宋.\r\n\r\n当前字体为:" + (iniNormal?.Name ?? "未设置");
            btnHead.ScreenTip = "设置所选文字为: " + (iniHead?.Name ?? "?") + " " + (iniHead?.Size ?? 0) + "pt";
            btnName.ScreenTip = "设置所选文字为: " + (iniName?.Name ?? "?") + " " + (iniName?.Size ?? 0) + "pt";
            btnFoot.ScreenTip = "设置所选文字为: " + (iniFoot?.Name ?? "?") + " " + (iniFoot?.Size ?? 0) + "pt";

            mnuScript.Items.Clear();
            if (Properties.Settings.Default.VBA历史 == null)
            {
                mnuScript.Visible = false;
                return;
            }
            foreach (string item in Properties.Settings.Default.VBA历史)
            {
                string[] it = item.Split('|');
                if (it.Length < 2) continue;
                var btnVBAlist = this.Factory.CreateRibbonButton();
                btnVBAlist.Label = it[0];
                btnVBAlist.ScreenTip = it[1];
                btnVBAlist.Click += BtnVBAlist_Click;
                mnuScript.Items.Add(btnVBAlist);
            }
        }

        private Font FormatFontSize(Font font)
        {
            if (font == null) return null;
            string fontName = font.Name;
            float sfs = font.SizeInPoints;
            if (Math.Abs(sfs - 10.5f) < 0.01f)
                return new Font(fontName, 10.5f, font.Style);
            sfs = (float)Math.Round(sfs);
            return new Font(fontName, sfs, font.Style);
        }

        internal string GetDllPath()
        {
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            string fullPath = Uri.UnescapeDataString(new UriBuilder(codeBase).Path);
            return Path.GetDirectoryName(fullPath);
        }

        // ============ 关于组 ============
        private void btnXkonglong_Click(object sender, RibbonControlEventArgs e)
        {
            try { Process.Start("https://gw.xkonglong.com"); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n请打开浏览器输入网址访问:\r\n https://gw.xkonglong.com");
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            using (var about = new AboutBox1())
            {
                if (about.ShowDialog() == DialogResult.OK) ReadSetting();
            }
        }

        private void btnEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string aVer = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
                Version oVer = System.Environment.OSVersion.Version;
                string appVer = System.Windows.Forms.Application.ProductName + System.Windows.Forms.Application.ProductVersion;
                string body = "%0d%0a%0d%0a以下我的系统环境信息,供您参考." + "\r\n当前 Windows Ver:" + oVer +
                              "\r\n当前程序版本 ver:" + appVer;
                Process.Start("mailto:xkonglong@qq.com?subject=关于小恐龙公文排版助手v" + aVer + "的意见反馈&body=" + body);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "为发送反馈获取系统信息时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            try { Process.Start(Properties.Settings.Default.在线帮助); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n请打开浏览器输入网址访问:\r\n http://xkonglong.com/wordadd_help/");
            }
        }

        // ============ 功能区角标 ============
        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try { WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFilePageSetup].Show(); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try { WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFormatParagraph].Show(); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void group3_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try { WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFormatFont].Show(); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        // ============ WPS 图标适配（从原 Ribbon1_Load 中抽出） ============
        private void ApplyWpsImages()
        {
            mnuTitle1.OfficeImageId = "_1";
            btnPaper.OfficeImageId = "FileServerTransferDatabase";
            btnInsert2.OfficeImageId = "IndentIncrease";
            glyFormat.OfficeImageId = "AccessFormWizard";
            btnFormatFullLine.OfficeImageId = "DiagramReverseClassic";
            btnFormatR2N.OfficeImageId = "EastAsianEditingMarks";
            btnFormatFrame2txt.OfficeImageId = "EditQuery";
            btnFormatParaPage.OfficeImageId = "CrossReferenceInsert";
            btnNormal.OfficeImageId = "OtherActionsMenu";
            btnFontChangeDialog.OfficeImageId = "QuickStyleSets";
            btnNum2Txt.OfficeImageId = "FileWorkflowTasks";
            btnHead.OfficeImageId = "FieldInsert";
            btnFoot.OfficeImageId = "AttachFile";
            btnFootDouble.OfficeImageId = btnFootSingle.OfficeImageId;
            btnDate.OfficeImageId = "ViewAllProposals";
            btnInsertPaper.OfficeImageId = "PageOrientationGallery";
            glyFoot.OfficeImageId = "PageNambersInFooterInsertGallery";
            glySymbol.OfficeImageId = "EncodingMenu";
            btnNumberOnLocation.OfficeImageId = "FormControlEditBox";
            btnInsertSignPicSave.OfficeImageId = "SignatureShow";
            btnSaveAsPicAt.OfficeImageId = "PictureStylesGallery";
            glyPassword.OfficeImageId = "AdpPrimaryKey";
            btnClearOpenPW.OfficeImageId = "CoverPageRemove";
            btnClearProtectPW.OfficeImageId = "InkDeleteAllInk";
            btnSpeak.OfficeImageId = "AnimationAudio";
            btnSpeakOption.OfficeImageId = "ConditionalFormattingHighlightEqualTo";
            btnXkonglong.OfficeImageId = "OutlookGlobe";
            btnEmail.OfficeImageId = "WorkflowPending";
            mnuSelect.OfficeImageId = "SelectionPane";
            btnSelectObject.OfficeImageId = "SnapToObjects";
            mnuConvert.OfficeImageId = "ArrangeByRecurrence";
            btnAI.OfficeImageId = "GroupTextBreak";

            btnTitleHei.Image = Properties.Resources.TableCellAlignTop;
            btnInsertSignPic.Image = Properties.Resources.StampTool;

            btnPaper.OfficeImageId = "PrintPreviewMargins";
            btnInsert2.OfficeImageId = "ParagraphIndent2Spaces";
            btnLeft.OfficeImageId = "DecreaseIndent";
            glyFormat.OfficeImageId = "SmartRearrange";
            btnFormatAddSpace.OfficeImageId = "TextOnFormField";
            btnNormal.OfficeImageId = "CellAlignment_6";
            btnTitleHei.Image = null;
            btnTitleHei.OfficeImageId = "CellAlignment_1";
            btnNum2Txt.OfficeImageId = "Numbering";
            glyFont.OfficeImageId = "ReplaceFont";
            btnFontChangeDialog.OfficeImageId = "ReplaceFont";
            btnHead.OfficeImageId = "BookProtection";
            btnName.OfficeImageId = "CellAlignment_2";
            btnFoot.OfficeImageId = "CellAlignment_7";
            btnStyle.OfficeImageId = "StyleFormat";
            btnStyleR.OfficeImageId = "StyleShapeTool";
            mnuInsertGWObject.Image = Properties.Resources.OpenOfficialDocumentbubbleImage;
            btnFootSingle.OfficeImageId = "AlignBottom";
            btnFootDouble.OfficeImageId = "AlignBottom";
            btnDate.OfficeImageId = "InsertDateTime";
            btnInsertPaper.OfficeImageId = "InsertHorizonBlankPage";
            glyFoot.OfficeImageId = "PageNumber";
            glySymbol.OfficeImageId = "InsertSymbol";
            btnNumberOnLocation.OfficeImageId = "PageNumber";
            btnInsertSignPicSave.OfficeImageId = "ImportDataSave";
            btnProtectDoc.OfficeImageId = "idocprotect";
            glyPassword.OfficeImageId = "ProtectEyes";
            btnClearOpenPW.OfficeImageId = "PasswordEncryption";
            btnClearProtectPW.OfficeImageId = "Lock";
            btnSpeak.OfficeImageId = "InsertSound";
            btnSpeakOption.OfficeImageId = "Setting_p";
            glyBak.Items[0].OfficeImageId = "ExportToPDF";
            glyBak.Items[1].OfficeImageId = "wpp2doc";
            glyBak.Items[2].OfficeImageId = "exporttopicture32";
            glyBak.Items[3].OfficeImageId = "exporttopicture32";
            btnSaveAsPicAt.OfficeImageId = "ChangePicture";
            btnXkonglong.OfficeImageId = "InsertHyperlink";
            btnEmail.OfficeImageId = "Mail";
            mnuSelect.OfficeImageId = "SelectObjects";
            btnSelectSharp.OfficeImageId = "SelectAll";
            btnSelectTables.OfficeImageId = "SelectTable";
            btnSelectObject.OfficeImageId = "SelectObjects";
            mnuConvert.OfficeImageId = "WrappingTopandBottom";
            mnuInLineSharp.OfficeImageId = "WrapText";
            btnSharp2OnText.OfficeImageId = "WrappingInFrontofText";
            btnSharp2Top.OfficeImageId = "WrappingInFrontofText";
            btnSharp2Wrapping.OfficeImageId = "WrappingSquare";
            btnSharp2TopBonttom.OfficeImageId = "WrappingTopandBottom";
            btnAI.OfficeImageId = "splitwindow";
            btnVBA.OfficeImageId = "SlideShow";
            btnClearVBA.OfficeImageId = "Delete";
            mnuVBA.OfficeImageId = "iplay";

            btnFormatAddSpace.OfficeImageId = "";
            btnFormatC2E.OfficeImageId = "";
            btnFormatDelEmptyLine.OfficeImageId = "";
            btnFormatDelGraphic.OfficeImageId = "";
            btnFormatDelSelectEmptyLine.OfficeImageId = "";
            btnFormatDelSelectSpace.OfficeImageId = "";
            btnFormatDelSpace.OfficeImageId = "";
            btnFormatE2C.OfficeImageId = "";
            btnFormatFrame2txt.OfficeImageId = "";
            btnFormatFullLine.OfficeImageId = "";
            btnFormatParaPage.OfficeImageId = "";
            btnFormatR2N.OfficeImageId = "";
            btnFormatRemoveURL.OfficeImageId = "";

            btnSetPW.Visible = false;
            btnFormatDelSpace.Visible = false;
            btnSelectCurrentPage.Visible = false;
            btnFormatParaPage.Visible = false;
        }
    }
}