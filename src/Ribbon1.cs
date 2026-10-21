using System;
using Microsoft.Office.Tools.Ribbon;
using MSWord = Microsoft.Office.Interop.Word;
using MSOffice = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using Microsoft.Office.Interop.Word;
using Font = System.Drawing.Font;
using Microsoft.Vbe.Interop;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace 公文排版
{
    public partial class Ribbon1
    {
        //一些值得注意的事情:
        //  word97-2003文档不支持 Table.Descr 属性.
        //

        MSWord.Application WordApp;  //定义WordApplication
        public System.Speech.Synthesis.SpeechSynthesizer sp;  //定义语音模块
        frmPaperSetup fp;//定义页面设置窗体
        string wps = "";  //主程序名称, 提升共用范围,便于其他函数使用.

        //字体变量
        Font iniTitle ;
        Font iniTitle1;
        Font iniTitle2;
        Font iniTitle3;
        Font iniTitle4;
        Font iniNormal;
        Font iniHead ;
        Font iniName ;
        Font iniFoot ;
        Font iniPageNumber;
        //////////////////////////////////

        //=======================以下是功能函数区域=======================
        /// <summary>
        /// 设置Word所选文字字体格式,包括字体名称,大小,是否粗体\斜体
        /// </summary>
        /// <param name="myFont">Font class 字体名称 </param>
        private void SetFont(Font myFont)
        {
            try
            {
               // float[] sSize = { 14.25f,15.75f, 21.75f,26.25f};
                //float wfs =(float) Math.Round(myFont.SizeInPoints, 1);
                WordApp.Selection.Font.Name = myFont.Name;
                WordApp.Selection.Font.Size = myFont.Size;

                //Word字号与系统字号并不完全一致,需要修改.
                //switch (myFont.Size)
                //{
                //    case 14.25f:
                //        WordApp.Selection.Font.Size = 14;
                //        break;
                //    case 15.75f:
                //        WordApp.Selection.Font.Size = 16;
                //        break;
                //    case 21.75f:
                //        WordApp.Selection.Font.Size = 22;
                //        break;
                //    case 26.25f:
                //        WordApp.Selection.Font.Size = 26;
                //        break;
                //}
                if (myFont.Bold) {
                    WordApp.Selection.Font.Bold = 1;
                };
                if (myFont.Italic) { WordApp.Selection.Font.Italic = 1; }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"设置字体时出错");
            }
        }


        /// <summary>
        /// 设置页码位置
        /// </summary>
        /// <param name="position">0 居中, 1 外侧,2 居右</param>
        private void PageNumberAligment(int position)
        {
            try
            {
                //如果是普通视图或者大纲视图,转为显示页面视图
                if (WordApp.ActiveWindow.ActivePane.View.Type == MSWord.WdViewType.wdNormalView || WordApp.ActiveWindow.ActivePane.View.Type == MSWord.WdViewType.wdOutlineView)
                {
                    WordApp.ActiveWindow.ActivePane.View.Type = MSWord.WdViewType.wdPrintView;
                }
                WordApp.ScreenUpdating = false;
                switch (position)
                {
                    case 0:  //居中
                        foreach (MSWord.Section s in WordApp.ActiveDocument.Sections)
                        {
                            s.PageSetup.OddAndEvenPagesHeaderFooter = 0;

                            //页脚居中
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageFooter;//进入页脚视图
                        InsertNumber(); //插入页码
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//恢复文档视图
                    break;

                    case 1:  //外侧
                        if (WordApp.ActiveWindow.Panes[1].Pages.Count <= 1)
                        {
                            MessageBox.Show("文档仅为1页时,无法设置外侧页码.", "插入页码", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (wps.Contains("WPS"))
                        {
                            WordApp.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = -1; //1=奇偶页不同
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekPrimaryFooter;//进入页脚视图
                            WordApp.Selection.ParagraphFormat.CharacterUnitRightIndent = 1;   //居右页码右缩进1字符  
                            WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                            InsertNumber(); //插入页码
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//恢复文档视图
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekEvenPagesFooter;
                            WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 1; //居左页码左缩进1字符
                            WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                            InsertNumber(); //插入页码
                            WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//恢复文档视图
                        }
                        else
                        {
                            foreach (MSWord.Section s in WordApp.ActiveDocument.Sections)
                            {
                                s.PageSetup.OddAndEvenPagesHeaderFooter = -1; //1=奇偶页不同

                                MSWord.Range r;

                                //奇数页页码
                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                //设置字体
                                r.Font.Size = (float)Math.Round(iniPageNumber.Size, 1);
                                r.Font.Name = iniPageNumber.Name;
                                if (iniPageNumber.Bold) r.Font.Bold = 1;
                                if (iniPageNumber.Italic) r.Font.Italic = 1;

                                //设置页码
                                r.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;   //居右
                                r.Text = "— ";
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Fields.Add(r, MSWord.WdFieldType.wdFieldEmpty, "PAGE \\* Arabic");
                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Text = " —";
                                r.ParagraphFormat.CharacterUnitRightIndent = 1;  //右缩进1字符

                                //偶数页页码
                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;

                                //设置字体
                                r.Font.Size = (float)Math.Round(iniPageNumber.Size, 1);
                                r.Font.Name = iniPageNumber.Name;
                                if (iniPageNumber.Bold) r.Font.Bold = 1;
                                if (iniPageNumber.Italic) r.Font.Italic = 1;

                                //设置页码
                                r.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;   //居左
                                r.Text = "— ";
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Fields.Add(r, MSWord.WdFieldType.wdFieldEmpty, "PAGE \\* Arabic");
                                r = s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                                r.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                                r.Text = " —";
                                r.ParagraphFormat.CharacterUnitFirstLineIndent = 1;  //左缩进1字符
                            }
                        }
                        break;

                    case 2: //居右
                        foreach (MSWord.Section s in WordApp.ActiveDocument.Sections)
                        {
                            s.PageSetup.OddAndEvenPagesHeaderFooter = 0; //0=奇偶页相同
                            //设置页脚居右
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                            s.Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                        }
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageFooter;//进入页脚视图
                        InsertNumber(); //插入页码
                        WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//恢复文档视图
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "pageNumberAligment 方法出错");
            }
            finally { WordApp.ScreenUpdating = true; }
        }



        /// <summary>
        /// 插入页码,设置格式为 — 1 —
        /// </summary>
        /// <param name="align">对齐方式,center=0,left=1,right=2</param>
        private void InsertNumber()
        {
            try
            {
                WordApp.Selection.Font.Size = (float)Math.Round(iniPageNumber.Size,1);
                WordApp.Selection.Font.Name = iniPageNumber.Name;
                if (iniPageNumber.Bold) WordApp.Selection.Font.Bold = 1;
                if (iniPageNumber.Italic) WordApp.Selection.Font.Italic = 1;
                WordApp.Selection.TypeText("— ");
                WordApp.Selection.Fields.Add(WordApp.Selection.Range, MSWord.WdFieldType.wdFieldEmpty, "PAGE \\* Arabic");
                WordApp.Selection.TypeText(" —");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"InsertNumber 方法出错");
            }
        }
        
        

        /// <summary>
        /// 插入版记线
        /// </summary>
        /// <param name="LineNumber">版记行数,默认为1行</param>
        private void InsertLine(int LineNumber = 1)
        {
            //try
            {
                if (LineNumber < 1) LineNumber = 1;
                MSWord.Document WordDoc = WordApp.ActiveDocument;
                MSWord.Table newTable = WordDoc.Tables.Add(WordApp.Selection.Range, LineNumber, 1);

                //设置版记线永远在页面最底部.
                newTable.Rows.RelativeVerticalPosition = MSWord.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                //newTable.Rows.VerticalPosition = -999997f;   // wdTableBottom的浮点值
                newTable.Rows.VerticalPosition = (float)MSWord.WdFramePosition.wdFrameBottom;

                //newTable.Descr = "公文版记";  //表格的说明,旧版doc2003兼容文档不支持此属性

                //设置表格样式

                newTable.PreferredWidth = 443; //单位磅,大约是156mm
                newTable.Rows.Height = 28.5f; //表格行高
                
               // newTable.AllowPageBreaks = false;//不允许分割 wps2016不支持
                
                newTable.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowLeft; //居左
                

                //1磅=0.35146mm, 规范为上下粗线0.35mm, word里100pt=1磅     
                // 规范:中线为0.25mm, 合0.7磅, word里近似值为075pt

                newTable.Borders[MSWord.WdBorderType.wdBorderTop].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                newTable.Borders[MSWord.WdBorderType.wdBorderTop].LineWidth = MSWord.WdLineWidth.wdLineWidth100pt;    

                newTable.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                newTable.Borders[MSWord.WdBorderType.wdBorderBottom].LineWidth = MSWord.WdLineWidth.wdLineWidth100pt;

                if (LineNumber >1)//设置中线
                {
                    newTable.Borders[MSWord.WdBorderType.wdBorderHorizontal].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                    newTable.Borders[MSWord.WdBorderType.wdBorderHorizontal].LineWidth = MSWord.WdLineWidth.wdLineWidth075pt;
                }

                //设置版记字体
                newTable.Select();
                WordApp.Selection.Cells.VerticalAlignment =MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter; //文本上下居中
                SetFont(iniFoot);
            }
           /* catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "插入版记线出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
           */
        }
       

        /// <summary>
        /// 设置标题编号的方法
        /// </summary>
        /// <param name="sender">ribbonButton 对象</param>
        /// <param name="index">int 1,2,3,4代表四个标题级别,默认为1.</param>
        private void SetTitle(object sender, int index = 1)
        {
            try
            {
                MSWord.ListTemplate myTemplate = WordApp.ActiveDocument.ListTemplates.Add();  //添加编号
                myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;//编号与文字之间无空格无Tab
                myTemplate.ListLevels[1].TextPosition = 0; //设置第二行缩进量
                myTemplate.ListLevels[1].NumberPosition = 0;

                int at = Convert.ToInt32(((RibbonButton)sender).Tag);

                if (Convert.ToInt32(at) <= 0) at = 1; myTemplate.ListLevels[1].StartAt = at;//设置序列

                switch (index)
                {
                    case 1:
                        myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                        myTemplate.ListLevels[1].NumberFormat = "%1、";//一
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel1;    //设置大纲级别1
                        break;

                    case 2:
                        myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                        myTemplate.ListLevels[1].NumberFormat = "（%1）";//（一）
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel2;    //设置大纲级别2
                        break;

                    case 3:
                        
                        myTemplate.ListLevels[1].NumberFormat = "%1．";//1
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel3;//设置大纲级别3
                        break;

                    case 4:
                        
                        myTemplate.ListLevels[1].NumberFormat = "（%1）";//（1）
                        WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel4;//设置大纲级别4
                        break;
                }
                WordApp.Selection.Range.ListFormat.ApplyListTemplate(myTemplate); //应用新格式项目编号
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "设置标题格式时出错!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }


        /// <summary>
        /// 查找替换功能
        /// </summary>
        /// <param name="wildcard">是否支持通配符</param>
        /// <param name="findString">要查找的字符</param>
        /// <param name="replaceString">要替换的字符</param>
        /// <param name="range">替换范围, range=true时是所选范围,默认=false ,是全部</param>
        private void WordReplace(bool wildcard, string findString, string replaceString,bool range=false)
        {
            try
            {
                WordApp.Selection.Find.ClearFormatting();
                WordApp.Selection.Find.Replacement.ClearFormatting();
                if(range==true)//范围内
                {
                    WordApp.Selection.Find.Execute(findString, false, false, wildcard, false, false, true,
                    MSWord.WdFindWrap.wdFindStop, false, replaceString, MSWord.WdReplace.wdReplaceAll);
                }
                else   //全部
                {
                    WordApp.Selection.Find.Execute(findString, false, false, wildcard, false, false, true,MSWord.WdFindWrap.wdFindContinue, false, replaceString, MSWord.WdReplace.wdReplaceAll);
                }
                WordApp.Selection.Find.ClearFormatting();
                WordApp.Selection.Find.Replacement.ClearFormatting();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "排版替换功能出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        /// <summary>
        /// 生成pdf副本函数
        /// </summary>
        /// <param name="fileName">副本文件名</param>
        private void SavePdf(string filName)
        {
            WordApp.StatusBar = "正在生成 PDF 格式副本...";
            MSWord.WdExportFormat pdf = MSWord.WdExportFormat.wdExportFormatPDF;
            MSWord.WdExportCreateBookmarks CreateBookmarks = MSWord.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks;
            WordApp.ActiveDocument.ExportAsFixedFormat(filName, pdf, true,0,0,0,0,0,false,false,CreateBookmarks);
            WordApp.StatusBar = "已生成 PDF 格式副本:" + filName;
        }

        /// <summary>
        /// 生成 doc 格式副本
        /// </summary>
        /// <param name="filenName">副本文件名</param>
        /// <param name="oldFile">原文件名称</param>
        private void SaveDoc(string filenName, string oldFile)
        {
                //另存为命令会导致当前文档直接变为另存为的文档.所以需要用open命令打开原格式文档,并激活该文档窗口.
                WordApp.StatusBar = "正在生成 Word97-2003 格式副本...";
                WordApp.ActiveDocument.SaveAs2(filenName, MSWord.WdSaveFormat.wdFormatDocument97);
                WordApp.StatusBar = "已生成 Word97-2003 格式副本:" + filenName;
                WordApp.Documents.Open(oldFile);
                int w = WordApp.Windows.Count;
                if (w > 1)
                {
                    WordApp.Windows[w - 1].Activate();
                }
        }

        /// <summary>
        /// 插入印章的方法
        /// </summary>
        /// <param name="FileName">图片文件路径</param>
        /// <param name="LinkToFile">bool值,是否以链接方式插入图片</param>
        private void InsertSignPic(string FileName, bool LinkToFile = true)
        {
            MSWord.InlineShape pic = WordApp.Selection.Range.InlineShapes.AddPicture(FileName, LinkToFile, !LinkToFile);
            MSWord.Shape sPic = pic.ConvertToShape(); //图片转换为形状
            sPic.ZOrder(MSOffice.MsoZOrderCmd.msoBringInFrontOfText);//浮于文字上方
        }

        /// <summary>
        /// 替换字体
        /// </summary>
        /// <param name="oldFont">原字体中文名称</param>
        /// <param name="newFont">新字体中文名称</param>
        private void FontSwitch(string oldFont, string newFont)
        {
            try
            {
                //undo无法记录.如果word的undo为空,记录执行替换操作就会崩溃.
                // 替换中文字体
                WordApp.Selection.Find.ClearFormatting();  //清除查找格式
                WordApp.Selection.Find.Font.NameFarEast = oldFont; //设置中文所用字体 
                //WordApp.Selection.Find.Font.NameAscii = oldFont; //设置拉丁文所用字体
                WordApp.Selection.Find.Replacement.ClearFormatting(); //清除替换格式
                //WordApp.Selection.Find.Replacement.Font.Name = newFont;   //设置新的中文字体
                WordApp.Selection.Find.Replacement.Font.NameFarEast = newFont;   //设置新的中文字体
                //开始替换
                WordApp.Selection.Find.Execute("", false, false, false, false, false, true,
                        MSWord.WdFindWrap.wdFindContinue, true, "", MSWord.WdReplace.wdReplaceAll);
                WordApp.StatusBar = String.Format("字体{0}已替换为{1}.", oldFont, newFont);
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "替换字体时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        // 读取字体自定义设置
        private void ReadSetting()
        {
            //先读取配置
            iniTitle = Properties.Settings.Default.公文标题;
            iniTitle1 = Properties.Settings.Default.一级标题;
            iniTitle2 = Properties.Settings.Default.二级标题;
            iniTitle3 = Properties.Settings.Default.三级标题;
            iniTitle4 = Properties.Settings.Default.四级标题;
            iniNormal = Properties.Settings.Default.正文;
            iniHead = Properties.Settings.Default.版头;

            Debug.WriteLine(iniHead.Style.ToString());
            iniName = Properties.Settings.Default.签发人姓名;
            iniFoot = Properties.Settings.Default.版记;
            iniPageNumber = Properties.Settings.Default.页码;

            //再统一字体大小
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

            btnTitleHei.SuperTip = "设置公文大标题,规范为二号,小标宋.\r\n\r\n当前字体为:" + iniTitle.Name;
            btnNormal.SuperTip= "设置所选文字为正文格式,规范为三号,仿宋.\r\n\r\n当前字体为:" + iniNormal.Name;
            btnHead.ScreenTip = "设置所选文字为: " + iniHead.Name+" "+iniHead.Size+"pt";
            btnName.ScreenTip= "设置所选文字为: " + iniName.Name + " " + iniName.Size + "pt";
            btnFoot.ScreenTip= "设置所选文字为: " + iniFoot.Name + " " + iniFoot.Size + "pt";

            //读取vba历史菜单

            if (Properties.Settings.Default.VBA历史==null)
            {
                mnuScript.Visible = false;
                return;
            }
            foreach (string item in Properties.Settings.Default.VBA历史)
            {
                string[] it = item.Split('|');
                RibbonButton btnVBAlist = this.Factory.CreateRibbonButton();

                btnVBAlist.Label = it[0];
                btnVBAlist.ScreenTip = it[1];
                btnVBAlist.Click += BtnVBAlist_Click;
                mnuScript.Items.Add(btnVBAlist);
            }


        }
        /// <summary>
        /// 统一word和系统字体的大小
        /// </summary>
        /// <param name="font">原字体</param>
        /// <returns>新字体</returns>
        private Font FormatFontSize(Font font)
        {
            string fontName = font.Name;
            float sfs = font.SizeInPoints;
            switch (sfs)
            {
                case 10.5f:
                    font = new Font(fontName, 10.5f, font.Style);
                    break;
                default:
                    sfs = (float)Math.Round(sfs);
                    font = new Font(fontName, sfs, font.Style);
                    break;
            }
            return font;
        }

        /// <summary>
        /// 获取本插件所在路径
        /// </summary>
        /// <returns>返回路径string</returns>
        string GetDllPath()
        {
            //获取程序集完整路径  file:///xxx//xx//xx.dll
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            //去掉file:///
            string fullPath = Uri.UnescapeDataString(new UriBuilder(codeBase).Path);

            //获取启动路径
            string startPath = Path.GetDirectoryName(fullPath);
            return startPath;
        }

        /// <summary>
        ///  插入版头的方法
        /// </summary>
        /// <param name="autoText">要插入的自动图文集名称</param>
        void InsertHead(string autoText)
        {
            //载入公文部件
            string startPath = GetDllPath();
            string dotFilename = Path.Combine(startPath, "Resources\\公文部件.dotx");
            Debug.WriteLine("公文部件:" + dotFilename);
            //判断dotx是否存在,存在就附加到当前文档
            if (!File.Exists(dotFilename))
            {
                MessageBox.Show("未找到公文部件,请重新安装/修复本插件!", "载入公文部件时出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                try
                {
                    WordApp.AddIns.Add(dotFilename, Install: true);
                }
                catch (Exception ex)
                {
                    string err = ex.Message + "\r\n\r\n" + ex.Source;
                    MessageBox.Show(err, "加载公文部件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            try
            {
                //插入该模板的图文集
                WordApp.Selection.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);
                MSWord.Template template = WordApp.Templates[dotFilename];
                if (template.BuildingBlockEntries.Count < 1) { return; }
                for (int i = 1; i <= template.BuildingBlockEntries.Count; i++)
                {
                    Debug.WriteLine(template.BuildingBlockEntries.Item(i).Name);
                }
                template.BuildingBlockEntries.Item(autoText).Insert(Where: WordApp.Selection.Range);
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "插入版头时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        /// <summary>
        /// 单个的文本框转文字功能,  目前无法实现循环.
        /// 对选定的文本框转为文字
        /// </summary>
        private void Frame2txt()
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("文本框转文字");

            //MessageBox.Show(WordApp.ActiveDocument.Shapes.Count.ToString());
            //foreach (Shape shp in WordApp.ActiveDocument.Shapes)
            //{
            //    //WordApp.Selection.EndKey(MSWord.WdUnits.wdStory);
            //    WordApp.Selection.TypeText(shp.TextFrame.TextRange.Text);
            //    shp.Delete();
            //}
            if (WordApp.Selection.ShapeRange.Count > 0)
            {
                Shape sh = WordApp.Selection.ShapeRange[1];
                string t = sh.TextFrame.TextRange.Text;
                sh.Delete();
                WordApp.Selection.TypeText(t);
            }
            undo.EndCustomRecord();
        }


        //==============================================
        //              以上是功能函数区域
        //              以下是事件代码区域
        //==============================================

        private void Ribbon1_Close(object sender, EventArgs e) //功能区面板关闭事件
        {
            try
            {
                if (sp.Equals(null))
                {
                    sp.Dispose();//释放语音模块
                }
                if (!fp.IsDisposed)
                {
                    fp.Dispose();//释放页面设置对话框
                }
            }
            catch { }
        }

        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)//载入功能区面板
        {
            //调试模式开启试验功能.
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

                //针对 word 2007的提示
                if (wps.Contains("2007"))
                {
                    MessageBox.Show("公文排版助手不再支持 MS Office 2007 \r\n\t 请安装旧版插件(v1.0.8.3 版本),或者升级您的 Office 版本! ", "小恐龙公文排版助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                //针对WPS 的一些设置.
                if (wps.Contains("WPS"))
                {
                    //wps更改imageMSO图片,以下都是专业版支持的
                    mnuTitle1.OfficeImageId = "_1";
                    btnPaper.OfficeImageId = "FileServerTransferDatabase"; //页面设置
                    btnInsert2.OfficeImageId = "IndentIncrease";    //中文缩进
                    glyFormat.OfficeImageId = "AccessFormWizard";
                    //btnFormatFullLine.OfficeImageId = "DiagramReverseClassic";  //最高支持到 WPS文字(11.1.0.11045)
                    
                    btnFormatFullLine.OfficeImageId = "DiagramReverseClassic";
                    btnFormatR2N.OfficeImageId = "EastAsianEditingMarks";
                    btnFormatFrame2txt.OfficeImageId = "EditQuery";
                    btnFormatParaPage.OfficeImageId = "CrossReferenceInsert";

                    btnNormal.OfficeImageId = "OtherActionsMenu";
                    //glyFont.OfficeImageId = "QuickStyleSets";
                    btnFontChangeDialog.OfficeImageId = "QuickStyleSets";
                    btnNum2Txt.OfficeImageId = "FileWorkflowTasks";
                    btnHead.OfficeImageId = "FieldInsert";
                    btnFoot.OfficeImageId = "AttachFile";
                    //mnuInsertGWObject.OfficeImageId = "FileMarkAsFinal";
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

                    //2016还需要一些
                    btnTitleHei.Image = Properties.Resources.TableCellAlignTop;
                    btnInsertSignPic.Image = Properties.Resources.StampTool;


                    //这些是wps个人版支持的图标
                    btnPaper.OfficeImageId = "PrintPreviewMargins";   //wps个人版 页面设置
                    btnInsert2.OfficeImageId = "ParagraphIndent2Spaces";    //wps个人版 中文缩进
                    btnLeft.OfficeImageId = "DecreaseIndent";  //段落顶格
                    glyFormat.OfficeImageId = "SmartRearrange";//排版
                    btnFormatAddSpace.OfficeImageId = "TextOnFormField";//中英文之间添加空格


                    btnNormal.OfficeImageId = "CellAlignment_6";//正文
                    btnTitleHei.Image = null;
                    btnTitleHei.OfficeImageId = "CellAlignment_1";//公文标题

                    btnNum2Txt.OfficeImageId = "Numbering";//编号转文字
                    glyFont.OfficeImageId = "ReplaceFont";// "Style";//字体更改菜单
                    btnFontChangeDialog.OfficeImageId = "ReplaceFont";//字体更改对话框

                    btnHead.OfficeImageId = "BookProtection"; //三号黑体,密级
                    btnName.OfficeImageId = "CellAlignment_2";//三号楷体,签发人
                    btnFoot.OfficeImageId = "CellAlignment_7";//四号宋体,版记
                    btnStyle.OfficeImageId = "StyleFormat";//样式
                    btnStyleR.OfficeImageId = "StyleShapeTool";//强制重建 

                    mnuInsertGWObject.Image = Properties.Resources.OpenOfficialDocumentbubbleImage;  //公文部件
                    //mnuInsertGWObject.OfficeImageId = "OfficialTemplate";  //"OfficialHead";//公文部件
                    //btnHeadSingle.Image = null;
                    //btnHeadSingle.OfficeImageId = "OpenOfficialDocumentbubbleImage"; //发文机关
                    //btnHeadMulti.Image = null;
                    //btnHeadMulti.OfficeImageId = "OpenOfficialDocumentbubbleImage"; //发文机关
                    btnFootSingle.OfficeImageId = "AlignBottom";//插入版记
                    btnFootDouble.OfficeImageId = "AlignBottom";//插入版记

                    btnDate.OfficeImageId = "InsertDateTime";//插入日期
                    btnInsertPaper.OfficeImageId = "InsertHorizonBlankPage";//插入横页
                    glyFoot.OfficeImageId = "PageNumber";//插入页码
                    glySymbol.OfficeImageId = "InsertSymbol";//插入符号
                    btnNumberOnLocation.OfficeImageId= "PageNumber";//插入页码位置-光标处

                    //btnInsertSignPic.Image = Properties.Resources.StampTool; //插入公章
                    //btnInsertSignPic.OfficeImageId = "Save";////插入公章
                    btnInsertSignPicSave.OfficeImageId = "ImportDataSave"; //保存并插入公章

                    btnProtectDoc.OfficeImageId = "idocprotect";//限制编辑
                    glyPassword.OfficeImageId = "ProtectEyes";// "DocumnetsEncryption";//加密
                    btnClearOpenPW.OfficeImageId = "PasswordEncryption";  //清除打开密码
                    btnClearProtectPW.OfficeImageId = "Lock";  //解除修改密码
                    btnSpeak.OfficeImageId = "InsertSound";//朗读
                    btnSpeakOption.OfficeImageId = "Setting_p";//朗读选项

                    //glyBak.Items[0].Image = null;
                    glyBak.Items[0].OfficeImageId="ExportToPDF";
                    glyBak.Items[1].OfficeImageId = "wpp2doc";
                    glyBak.Items[2].OfficeImageId = "exporttopicture32";
                    glyBak.Items[3].OfficeImageId = "exporttopicture32";
                    btnSaveAsPicAt.OfficeImageId = "ChangePicture";

                    btnXkonglong.OfficeImageId = "InsertHyperlink";
                    btnEmail.OfficeImageId = "Mail";  //建议反馈

                    mnuSelect.OfficeImageId = "SelectObjects";
                    btnSelectSharp.OfficeImageId = "SelectAll";
                    btnSelectTables.OfficeImageId = "SelectTable";
                    btnSelectObject.OfficeImageId = "SelectObjects";

                    mnuConvert.OfficeImageId= "WrappingTopandBottom";
                    mnuInLineSharp.OfficeImageId = "WrapText";
                    btnSharp2OnText.OfficeImageId = "WrappingInFrontofText";
                    btnSharp2Top.OfficeImageId = "WrappingInFrontofText";
                    btnSharp2Wrapping.OfficeImageId = "WrappingSquare";
                    btnSharp2TopBonttom.OfficeImageId = "WrappingTopandBottom";

                    btnAI.OfficeImageId = "splitwindow";// "Split";
                    btnVBA.OfficeImageId = "SlideShow"; //运行VBA
                    btnClearVBA.OfficeImageId = "Delete";//"清理代码";
                    mnuVBA.OfficeImageId="iplay";  //预设脚本


                    //取消wps下排版子菜单图标, 最新版不支持了
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

                    /*
                    //wps需要载入按钮图片
                    btnPaper.Image = Properties.Resources.PageMarginsGallery;

                    //排版菜单子项按钮图标, wps新不如旧,懒得上图标了.
                    //
                    btnInsert2.Image = Properties.Resources.AlignJustifyMedium;
                    btnLeft.Image = Properties.Resources.IndentDecrease;
                    glyFormat.Image = Properties.Resources.AsianTypography;
                    btnTitleHei.Image = Properties.Resources.TableCellAlignTop;
                    btnNormal.Image = Properties.Resources.GroupTextBoxText;
                    glyFont.Image = Properties.Resources.FontsReplaceFonts;
                    btnFontChangeDialog.Image = Properties.Resources.FontsReplaceFonts;
                    btnNum2Txt.Image = Properties.Resources.FormattingUnique;
                    btnHead.Image = Properties.Resources.HeaderFooterEditHeader;
                    btnName.Image = Properties.Resources.NewDistributionList;
                    btnFoot.Image = Properties.Resources.HeaderFooterEditFooter;
                    btnStyle.Image = Properties.Resources.ChangeStylesMenu;
                    btnStyleR.Image = Properties.Resources.ChangeStylesMenu;
                    mnuInsertGWObject.Image = Properties.Resources.QuickSearchSubject;
                    btnDate.Image = Properties.Resources.DateAndTimeInsert;
                    btnInsertPaper.Image = Properties.Resources.PageOrientationPortraitLandscape;
                    btnFootSingle.Image = Properties.Resources.TableCellAlignBottom;
                    btnFootDouble.Image = Properties.Resources.TableCellAlignBottomCenter;
                    glyFoot.Image = Properties.Resources.HeaderFooterPageNumberInsert;
                    //btnNumberCenter.Image = Properties.Resources.AlignCenter;
                    //btnNumberOutside.Image = Properties.Resources.AlignRight;
                    btnNumberOnLocation.Image = Properties.Resources.CustomPageMargins;
                    glySymbol.Image = Properties.Resources.StarUnratedFull;
                    btnInsertSignPic.Image = Properties.Resources.StampTool;
                    btnInsertSignPicLink.Image = Properties.Resources.EditLinks;
                    btnInsertSignPicSave.Image = Properties.Resources.SaveAttachments;
                    glyBak.Image = Properties.Resources.FileSaveAs;
                    btnSaveAsPDF.Image = Properties.Resources.FileEmailAsPdfEmailAttachment;
                    btnSaveAsDoc.Image = Properties.Resources.FileSaveAsWord97_2003;
                    btnSaveAsPic.Image = Properties.Resources.SaveAsPng;
                    btnSaveAsJPG.Image = Properties.Resources.SaveAsJpg;
                    btnSaveAsPicAt.Image = Properties.Resources.PicturePropertiesAlternativeText;
                    btnProtectDoc.Image = Properties.Resources.ProtectDocument;
                    glyPassword.Image = Properties.Resources.PasswordProtectSection;
                    btnSetPW.Image = Properties.Resources.WebControlPassword;
                    btnClearOpenPW.Image = Properties.Resources.SetDatabasePassword;
                    btnClearProtectPW.Image = Properties.Resources.HighlightClear;
                    btnSpeak.Image = Properties.Resources.SpeakCells;
                    btnSpeakOption.Image = Properties.Resources.GroupAudioVideoRecordingOptions;
                    btnXkonglong.Image = Properties.Resources.GroupInsertLinks;
                    btnHelp.Image = Properties.Resources.Help;
                    btnEmail.Image = Properties.Resources.OfficeButtonFeedback;
                    */
                    btnSetPW.Visible = false;//wps不兼容,暂时隐藏.
                    btnFormatDelSpace.Visible = false; //wps 不支持替换^w,所以隐藏"删除空白"按钮
                    btnSelectCurrentPage.Visible = false; //wps 不支持选择当前页,隐藏
                    btnFormatParaPage.Visible = false; //WPS不支持直接调用内部命令,隐藏.

                }
            }
            catch (Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
            }

            ///语音功能初始化
            try
            {
                //初始化语音朗读
                sp = new System.Speech.Synthesis.SpeechSynthesizer();
                sp.SpeakCompleted += Sp_SpeakCompleted;
            }            
            catch(Exception ex)
            {
                btnSpeak.Enabled = false;
                btnSpeak.ScreenTip = "操作系统语音模块加载失败，该功能无法使用！";
            }
            ///载入字体

            try
            {
                //载入字体
                cobFonts.Items.Clear();
                InstalledFontCollection ifc = new InstalledFontCollection();
                FontFamily[] ffs = ifc.Families;
                //Font f;
                cobFonts.Items.Clear();
                foreach (FontFamily ff in ffs)
                {
                    // 载入符合条件的中文字体到下拉列表
                    if (ff.Name.Contains("宋") || ff.Name.Contains("黑") || ff.Name.Contains("楷"))
                    {
                        RibbonDropDownItem rItem = Factory.CreateRibbonDropDownItem();
                        rItem.Label = ff.Name;
                        
                        cobFonts.Items.Add(rItem);
                    }
                }
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "读取字体列表时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }


            //初始化标题文字菜单
            string[] index = { "0", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十" };

            try
            {
                WordApp = Globals.ThisAddIn.Application;  //初始化WordApplication
                WordApp.WindowSelectionChange += WordApp_WindowSelectionChange; //注册文字选择事件

                //载入mnuTitle序号值
                mnuTitle1.Items.Clear();
                mnuTitle2.Items.Clear();
                mnuTitle3.Items.Clear();
                mnuTitle4.Items.Clear();

                for (int i = 1; i < 21; i++)
                {
                    RibbonButton mnuTitle1_item = this.Factory.CreateRibbonButton();
                    RibbonButton mnuTitle2_item = this.Factory.CreateRibbonButton();
                    RibbonButton mnuTitle3_item = this.Factory.CreateRibbonButton();
                    RibbonButton mnuTitle4_item = this.Factory.CreateRibbonButton();

                    mnuTitle1_item.Label = "标题" + index[i]+"、";
                    mnuTitle2_item.Label="标题（"+index[i] + "）";
                    mnuTitle3_item.Label = "标题" + i+".";

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
                //初始化盖章按钮默认值
                btnInsertSignPic.SuperTip = btnInsertSignPicLink.SuperTip;
                btnInsertSignPic.Tag = "Link";

                //选定范围内的按钮默认为禁用
                //btnFormatDelSelectEmptyLine.Enabled = false;
                //btnFormatDelSelectSpace.Enabled = false;

                if(WordApp.Documents.Count<1)   Globals.ThisAddIn.ContorlEnable(false);
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "小恐龙公文排版助手载入时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //载入字体设置
            ReadSetting();
            LoadVBA();
        }

 
        private void WordApp_WindowSelectionChange(MSWord.Selection Sel) //WordApp 文字选择事件
        {
            try
            {
                if (WordApp.Selection.Range.Font.EmphasisMark == MSWord.WdEmphasisMark.wdEmphasisMarkUnderSolidCircle)
                {
                    btnUnderDot.Checked = true;  //下划线按钮
                }
                else
                {
                    btnUnderDot.Checked = false;
                }

                if (WordApp.Selection.Range.Font.Hidden == 0)
                {
                    btnHideFont.Checked = false;
                }
                else
                {
                    btnHideFont.Checked = true;
                }

                //如果文档是保护状态,控件变灰
                if (WordApp.ActiveDocument.ProtectionType != MSWord.WdProtectionType.wdNoProtection)
                {
                    Globals.ThisAddIn.ContorlEnable(false);
                    btnSpeak.Enabled = true;
                    glyBak.Enabled = true;
                }
                else
                {
                    Globals.ThisAddIn.ContorlEnable();
                }

                if (WordApp.Selection.Type == MSWord.WdSelectionType.wdSelectionNormal)
                {
                    if (WordApp.Selection.Text.Length > 1)
                    {
                        btnFormatDelSelectSpace.Enabled = true;
                        btnFormatDelSelectEmptyLine.Enabled = true;
                    }
                    else
                    {
                        btnFormatDelSelectSpace.Enabled = false;
                        btnFormatDelSelectEmptyLine.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
            }
        }

        //=============标题和文字格式功能区===============
        private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)//打开段落设置对话框
        {
            try
            {
                WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFormatParagraph].Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void group3_DialogLauncherClick(object sender, RibbonControlEventArgs e) //打开字体设置对话框
        {
            try
            {
                WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFormatFont].Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void MnuTitle4_item_Click(object sender, RibbonControlEventArgs e)//四级标题
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为四级标题");
                SetFont(iniTitle4);
                SetTitle(sender, 4);
            undo.EndCustomRecord();
        }

        private void MnuTitle3_item_Click(object sender, RibbonControlEventArgs e)//三级标题
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为三级标题");
                SetFont(iniTitle3);
                SetTitle(sender, 3);
            undo.EndCustomRecord();
        }

        private void MnuTitle2_item_Click(object sender, RibbonControlEventArgs e) //二级标题
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为二级标题");
                SetFont(iniTitle2);
                SetTitle(sender, 2);
            undo.EndCustomRecord();
        }

        private void MnuTitle1_item_Click(object sender, RibbonControlEventArgs e)//一级标题
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为一级标题");
                SetFont(iniTitle1);
                SetTitle(sender, 1);
            undo.EndCustomRecord();
        }

        private void BtnNum2Txt_Click(object sender, RibbonControlEventArgs e)//编号转文字
        {
            try
            {
                string text = "编号转换为文字后,现有的项目将无法自动编号.\r\n你确定要继续吗?";
                string caption = "编号转为文字";

                DialogResult dr = MessageBox.Show(text, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr == DialogResult.OK)
                {
                    MSWord.UndoRecord undo = WordApp.UndoRecord;
                    undo.StartCustomRecord("编号转为文字");
                        WordApp.ActiveDocument.Content.ListFormat.ConvertNumbersToText();
                    //WordApp.Selection.Range.ListFormat.ConvertNumbersToText();  对所选内容转换,但用户可能会很困扰.
                    undo.EndCustomRecord();
                    WordApp.StatusBar = "编号已全部转为文字形式!";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnNormal_Click(object sender, RibbonControlEventArgs e)//正文
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为公文:正文");
                SetFont(iniNormal); //三号仿宋
            undo.EndCustomRecord();
            try
            {
                undo.StartCustomRecord("所选文字取消编号和大纲级别");
                    WordApp.Selection.Range.ListFormat.RemoveNumbers(MSWord.WdNumberType.wdNumberAllNumbers);
                    WordApp.Selection.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevelBodyText;  //取消大纲级别
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnTitleHei_Click(object sender, RibbonControlEventArgs e)//公文标题
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为公文:标题");
                SetFont(iniTitle);
            undo.EndCustomRecord();
            try
            {
                undo.StartCustomRecord("设置标题居中");
                    WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter; //居中
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cobFonts_TextChanged(object sender, RibbonControlEventArgs e) //字体列表框设置标题字体
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("所选文字设为二号"+ cobFonts.Text);
                WordApp.Selection.Font.Name = cobFonts.Text;
                WordApp.Selection.Font.Size = 22;
            undo.EndCustomRecord();
            
            try
            {
                undo.StartCustomRecord("设置标题居中");
                    WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter; //居中
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnHead_Click(object sender, RibbonControlEventArgs e)//黑体三 版头
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("设置所选字体");
                SetFont(iniHead);
            undo.EndCustomRecord();
        }

        private void btnName_Click(object sender, RibbonControlEventArgs e)//楷体三 签发人
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("设置所选字体");
               SetFont(iniName);
            undo.EndCustomRecord();
        }

        private void btnFoot_Click(object sender, RibbonControlEventArgs e)//仿宋四
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("设置所选字体");
                SetFont(iniFoot);
            undo.EndCustomRecord();
        }

        private void btnUnderDot_Click(object sender, RibbonControlEventArgs e)// 着重号
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("着重号");
                if (WordApp.Selection.Font.EmphasisMark != MSWord.WdEmphasisMark.wdEmphasisMarkUnderSolidCircle)
                {
                    WordApp.Selection.Font.EmphasisMark = MSWord.WdEmphasisMark.wdEmphasisMarkUnderSolidCircle;
                }
                else
                {
                    WordApp.Selection.Font.EmphasisMark = MSWord.WdEmphasisMark.wdEmphasisMarkNone;
                }
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void btnHideFont_Click(object sender, RibbonControlEventArgs e)//隐藏字体
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("隐藏");
                int hide = WordApp.Selection.Font.Hidden;  //0为不隐藏,-1为隐藏, 9999为混合
                if (hide == -1)
                {
                    WordApp.Selection.Font.Hidden = 0;
                }

                else
                {
                    WordApp.Selection.Font.Hidden = -1;
                }
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //=========== 排版功能区================
        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)//打开页面设置
        {
            try
            {
                WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFilePageSetup].Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnPaper_Click(object sender, RibbonControlEventArgs e)//显示页面设置对话框
        {
            if (wps.Contains("2010")|| wps.Contains("2007"))
            {
                fp = new frmPaperSetup();
                fp.ShowDialog();
                btnPaper.Checked = false;
                return;

            }
         
            ClsWinWrap clsWinWrap = new ClsWinWrap(WordApp.ActiveWindow.Hwnd);  

            if (fp == null || fp.IsDisposed)  //如果没有页面设置对话框,则新建一个
            {
                fp = new frmPaperSetup();
                clsWinWrap.Show(fp);  //指定word是它的主窗体
                // clsWinWrap.ShowDialog(fp);  此方法无法调整ribbin按钮状态
                btnPaper.Checked = true;
            }
            else
            {
                WordApp.StatusBar = "页面设置对话框已打开";
                //MessageBox.Show("fp<>null");
                fp.Activate();
                btnPaper.Checked = true;
            }
        }

        private void btnInsert2_Click(object sender, RibbonControlEventArgs e)//中文缩进. 首行缩进2字符
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("首行缩进2字符");
                    WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnLeft_Click(object sender, RibbonControlEventArgs e) //段落顶格
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("段落顶格");
                WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                WordApp.Selection.ParagraphFormat.FirstLineIndent = 0;
                WordApp.Selection.ParagraphFormat.LeftIndent = 0;
                WordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = 0;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnHeadNoLine_Click(object sender, RibbonControlEventArgs e)//删除页眉横线
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除页眉横线");
                Debug.WriteLine("首页不同:" + WordApp.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter);
                Debug.WriteLine("奇偶页不同:" + WordApp.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter);
                WordApp.ScreenUpdating = false;
                if (WordApp.ActiveWindow.ActivePane.View.SeekView != MSWord.WdSeekView.wdSeekCurrentPageHeader)  //打开当前页面页眉
                {
                    WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
                }
                WordApp.Selection.WholeStory();
                WordApp.Selection.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleNone; //取消边框线


                WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                WordApp.ScreenUpdating = true;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "删除页面横线失败", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                WordApp.ScreenUpdating = true;
            }

        }

        //============================================
        //   排版菜单功能 2021年10月29日全面调整v1900
        //============================================

        private void btnFormatFullLine_Click(object sender, RibbonControlEventArgs e)//断句重排,支持选定或全局
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!","断句重排",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            WordReplace(true, "([!。：！？])^13@", "\\1",s);
            WordApp.ScreenUpdating = true;
        }


        private void btnFormatDelEmptyLine_Click(object sender, RibbonControlEventArgs e) //删除空行
        {
            int n = 0;
            try
            {
                MSWord.Paragraphs p = WordApp.Selection.Paragraphs;   //默认为选择范围

                if (WordApp.Selection.Characters.Count <= 1)  //如果没有选定范围,就视为全局
                { 
                    DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "删除空行", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dr == DialogResult.Cancel) return;
                    p = (MSWord.Paragraphs)WordApp.ActiveDocument.Paragraphs;
                }
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除空行");

                WordApp.ScreenUpdating = false;
                foreach (MSWord.Paragraph i in p)
                {
                    
                    if (i.Range.Text.Trim().Length <= 0 && i.Range.ShapeRange.Count == 0)
                    {
                        //Debug.WriteLine("包含的shapes数量:"+i.Range.ShapeRange.Count); //管用
                        i.Range.Delete();
                        n++;    
                    }
                }

                undo.EndCustomRecord();
                WordApp.StatusBar = "已删除 " + n + " 个空行!";
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "删除空行时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                WordApp.ScreenUpdating = true;
            }
        }

        //已废弃,暂时隐藏,未删除
        //删除选定范围内的空行
        private void btnFormatDelSelectEmptyLine_Click(object sender, RibbonControlEventArgs e) 
        {
            int n = 0;
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除空行");

                MSWord.Paragraphs p = WordApp.Selection.Paragraphs;

                WordApp.ScreenUpdating = false;
                foreach (MSWord.Paragraph i in p)
                {

                    if (i.Range.Text.Trim().Length <= 0 && i.Range.ShapeRange.Count == 0)
                    {
                        //Debug.WriteLine("包含的shapes数量:"+i.Range.ShapeRange.Count); //管用
                        i.Range.Delete();
                        n++;
                    }
                }

                undo.EndCustomRecord();
                WordApp.StatusBar = "已删除 " + n + " 个空行!";

            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "删除空行时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                WordApp.ScreenUpdating = true;

            }
        }



        private void btnFormatDelSpace_Click(object sender, RibbonControlEventArgs e) //删除空白,支持选定范围或全局
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "删除空白", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            WordReplace(false, "^w", "", s); //wps不支持^w的替换,所以改为wps下不显示本按钮　
            WordReplace(false, "　", "", s);　//中文空格
            WordApp.ScreenUpdating = true;
        }

        //换行符转为回车符.支持选定范围或全局
        private void btnFormatR2N_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "换行符转回车", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            WordReplace(false, "^l", "^p",s);
            WordApp.ScreenUpdating = true;
        }

        //删除图形,支持选定范围或全局
        private void btnFormatDelGraphic_Click(object sender, RibbonControlEventArgs e) 
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "删除图形", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            WordReplace(false, "^g", "", s);
            WordApp.ScreenUpdating = true;
        }

        //移走链接,支持选定范围或全局
        private void btnFormatRemoveURL_Click(object sender, RibbonControlEventArgs e) 
        {
            try
            {
                bool s = false;  //是否选定范围布尔值
                if (WordApp.Selection.Characters.Count > 1)
                {
                    s = true;  //如果选定范围,则为选定范围
                }
                else
                {
                    DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "移走链接", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dr == DialogResult.Cancel) return;
                }
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("移走链接");
                WordApp.ScreenUpdating = false;
                if(s) WordApp.Selection.Fields.Unlink();
                else WordApp.ActiveDocument.Fields.Unlink();
                WordApp.ScreenUpdating = true;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //删除半角空格,支持选定范围或全局.
        private void btnFormatDelSelectSpace_Click(object sender, RibbonControlEventArgs e) 
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "删除半角空格", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            WordReplace(false, " ", "", s);
            WordApp.ScreenUpdating = true;
        }

        //中英文之间加空格,支持选定范围或全局.
        private void btnFormatAddSpace_Click(object sender, RibbonControlEventArgs e)
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "中英文之间加空格", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            WordReplace(true, "([A-Za-z])([㑇-﨩])", "\\1 \\2",s);
            WordReplace(true, "([㑇-﨩])([A-Za-z])", "\\1 \\2",s);
            WordApp.ScreenUpdating = true;
        }

        //英文标点符号转为中文,句号和单双引号要单独替换,支持选定范围或全局.
        private void btnFormatE2C_Click(object sender, RibbonControlEventArgs e) 
        {
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围>1,则为选定范围 . 不选中或选中1个字符都视为未选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "标点符号(英转中)", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            //替换句号,但是数字后缀的.不替换.
            WordApp.ScreenUpdating = false;
            string strFind = "([!0-9])\\.([!0-9])";
            string strRep = "\\1。\\2";
            WordReplace(true, strFind, strRep,s);

            //替换其他标点符号
            string[] eArray = { ",", ";", ":", "?", "!", "-", "~", "(", ")", "[", "]" };
            string[] cArray = { "，", "；", "：", "？", "！", "—", "～", "（", "）", "［", "］" };
            if (cArray.Length != eArray.Length) { MessageBox.Show("替换数组有误"); }
            for (int i = 0; i < eArray.Length - 1; i++)
            {
                WordReplace(false, eArray[i], cArray[i],s);
            }
            //替换双引号
            if (wps.Contains("WPS"))  
            {
                //wps下,不支持\1的替换语句
                //2次替换法, 新版wps将中英文双引号视为一种符号,无法折中解决.
            }
            else
            {
                strFind = "\"(*)\"";
                strRep = "“\\1”";
                WordReplace(true, strFind, strRep, s);
            }
            
            WordApp.ScreenUpdating = true;
        }

        //中文标点符号转为英文,数组中会多句号和单双引号.,支持选定范围或全局.
        private void btnFormatC2E_Click(object sender, RibbonControlEventArgs e) 
        {
            //关闭引号自动更正提示
            if (WordApp.Options.AutoFormatAsYouTypeReplaceQuotes)
            {
                string text = "Word自动更正选项中,已开启输入直引号自动更正为弯引号." +
                    "\r\n此功能将影响中文引号的转换.要关闭此选项并开始转换吗?" +
                    "\r\n\r\n点击\"是\"将关闭该自动更正选项并转换.\r\n点击\"否\"将直接转换,引号可能被Word自动更正." +
                    "\r\n点击\"取消\"退出该操作.";
                string caption = "中文标点符号转为英文";
                DialogResult dr = MessageBox.Show(text, caption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    WordApp.Options.AutoFormatAsYouTypeReplaceQuotes = false;
                }
                else if (dr == DialogResult.Cancel)
                {
                    return;
                }
            }
            bool s = false;  //是否选定范围布尔值
            if (WordApp.Selection.Characters.Count > 1)
            {
                s = true;  //如果选定范围,则为选定范围
            }
            else
            {
                DialogResult dr = MessageBox.Show("您未选定范围,将执行全局操作!", "标点符号(中转英)", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.Cancel) return;
            }
            WordApp.ScreenUpdating = false;
            string[] cArray = { "。", "，", "；", "：", "？", "！", "—", "～", "（", "）", "［", "］", "“", "”", "‘", "’" };
            string[] eArray = { ".", ",", ";", ":", "?", "!", "-", "~", "(", ")", "[", "]", "\"", "\"", "'", "'" };
            if (cArray.Length != eArray.Length) { MessageBox.Show("替换数组有误");return; }
            for (int i = 0; i < eArray.Length - 1; i++)
            {
                WordReplace(false, cArray[i], eArray[i],s);
            }
            WordApp.ScreenUpdating = true;
        }

        //调用将单个文本框转换为文字功能
        private void btnFormatFrame2txt_Click(object sender, RibbonControlEventArgs e) 
        {
            Frame2txt();
        }

        //段前分页, 即所选段落另起一页
        private void btnFormatParaPage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                WordApp.Run("ParaPageBreakBefore");
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "所选段落另起一页", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
    }








        //============插入功能区================
        private void btnDate_Click(object sender, RibbonControlEventArgs e)//插入中文日期
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入日期");
            SetFont(iniNormal); //同正文字体
            try
            {
                WordApp.Selection.InsertDateTime("YYYY年M月D日", false);
                WordApp.Selection.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                undo.EndCustomRecord();
            }
        }

        private void btnLine1_Click(object sender, RibbonControlEventArgs e) //插入红色反线
        {
            try
            {
                WordApp.ScreenUpdating = false;
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("插入间隔线");
                
                WordApp.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                WordApp.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineWidth = MSWord.WdLineWidth.wdLineWidth225pt;
                undo.EndCustomRecord();

                undo.StartCustomRecord("间隔线设为红色");
                WordApp.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].Color = MSWord.WdColor.wdColorRed;
                
                undo.EndCustomRecord();
                WordApp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnFootSingle_Click(object sender, RibbonControlEventArgs e) //插入版记线单行
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入单行版记");
            InsertLine();
            undo.EndCustomRecord();
        }

        private void btnFootDouble_Click(object sender, RibbonControlEventArgs e) //插入版记线双行
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入双行版记");
            InsertLine(2);
            undo.EndCustomRecord();
        }
        


        private void btnInsertPaper_Click(object sender, RibbonControlEventArgs e) //插入横页
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("插入横页");

                WordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;

                WordApp.Selection.InsertBreak(MSWord.WdBreakType.wdSectionBreakNextPage);//插入分节符
                WordApp.Selection.PageSetup.Orientation = MSWord.WdOrientation.wdOrientLandscape;//将当前页面调整为横向
                undo.EndCustomRecord();

                undo.StartCustomRecord("插入竖页");
                WordApp.Selection.InsertBreak(MSWord.WdBreakType.wdSectionBreakNextPage);//再插入分节符
                WordApp.Selection.PageSetup.Orientation = MSWord.WdOrientation.wdOrientPortrait;//将当前页面调整为纵向
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "插入横页失败", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void galSymbol_Click(object sender, RibbonControlEventArgs e)//插入符号
        {
            try
            {
                string f = glySymbol.SelectedItem.Label;

                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("插入符号:" +f);
                    WordApp.Selection.TypeText(f);
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnSymbol_zkh_Click(object sender, RibbonControlEventArgs e)//[]替换年份的六角括号
        {
            WordReplace(true, @"\[([0-9]{1,4})\]", @"〔\1〕");
            WordReplace(true, "［([0-9]{1,4})］", @"〔\1〕");

        }
        private void btnSymbol_hkh_Click(object sender, RibbonControlEventArgs e) //【】替换为六角括号
        {

            WordReplace(true, "【([0-9]{1,4})】", @"〔\1〕");

        }

        private void btnNumberCenter_Click(object sender, RibbonControlEventArgs e) //居中页码按钮
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入居中页码");
                PageNumberAligment(0);
            undo.EndCustomRecord();
        }

        private void BtnNumberRight_Click(object sender, RibbonControlEventArgs e) //居右页码按钮
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入右侧页码");
                PageNumberAligment(2);
            undo.EndCustomRecord();
        }

        private void btnNumberOutside_Click(object sender, RibbonControlEventArgs e) //外侧页码按钮
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入外侧页码");
                PageNumberAligment(1);
            undo.EndCustomRecord();
        }

        private void BtnNumberOnLocation_Click(object sender, RibbonControlEventArgs e) //光标处插入页码按钮
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("光标处插入页码");
                InsertNumber();
            undo.EndCustomRecord();
        }


        private void btnInsertSignPicLink_Click(object sender, RibbonControlEventArgs e)//以链接方式插入印章
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("以链接方式插入印章");
                btnInsertSignPic.Tag = "Link";
                btnInsertSignPic.SuperTip = btnInsertSignPicLink.SuperTip;
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "请选择以链接方式插入的印章";
                ofd.Filter = "支持透明度的图片(*.PNG,*.gif,*.TIFF)|*.png;*.gif;*.tiff|所有文件(*.*)|*.*";
                DialogResult dr = ofd.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    string filename = ofd.FileName;
                    InsertSignPic(filename);
                }
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开页面设置对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        private void btnInsertSignPicSave_Click(object sender, RibbonControlEventArgs e)//以嵌入方式插入印章
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("以嵌入方式插入印章");
                btnInsertSignPic.Tag = "";
                btnInsertSignPic.SuperTip = btnInsertSignPicSave.SuperTip;
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "请选择以保存方式插入的印章";
                ofd.Filter = "支持透明度的图片(*.PNG,*.gif,*.TIFF)|*.png;*.gif;*.tiff|所有文件(*.*)|*.*";
                DialogResult dr = ofd.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    string filename = ofd.FileName;
                    InsertSignPic(filename, false);
                }
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开页面设置对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    

        private void btnInsertSignPic_Click(object sender, RibbonControlEventArgs e) //插入印章
        {
            if ((string)btnInsertSignPic.Tag == "Link")
            {
                btnInsertSignPicLink_Click(sender, e);
            }
            else
            {
                btnInsertSignPicSave_Click(sender, e);
            }
        }


        //============其他和关于功能区=============


        private void btnXkonglong_Click(object sender, RibbonControlEventArgs e) //访问website
        {
            try
            {
                Process.Start("https://gw.xkonglong.com");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n请打开浏览器输入网址访问:\r\n https://gw.xkonglong.com");
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)//关于对话框
        {
            AboutBox1 about = new AboutBox1();
            DialogResult dr = about.ShowDialog();
            if (dr==DialogResult.OK) ReadSetting();
        }

        private void btnEmail_Click(object sender, RibbonControlEventArgs e) //发送反馈
        {
            try
            {
                string aVer = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(); //程序集版本
                System.Version oVer = System.Environment.OSVersion.Version; //OS version
                string appVer = System.Windows.Forms.Application.ProductName + System.Windows.Forms.Application.ProductVersion;         ;//获取word或wps的宿主程序版本  
                string body = "%0d%0a%0d%0a以下我的系统环境信息,供您参考."+"\r\n当前 Windows Ver:" + oVer.ToString()+ "\r\n当前程序版本 ver:"+appVer;

                Process.Start("mailto:xkonglong@qq.com?subject=关于小恐龙公文排版助手v" + aVer + "的意见反馈&body=" + body);

                //string msg = "反馈程序错误时,您可通过“关于插件”复制相关系统信息,\r\n以便我能更好的排查和解决问题。\r\n点击确定将打开反馈网页\r\nhttps://support.qq.com/products/57503";
                //MessageBox.Show(msg, "小恐龙公文排版助手",MessageBoxButtons.OK,MessageBoxIcon.Information);
                //Process.Start("https://support.qq.com/products/57503");
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "为发送反馈获取系统信息时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void btnHelp_Click(object sender, RibbonControlEventArgs e)//使用帮助
        {
            try
            {
                //Process.Start("http://xkonglong.com/wordadd_help/");
                Process.Start(Properties.Settings.Default.在线帮助);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+ "\r\n请打开浏览器输入网址访问:\r\n http://xkonglong.com/wordadd_help/");
            }
        }


        //副本功能区

        private void btnSaveAsPDF_Click(object sender, RibbonControlEventArgs e)//生成pdf副本
        {
            string outFile = "";

            //如果未存档,先提示存档.调用word的保存功能.如果取消保存则退出.
            if (WordApp.ActiveDocument.Path.Length==0) 
            {
                DialogResult dr = MessageBox.Show("生成副本前请先保存本文档,要保存文档吗?", 
                    "生成PDF副本提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    try
                    {
                        WordApp.ActiveDocument.Save();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("由于文档未保存成功,副本也未保存.", ex.Message,MessageBoxButtons.OK,MessageBoxIcon.Stop);
                        return;
                    }
                }
                else { return; }
            }
            //尝试生成pdf副本操作
            try
            {
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                fn = Path.GetFileNameWithoutExtension(fn); //不含扩展名的文件名
                //MessageBox.Show(ex);
                //string[] nfn = fn.Split(".".ToCharArray());
                
                outFile = fp + "\\" + fn + ".pdf";  //outFile= 输出副本的名称

                //如果副本存在同名文件,则提示
                if (System.IO.File.Exists(outFile)) 
                {
                    DialogResult drr = MessageBox.Show("文件 \r\n" + outFile + "\r\n已存在! " +
                        "\r\n点击\"是\"将覆盖同名文件.\r\n点击\"否\"将打开保存对话框更改路径或文件名.\r\n点击\"取消\"将取消生成副本.",
                        "生成副本时发现同名文件", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (drr==DialogResult.No)  //为否,提示改名
                    {
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.AddExtension = true;
                        sfd.DefaultExt = "pdf";
                        sfd.Filter = "PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
                        sfd.InitialDirectory = WordApp.ActiveDocument.Path;
                        sfd.FileName = outFile;
                        DialogResult dr = sfd.ShowDialog();
                        if (dr == DialogResult.OK)
                        {
                            SavePdf(sfd.FileName); //新文件名保存
                        }
                    }
                    else if(drr==DialogResult.Yes) //是,继续覆盖
                    {
                        SavePdf(outFile);
                    }
                    else  //取消退出
                    {
                        return;
                    }
                }
                else  //否则直接保存
                {
                    SavePdf(outFile);
                }

            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "生成pdf副本时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    
        private void btnSaveAsDoc_Click(object sender, RibbonControlEventArgs e)//生成doc副本
        {
            string outFile = "";

            //如果未存档,先提示存档.调用word的保存功能.如果取消保存则退出.
            if (WordApp.ActiveDocument.Path.Length==0)
            {
                DialogResult dr = MessageBox.Show("生成副本前请先保存本文档,要保存文档吗?", "生成DOC副本提示",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    try
                    {
                        WordApp.ActiveDocument.Save();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message+ "\r\n\r\n由于文档未保存,生成副本操作被取消.", "生成副本失败",
                            MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }
                else { return; }
            }

            //如果文件已存档,则尝试生成副本
            try
            {
                //获取文档路径和名称
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                

                string oldFile = fp + "\\" + fn;     //oldFile=原始文档名称
                //string[] nfn = fn.Split(".".ToCharArray());
                //fn = nfn[0];
                fn = Path.GetFileNameWithoutExtension(fn);  //不含扩展名的文件名

                outFile = fp + "\\" + fn + ".doc";  //outFile=副本文档名称

                if (System.IO.File.Exists(outFile))//如果发现同名副本,提示是否继续.
                {
                    DialogResult drr = MessageBox.Show("文件 \r\n" + outFile + "\r\n已存在! " +
                        "\r\n点击\"是\"将覆盖同名文件.\r\n点击\"否\"将打开保存对话框更改路径或文件名.\r\n点击\"取消\"将取消生成副本.",
                        "生成副本时发现同名文件", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    //选择了否,将改名保存.
                    if (drr == DialogResult.No) 
                    {
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.AddExtension = true;
                        sfd.DefaultExt = "doc";
                        sfd.Filter = "Word97-2003格式文档(*.doc)|*.doc|所有文件 (*.*)|*.*";
                        sfd.InitialDirectory = WordApp.ActiveDocument.Path;
                        sfd.FileName = outFile;
                        DialogResult dr = sfd.ShowDialog();
                        
                        if (dr == DialogResult.OK)
                        {
                            SaveDoc(sfd.FileName,oldFile);//新文件名保存
                        }
                    }
                    //选择了是,将直接覆盖.
                    else if(drr==DialogResult.Yes)
                    {
                        SaveDoc(outFile, oldFile);
                    }
                    //选择了取消,取消操作
                    else if(drr==DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else   //未发现同名副本
                {
                    SaveDoc(outFile, oldFile);
                }
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "生成doc副本时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
            }
        }
        private void btnSaveAsPic_Click(object sender, RibbonControlEventArgs e)//文字转png图片
        {
            try
            {
                //如果未存档,先提示存档.调用word的保存功能.如果取消保存则退出.
                if (WordApp.ActiveDocument.Path.Length==0)
                {
                    DialogResult dr = MessageBox.Show("生成副本前请先保存本文档,要保存文档吗?", "生成DOC副本提示",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        try
                        {
                            WordApp.ActiveDocument.Save();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\r\n\r\n由于文档未保存,生成副本操作被取消.", "生成副本失败",
                                MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                    }
                    else { return; }
                }

                //如果文件已存档,则尝试生成副本,因为涉及到多张图片的可能性,所以直接覆盖不提示.
                //获取文档路径和名称
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                string oldFile = fp + "\\" + fn;     //oldFile=文档完整路径和名称
                MSWord.Pages pages = WordApp.ActiveDocument.ActiveWindow.ActivePane.Pages; //获取页面集合
                
                for (int i = 1; i <= pages.Count; i++)
                {
                    WordApp.StatusBar = "正在将页面转换成图片 " + i+"...";
                    MSWord.Page page = pages[i];
                    System.IO.Stream st = new System.IO.MemoryStream((byte[])(page.EnhMetaFileBits));
                    Image img = Image.FromStream(st);
                    st.Dispose();//回收资源,避免连续存图崩溃.
                    Bitmap bmp = new Bitmap(img, 1240, 1754);
                    if (img.Width > img.Height) //如果是横页
                    {
                        bmp = new Bitmap(img, 1754, 1240);
                    }
                    img.Dispose();
                    bmp.Save(String.Format(oldFile + "_{0}.png", i), System.Drawing.Imaging.ImageFormat.Png);
                    bmp.Dispose();
                    
                }
                WordApp.StatusBar = "文字转图片已完成,请在打开的文件夹中查看";
                Process.Start(fp); //打开所在文件夹

            }
            catch (Exception ex)
            {
               // Globals.ThisAddIn.Logging(ex.Source, ex.Message);
                WordApp.StatusBar = ex.Message;
            }
        }
        
        private void BtnSaveAsJPG_Click(object sender, RibbonControlEventArgs e)  //保存图片副本为jpg格式
        {
            float dpi = 150;
            try
            {
                dpi = Properties.Settings.Default.DPI;
            }
            catch { }
            try
            {
                //如果未存档,先提示存档.调用word的保存功能.如果取消保存则退出.
                if (WordApp.ActiveDocument.Path.Length==0)
                {
                    DialogResult dr = MessageBox.Show("生成副本前请先保存本文档,要保存文档吗?", "生成DOC副本提示",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        try
                        {
                            WordApp.ActiveDocument.Save();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\r\n\r\n由于文档未保存,生成副本操作被取消.", "生成副本失败",
                                MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                    }
                    else { return; }
                }

                //如果文件已存档,则尝试生成副本,因为涉及到多张图片的可能性,所以直接覆盖不提示.
                //获取文档路径和名称
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                string oldFile = fp + "\\" + fn;     //oldFile=文档完整路径和名称
                MSWord.Pages pages = WordApp.ActiveDocument.ActiveWindow.ActivePane.Pages; //获取页面集合

                for (int i = 1; i <= pages.Count; i++)
                {
                    WordApp.StatusBar = "正在将页面转换成图片 " + i + "...";
                    MSWord.Page page = pages[i];
                    using (Stream st = new MemoryStream((byte[])(page.EnhMetaFileBits)))
                    using (Image img = Image.FromStream(st))
                    {
                        int w, h;
                        if (img.Width > img.Height) //如果是横页
                        {
                            w = 1754; h = 1240;
                        }
                        else
                        {
                            w = 1240; h = 1754;
                        }
                        using (Bitmap b = new Bitmap(w, h))
                        {
                            //b.SetResolution(150, 150);
                            //b.SetResolution(300, 300); //提高精度,2019.12.2
                            b.SetResolution(dpi, dpi);
                            using (Graphics g = Graphics.FromImage(b))
                            {
                                g.Clear(Color.White);
                                //g.DrawImageUnscaled(img, 0, 0);
                                g.DrawImage(img, 0, 0, w, h);
                            }
                            b.Save(String.Format(oldFile + "_{0}.jpg", i), System.Drawing.Imaging.ImageFormat.Jpeg);
                        }

                    }
                }
                WordApp.StatusBar = "文字转图片已完成,请在打开的文件夹中查看";
                Process.Start(fp); //打开所在文件夹

            }
            catch (Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
                WordApp.StatusBar = ex.Message;
            }

        }
        private void btnSaveAsPicAt_Click(object sender, RibbonControlEventArgs e)//保存当前页面为图片
        {
            try
            
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.AddExtension = true;
                sfd.Title = "将当前页面保存为图片,您可以根据需求选择图片格式.";
                sfd.FileName = DateTime.Now.ToString("yyyyMMdd-hhmmss");
                sfd.DefaultExt = ".png|*.jpg|*.tiff";
                sfd.Filter = "A4尺寸网络交换图片(*.png)|*.png|原始尺寸网络交换图片(*.png)|*.png|原始尺寸有损压缩图片(*.jpg)|*.jpg|原始尺寸高精度文档图片(*.tiff)|*.tiff|所有文件(*.*)|*.*";
                DialogResult dr = sfd.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    WordApp.StatusBar = "正在将当前页面保存成图片... ";
                    int i = WordApp.Selection.Information[MSWord.WdInformation.wdActiveEndPageNumber];
                    MSWord.Page page = WordApp.ActiveDocument.ActiveWindow.ActivePane.Pages[i];    //获取当前页面

                    Stream sr = new MemoryStream((byte[])(page.EnhMetaFileBits));
                    Image img = Image.FromStream(sr);
                    sr.Dispose();  //销毁回收内存,避免连续存图导致内存溢出,word崩溃.

                    switch (sfd.FilterIndex)
                    {
                        case 1:
                            Bitmap png;
                            if (img.Width > img.Height) //如果是横页
                            {
                                png = new Bitmap(img, 1754, 1240);
                            }
                            else
                            {
                                 png= new Bitmap(img, 1240, 1754); //a4尺寸图片,竖向.
                            }

                            png.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Png);
                            png.Dispose();//销毁回收内存,避免连续存图导致内存溢出,word崩溃.
                            break;

                        case 2:
                            img.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Png);
                            img.Dispose();
                            break;

                        case 3:
                            //使用using,确保变量释放,避免内存溢出
                            using (Bitmap b = new Bitmap(img.Width, img.Height))
                            {
                                b.SetResolution(img.HorizontalResolution, img.VerticalResolution);

                                using (Graphics g = Graphics.FromImage(b))
                                {
                                    g.Clear(Color.White);
                                    g.DrawImageUnscaled(img, 0, 0);
                                }
                                b.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                            }
                            break;

                        case 4:
                            img.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Tiff);
                            img.Dispose();//销毁回收内存,避免连续存图导致内存溢出,word崩溃.
                            break;
                    }
                    WordApp.StatusBar = "页面保存图片成功!";
                    Process.Start(sfd.FileName); //打开图片.

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                WordApp.StatusBar = "页面保存图片失败.";
            }
            

        }
        private void btnProtectDoc_Click(object sender, RibbonControlEventArgs e)//限制编辑,保护文档
        {
            try
            {
                WordApp.TaskPanes[MSWord.WdTaskPanes.wdTaskPaneDocumentProtection].Visible = true;  //显示限制编辑面板.
                if (WordApp.ActiveDocument.ProtectionType == MSWord.WdProtectionType.wdNoProtection)  //如果未保护状态,显示加密对话框
                {
                    WordApp.Dialogs[MSWord.WdWordDialog.wdDialogToolsProtectDocument].Show();
                    btnProtectDoc.Checked = false;
                    return;
                }
                else
                {
                    btnProtectDoc.Checked = false;
                }
            }
            catch (Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
                WordApp.StatusBar = ex.Message;
            }

        }

        private void DocPassWord(string passWord)  //设置文档打开密码
        {
            try
            { 
                WordApp.ActiveDocument.Password = passWord;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSetPW_Click(object sender, RibbonControlEventArgs e) //设置密码
        {
            try
            {
                WordApp.Dialogs[MSWord.WdWordDialog.wdDialogToolsOptionsSecurity].Show();
                
            }
            catch(Exception ex)
            {
                if (wps.Contains("WPS"))
                {
                    MessageBox.Show("对不起,您的WPS暂时还不兼容这一调用内置对话框的方法!","打开内置对话框时出错");
                    btnSetPW.Enabled = false;
                }
                MessageBox.Show(ex.Message);
                
            }
            

        }

        private void btnClaerPW_Click(object sender, RibbonControlEventArgs e)//清除打开密码
        {
            DialogResult dr = MessageBox.Show("你确定要清除本文档的\"打开密码\"吗?\r\n 点击确定将直接清除密码.", 
                "文档加密", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (dr == DialogResult.OK)
            {
                DocPassWord("");
                WordApp.StatusBar = "警告:文档打开密码已被清除!";
            }

        }

        private void btnClearProtectPW_Click(object sender, RibbonControlEventArgs e) //清除修改密码
        {
            DialogResult dr = MessageBox.Show("你确定要清除本文档的\"修改密码\"吗?\r\n 点击确定将直接清除密码.",
               "文档加密", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (dr == DialogResult.OK)
            {
                WordApp.ActiveDocument.WritePassword = "";
                WordApp.StatusBar = "警告:文档修改密码已被清除!";
            }

        }

        private void Sp_SpeakCompleted(object sender, System.Speech.Synthesis.SpeakCompletedEventArgs e)//朗读完成事件
        {
            this.btnSpeak.OfficeImageId = "SpeakCells";
            this.btnSpeak.Tag = "";
            this.btnSpeak.Label = "朗读";
            WordApp.StatusBar = "朗读完成!";
            this.btnSpeak.Checked = false;
            this.btnSpeak.SuperTip = "单击朗读,再次单击停止.";
            //throw new NotImplementedException();
        }
        private void BtnSpeakOption_Click(object sender, RibbonControlEventArgs e)  //朗读选项
        {

            ClsWinWrap clsWinWrap = new ClsWinWrap(WordApp.ActiveWindow.Hwnd);
            clsWinWrap.Show(new frmSpeakOption());  //打开窗口，指定它的父亲是Word
        }

        private void btnSpeak_Click(object sender, RibbonControlEventArgs e) //朗读按钮
        {
            //WordApp.Run("SpeakStopSpeaking");  //宏命令无效
            try
            {
                if (this.btnSpeak.Tag.ToString().Length==0)
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
                        //WordApp.Selection.WholeStory();
                        //WordApp.Selection.ExtendMode = true;
                        //WordApp.Selection.Extend();//单词
                        //WordApp.Selection.Extend();//句子
                        //WordApp.Selection.Extend();//段落
                        WordApp.Selection.MoveEnd(MSWord.WdUnits.wdParagraph);
                        //WordApp.Selection.ExtendMode = false;
                        sp.SpeakAsync(WordApp.Selection.Text);
                    }


                    return;
                }
                else
                {
                    
                    sp.SpeakAsyncCancelAll();
                    this.btnSpeak.Tag = "";
                    this.btnSpeak.Label = "朗读";
                    this.btnSpeak.Checked = false;
                    this.btnSpeak.OfficeImageId = "SpeakCells";
                    this.btnSpeak.SuperTip = "单击朗读,再次单击停止.";
                    return;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// 清理用户自定义的样式
        /// </summary>
        private void clearCustomStyles() 
        {
            try
            {
                foreach (MSWord.Style sy in WordApp.ActiveDocument.Styles)
                {
                    if (sy.BuiltIn == false)
                        sy.Delete();
                }
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "清理自定义样式时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        /// <summary>
        /// 新建用户自定义样式
        /// </summary>
        private void creatCustomStyle()
        {
            try
            {
                //检查是否存在公文风格,th=公文标题,t=正文,ts=缩进正文,t1~t4=1-4标题

                bool t4, t3, t2, t1, t, th,ts,fj;
                t4 = t3 = t2 = t1 = t = th=ts =fj= false;
                foreach (MSWord.Style name in WordApp.ActiveDocument.Styles)
                {
                    switch (name.NameLocal)
                    {
                        case "公文:4级标题":
                            t4 = true;
                            break;
                        case "公文:3级标题":
                            t3 = true;
                            break;
                        case "公文:2级标题":
                            t2 = true;
                            break;
                        case "公文:1级标题":
                            t1 = true;
                            break;
                        case "公文:标题":
                            th = true;
                            break;
                        case "公文:正文":
                            t = true;
                            break;
                        case "公文:正文(缩进)":
                            ts = true;
                            break;

                    }
                }
                if ((t4 & t3 & t2 & t1 & t & th & ts)==true)//如果存在公文样式,退出.
                {
                    Debug.WriteLine("公文样式已存在!");
                    return;
                }

                MSWord.ListTemplate myTemplate;    //定义list模板
                MSWord.Style syT4, syT3, syT2, syT1, sy, syTh,sys;     //定义风格
                //如果不存在,就新建公文风格
                if (!t4)
                {
                    syT4 = WordApp.ActiveDocument.Styles.Add("公文:4级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    //syT4.AutomaticallyUpdate = true;   //是否自动更新
                    syT4.set_BaseStyle(""); //样式基于无格式样式.
                    syT4.Font.NameFarEast = iniTitle4.Name;   //公文字体
                    syT4.Font.Size =iniTitle4.SizeInPoints;    //公文字号
                    if (iniTitle4.Bold) syT4.Font.Bold = 1;   //粗体
                    if (iniTitle4.Italic) syT4.Font.Italic = 1;   //斜体


                    syT4.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel4; //4级大纲级别

                    //设置编号
                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add(); 
                    myTemplate.ListLevels[1].NumberFormat = "（%1）";//（1）
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;//编号与文字之间无空格无Tab
                    myTemplate.ListLevels[1].TextPosition = 0; //设置第二行缩进量
                    syT4.LinkToListTemplate(myTemplate);  //公文风格链接列表模板
                    myTemplate = null;

                }



                if (!t3)
                {
                    syT3 = WordApp.ActiveDocument.Styles.Add("公文:3级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT3.set_BaseStyle(""); //样式基于无格式样式.
                    syT3.Font.NameFarEast = iniTitle3.Name;  //设置字体
                    syT3.Font.Size = iniTitle3.SizeInPoints; ;
                    if (iniTitle3.Bold) syT3.Font.Bold = 1;   //粗体
                    if (iniTitle3.Italic) syT3.Font.Italic = 1;   //斜体


                    syT3.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel3; //3级大纲级别
                    
                    //设置编号
                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberFormat = "%1．";//1．
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;//编号与文字之间无空格无Tab
                    myTemplate.ListLevels[1].TextPosition = 0; //设置第二行缩进量
                    syT3.LinkToListTemplate(myTemplate);  //公文风格链接列表模板
                    myTemplate = null;


                }

                if (!t2)
                {
                    syT2 = WordApp.ActiveDocument.Styles.Add("公文:2级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT2.set_BaseStyle(""); //样式基于无格式样式.
                    syT2.Font.NameFarEast = iniTitle2.Name;
                    syT2.Font.Size = iniTitle2.SizeInPoints;
                    if (iniTitle2.Bold) syT2.Font.Bold = 1;   //粗体
                    if (iniTitle2.Italic) syT2.Font.Italic = 1;   //斜体


                    syT2.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel2; //2级大纲级别

                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                    myTemplate.ListLevels[1].NumberFormat = "（%1）";//（一）
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone; //编号与文字之间无空格无Tab
                    myTemplate.ListLevels[1].TextPosition = 0; //设置第二行缩进量
                    
                    syT2.LinkToListTemplate(myTemplate);
                    myTemplate = null;

    
                }
                if (!t1)
                {
                    syT1 = WordApp.ActiveDocument.Styles.Add("公文:1级标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syT1.set_BaseStyle(""); //样式基于无格式样式.
                    syT1.Font.NameFarEast = iniTitle1.Name;
                    syT1.Font.Size = iniTitle1.SizeInPoints;
                    if (iniTitle1.Bold) syT1.Font.Bold = 1;   //粗体
                    if (iniTitle1.Italic) syT1.Font.Italic = 1;   //斜体


                    syT1.ParagraphFormat.OutlineLevel = MSWord.WdOutlineLevel.wdOutlineLevel1; //一级大纲级别

                    myTemplate = WordApp.ActiveDocument.ListTemplates.Add();
                    myTemplate.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
                    myTemplate.ListLevels[1].NumberFormat = "%1、";//一、
                    myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone; //编号与文字之间无空格无Tab
                    myTemplate.ListLevels[1].TextPosition = 0; //设置第二行缩进量
                    syT1.LinkToListTemplate(myTemplate);
                    myTemplate = null;



                }
                if (!t)
                {

                    sy = WordApp.ActiveDocument.Styles.Add("公文:正文", MSWord.WdStyleType.wdStyleTypeParagraph);
                    sy.set_BaseStyle(""); //样式基于无格式样式.
                    sy.Font.NameFarEast = iniNormal.Name;
                    sy.Font.Size = iniNormal.SizeInPoints;
                    if (iniNormal.Bold) sy.Font.Bold = 1;   //粗体
                    if (iniNormal.Italic) sy.Font.Italic = 1;   //斜体
                    sy.set_BaseStyle(MSWord.WdBuiltinStyle.wdStyleNormal);//基于样式为无格式

                    //指定快捷键 alt+shift+z
                    int key = WordApp.BuildKeyCode(MSWord.WdKey.wdKeyAlt, MSWord.WdKey.wdKeyShift, MSWord.WdKey.wdKeyZ);
                    WordApp.KeyBindings.Add(MSWord.WdKeyCategory.wdKeyCategoryStyle, "公文:正文", key); 

                }
                if (!ts)
                {
                    sys = WordApp.ActiveDocument.Styles.Add("公文:正文(缩进)", MSWord.WdStyleType.wdStyleTypeParagraph);
                    sys.set_BaseStyle(""); //样式基于无格式样式.
                    sys.Font.NameFarEast = iniNormal.Name;
                    sys.Font.Size = iniNormal.SizeInPoints ;
                    if (iniNormal.Bold) sys.Font.Bold = 1;   //粗体
                    if (iniNormal.Italic) sys.Font.Italic = 1;   //斜体
                    sys.set_BaseStyle(MSWord.WdBuiltinStyle.wdStyleNormal);//基于样式为无格式
                    sys.ParagraphFormat.CharacterUnitFirstLineIndent = 2;  //首行缩进2字符

                    //指定快捷键 alt+shift+s
                    int key = WordApp.BuildKeyCode(MSWord.WdKey.wdKeyAlt, MSWord.WdKey.wdKeyShift, MSWord.WdKey.wdKeyS);
                    WordApp.KeyBindings.Add(MSWord.WdKeyCategory.wdKeyCategoryStyle, "公文:正文(缩进)", key);

                }
                if (!th)
                {
                    syTh = WordApp.ActiveDocument.Styles.Add("公文:标题", MSWord.WdStyleType.wdStyleTypeParagraph);
                    syTh.set_BaseStyle(""); //样式基于无格式样式.

                    syTh.Font.NameFarEast = iniTitle.Name;

                    syTh.Font.Size = iniTitle.SizeInPoints;
                    if (iniTitle.Bold) syTh.Font.Bold = 1;   //粗体
                    if (iniTitle.Italic) syTh.Font.Italic = 1;   //斜体
                    syTh.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;   //居中
                    syTh.set_BaseStyle(MSWord.WdBuiltinStyle.wdStyleNormal);//基于样式为无格式
                }

                
            }
            catch (Exception ex)
            {
                string err = ex.Message + "比如指定了不存在的字体或者在英文环境使用了!请到关于-高级设置里先指定公文字体。\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "新建公文样式时发生错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void btnStyle_Click(object sender, RibbonControlEventArgs e)//生成样式按钮
        {
            //clearCustomStyles(); //清理自定义样式

            Cursor.Current = Cursors.WaitCursor;
            creatCustomStyle();
         
            WordApp.TaskPanes[MSWord.WdTaskPanes.wdTaskPaneFormatting].Visible = true; //显示样式任务面板
            WordApp.ShowStylePreviews = true;
            Cursor.Current = Cursors.Default; 

            
        }
        private void btnStyleR_Click(object sender, RibbonControlEventArgs e)//强制重建样式按钮
        {
            Cursor.Current = Cursors.WaitCursor;
            WordApp.Selection.ClearFormatting();
            clearCustomStyles(); //清理自定义样式
            creatCustomStyle();  //创建样式
            WordApp.TaskPanes[MSWord.WdTaskPanes.wdTaskPaneFormatting].Visible = true; //显示样式任务面板
            WordApp.ShowStylePreviews = true;
            Cursor.Current = Cursors.Default;
        }

        private void btnFontFSgb2312_Click(object sender, RibbonControlEventArgs e)//仿宋GB2312=>仿宋
        {
            FontSwitch("仿宋_GB2312", "仿宋");
        }

        private void btnFontKTgb2312_Click(object sender, RibbonControlEventArgs e)//楷体_gb2312=>楷体
        {
            FontSwitch("楷体_GB2312","楷体");
        }

        private void btnFontKT_Click(object sender, RibbonControlEventArgs e)//楷体=>楷体_gb2312
        {
            FontSwitch("楷体","楷体_GB2312" );
        }

        private void btnFontFS_Click(object sender, RibbonControlEventArgs e)//仿宋=>仿宋GB2312
        {
            FontSwitch("仿宋","仿宋_GB2312" );
        }

        private void btnFontBS2HT_Click(object sender, RibbonControlEventArgs e) //方正小标宋=>黑体
        {
            FontSwitch("方正小标宋_GBK", "黑体");
            FontSwitch("方正小标宋简体", "黑体");
        }

        private void btnFontHT2BS_Click(object sender, RibbonControlEventArgs e)//黑体=>方正小标宋_GBK
        {
            FontSwitch( "黑体", "方正小标宋_GBK");
        }

        private void group6_DialogLauncherClick(object sender, RibbonControlEventArgs e)//其他标签角标功能
        {
            try
            {
                WordApp.Dialogs[MSWord.WdWordDialog.wdDialogFileSummaryInfo].Show();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开Word内置对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void group5_DialogLauncherClick(object sender, RibbonControlEventArgs e)//关于标签角标功能
        {
            try
            {
                if (wps.Contains("WPS"))
                {
                    WordApp.ListCommands(false);
                }
                else
                {
                    WordApp.Dialogs[MSWord.WdWordDialog.wdDialogListCommands].Show();// 新建一篇命令列表文档.
                }
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开Word内置对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void BtnSymbol_index_Click(object sender, RibbonControlEventArgs e)  //数字转带圈序号
        {
            int at=1;
            string number = WordApp.Selection.Text;

            if (int.TryParse(number,out at))
            {
                Debug.WriteLine("数字=" + number);
                if (at > 20)
                {
                    MessageBox.Show("带圈序号最大只能到20","数字转换为带圈序号", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("带圈数字序列");

            MSWord.ListTemplate myTemplate2 = WordApp.ActiveDocument.ListTemplates.Add();  //添加法
            myTemplate2.ListLevels[1].NumberFormat = "%1";//①
            myTemplate2.ListLevels[1].NumberStyle = MSWord.WdListNumberStyle.wdListNumberStyleNumberInCircle;//带圈数字
            myTemplate2.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;//编号与文字之间无空格无Tab
            myTemplate2.ListLevels[1].TextPosition = 0; //设置第二行缩进量
            myTemplate2.ListLevels[1].NumberPosition = 0;

            
            if (at <= 0) at = 1; myTemplate2.ListLevels[1].StartAt = at;//设置序列
            WordApp.Selection.Range.ListFormat.ApplyListTemplate(myTemplate2); //应用新格式项目编号
            undo.EndCustomRecord();
            
        }

        private void GlyBak_Click(object sender, RibbonControlEventArgs e)  //点击副本下拉项目时执行的操作
        {
            switch(glyBak.SelectedItemIndex)
            {
                case 0:
                    btnSaveAsPDF_Click(sender, e);

                    break;
                case 1:
                    btnSaveAsDoc_Click(sender, e);

                    break;
                case 2:
                    btnSaveAsPic_Click(sender, e);

                    break;
                case 3:
                    BtnSaveAsJPG_Click(sender, e);

                    break;
            }
        }

        private void GlyBak_ItemsLoading(object sender, RibbonControlEventArgs e)  //wps下载入图片
        {
            try
            {
                if (wps.Contains("WPS"))
                {
                    //glyBak.Items[0].Image = Properties.Resources.FileEmailAsPdfEmailAttachment;
                    //glyBak.Items[1].Image = Properties.Resources.FileSaveAsWord97_2003;
                    //glyBak.Items[2].Image = Properties.Resources.SaveAsPng;
                    //glyBak.Items[3].Image = Properties.Resources.SaveAsJpg;
                }
            }
            catch (Exception ex)
            {
                //Globals.ThisAddIn.Logging(ex.Source, ex.Message);
            }
        }

 
        private void BtnFontChangeDialog_Click(object sender, RibbonControlEventArgs e) //显示字体更改对话框
        {
            FrmFontChangeDialog f = new FrmFontChangeDialog();
            f.ShowDialog();
            
        }



        private void btnHeadSingle_Click(object sender, RibbonControlEventArgs e)// 插入单行发文机关
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:单行发文机关");
            InsertHead("版头:单行发文机关");
            undo.EndCustomRecord();

        }
        private void btnHeadMulti_Click(object sender, RibbonControlEventArgs e) //插入联合发文机关
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:联合发文机关");
            InsertHead("版头:联合发文机关");
            undo.EndCustomRecord();
        }

        private void BtnLineH1_Click(object sender, RibbonControlEventArgs e) //插入函线上
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:函线(上沿)");
            InsertHead("版头:函线(上沿)");
            undo.EndCustomRecord();
        }

        private void BtnLineH2_Click(object sender, RibbonControlEventArgs e) //插入函线下
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入版头:函线(下沿)");
            InsertHead("版头:函线(下沿)");
            undo.EndCustomRecord();
        }

        private void BtnBindLine_Click(object sender, RibbonControlEventArgs e)//插入装订线
        {
            MSWord.UndoRecord undo = WordApp.UndoRecord;
            undo.StartCustomRecord("插入装订线");
            InsertHead("装订线");
            undo.EndCustomRecord();
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e) //测试按钮, debug时显示,"列出所有序号"
        {
            //ShowAllNumber();  //显示所有序号
            //WordApp.Run("TrustCenter"); //打开宏安全设置
            //delPageHeader();//删除页眉
            clearVBA();  //清理当前文档残留的vba代码
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)  //debug显示,列出所有内置面板
        {
            WordApp.Selection.TypeText("===========所有内置对话框 ===========\r\n");
            ShwoAllDialog();
            WordApp.Selection.TypeText("===========所有任务面板 ===========\r\n");
            ShowAllPanes();
        }




        /////////////////////////////////////////////////////////////////////
        ///
        ///       实验室代码
        /// 
        ////////////////////////////////////////////////////////////////////
        ///
        private void BtnSelectSharp_Click(object sender, RibbonControlEventArgs e)//选择全部图形
        {
            SelectShape();
            Globals.Ribbons.Ribbon1.Tabs[0].RibbonUI.ActivateTabMso("TabPictureToolsFormat");
        }

        private void BtnSelectTables_Click(object sender, RibbonControlEventArgs e)//选择全部表格
        {
            try
            {
                if (WordApp.ActiveDocument.ProtectionType >= 0) 
                {
                    WordApp.StatusBar="文档受到保护,无法选择!";
                    return; 
                }
                if (!WordApp.ActiveDocument.Equals(null))
                {
                    WordApp.Selection.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd);//光标移动
                    //曲线救国,通过给表格设置编辑者的方式,然后选择编辑者所拥有的内容.
                    foreach (MSWord.Table table in WordApp.ActiveDocument.Tables)
                    {
                        {
                            table.Range.Editors.Add(-1);   //-1代表 wdEditorEveryone
                        }
                    }
                    WordApp.ActiveDocument.SelectAllEditableRanges(-1);
                    WordApp.ActiveDocument.DeleteAllEditableRanges(-1);
                    WordApp.ActiveWindow.View.ShadeEditableRanges = 0;  //不显示可编辑权限的底纹色
                    Globals.Ribbons.Ribbon1.Tabs[0].RibbonUI.ActivateTabMso("TabTableToolsDesign");
                    //WordApp.ActiveDocument.UndoClear(); //此方法会清除所有的撤销记录.
                }
            }
            catch(Exception ex)
            {
                string err = ex.Message +"\r\n" + ex.Source;
                MessageBox.Show(err, "选择所有表格时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            
        }
        private void btnSelectCurrentPage_Click(object sender, RibbonControlEventArgs e) //全选当前页面
        {
            selectCurrentPage();
        }

        private void BtnHeadDelete_Click(object sender, RibbonControlEventArgs e)  //删除页眉,目前该按钮已隐藏.
        {
            delPageHeader();
        }

        private void btnSplit_Click(object sender, RibbonControlEventArgs e)  //拆分文档
        {
            splitDocBySection();
        }

        private void btnSharp2Top_Click(object sender, RibbonControlEventArgs e)  //所有衬于文字下方的图形转换为浮于文字上方
        {
            ConverUnderShape();
        }

        private void btnSharp2OnText_Click(object sender, RibbonControlEventArgs e) // 所有嵌入型图形转换为浮于文字上方
        {
            ConverSharp();
        }

        private void btnSharp2Wrapping_Click(object sender, RibbonControlEventArgs e)  //嵌入图形转为四周型
        {
            ConverSharp(2);
        }

        private void btnSharp2TopBonttom_Click(object sender, RibbonControlEventArgs e) //嵌入图形转为上下型
        {
            ConverSharp(3);
        }

        private void groupTest_DialogLauncherClick(object sender, RibbonControlEventArgs e)  //打开文档所在位置
        {
            try
            {
                if (WordApp.ActiveDocument.Saved)
                {
                    //获取word文档路径
                    string docPath = WordApp.ActiveDocument.FullName;
                    Uri uriFile = new Uri(docPath);
                    //判断是否本地文件
                    if (uriFile.IsFile)
                    {
                        Process.Start(@"explorer.exe", "/select,\"" + docPath + "\"");
                    }
                    else
                    {
                        string shortName = docPath;
                        for (int i = 0; i < 4; i++)
                        {
                            int index = shortName.IndexOf('/');
                            if (index >= 0)
                            {
                                shortName = shortName.Substring(index + 1);
                                docPath = System.Environment.GetEnvironmentVariable("OneDrive") + "/" + shortName;
                                docPath = docPath.Replace("/", "\\");
                            }
                        }
                        Process.Start(@"explorer.exe", "/select,\"" + docPath + "\"");
                        //Process.Start(docPath);
                    }
                }
            }
            catch(UriFormatException ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                WordApp.StatusBar = err;
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                //WordApp.StatusBar = err;
                MessageBox.Show(err, "打开文档所在文件夹时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
        }

        private void btnSelectObject_Click(object sender, RibbonControlEventArgs e)  //选择按钮
        {
            try
            {
                WordApp.Run("SelectDrawingObjects");
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "选择对象", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnVBA_Click(object sender, RibbonControlEventArgs e) //运行 VBA 代码按钮
        {
            OpenVba();
        }
        private void btnClearVBA_Click(object sender, RibbonControlEventArgs e) //清理代码按钮事件
        {
            clearVBA();
        }

        ///////////////////////////////////////////////////////////////
        ///    以上为实验室事件以下为实验室方法
        ///  //////////////////////////////////////////////////////////

        private void selectCurrentPage()//选择当前页全文
        {
            try
            {
                WordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
                WordApp.ActiveDocument.Bookmarks["\\page"].Range.Select(); //wps不支持
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "选择整页时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        void SaveSelect()  //保存所选内容为新文档
        {
            //WordApp.Dialogs[WdWordDialog.wdDialogFileSaveAs].Show();
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.AddExtension = true;
            sfd.Title = "将所选内容保存为可编辑的新文档";
            string fname = "新文档";
            try
            {
                fname = WordApp.Selection.Range.Paragraphs[1].Range.Text;
                fname = System.Text.RegularExpressions.Regex.Replace(fname, @"[^a-zA-Z0-9\u4e00-\u9fa5\s]", "");

                fname =fname.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").
                    Replace("\\","").Replace("*","").Replace("?", "").Replace(">", "").
                    Replace("/", "").Replace("<", "").Replace("|", "").Replace("\"", "").Replace(":", "");
                if (fname.Length > 20) fname = fname.Substring(0, 20);
            }
            finally
            {
                sfd.FileName = fname;
            }

            //此功能不包括pdf格式, 因为部分版本word(比如word365)会出错,且不提示.
            sfd.DefaultExt = ".docx|*.htm|*.mht|*.xml|*.txt";
            sfd.Filter = "Word文档(*.docx)|*.docx|标准 HTML 文档(*.htm)|*.htm|" +
                "单一 HTML 文档(*.mht)|*.mht|Word XML 文档(*.xml)|*.xml|文本文档(*.txt)|*.txt";
            DialogResult dr = sfd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                WordApp.StatusBar = "正在将所选内容存为可编辑的新文档... ";
                int wf = 0;
                switch (sfd.FilterIndex)
                {
                    case 1:
                        wf = 16;
                        break;
                    case 2:
                        wf = 8;
                        break;
                    case 3:
                        wf = 9;
                        break;
                    case 4:
                        wf = 12;
                        break;
                    case 5:
                        wf = 2;
                        break;
                }
                try
                {
                    WordApp.Selection.Range.ExportFragment(sfd.FileName, (WdSaveFormat)wf);
                    WordApp.StatusBar = "所选内容已保存!";
                    WordApp.Documents.Open(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message+ "\r\n"+ex.Source);
                    WordApp.StatusBar = "";
                }
            }
        }

        //保存长截图功能, 以文档存为html方式/选中内容保存 进行网页截图,但是效果不好.
        //1. html方式会改变文字排版,  尤其是选中内容保存会取消中文缩进.
        //2. 长截图质量低劣
        void SaveAsLongPic()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.AddExtension = true;
            sfd.Title = "将所选内容保存为图片";
            string fname = DateTime.Now.ToString("yyyyMMdd-hhmmss");
            sfd.FileName = fname +".png";
            DialogResult dr = sfd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                WordApp.StatusBar = "正在将当前页面保存成图片... ";

                string tmp = System.IO.Path.GetTempPath() + fname + ".html";
                WordApp.Selection.Range.ExportFragment(tmp, WdSaveFormat.wdFormatHTML);     //这个方法会改变格式,比如取消中文缩进
                //WordApp.ActiveDocument.SaveAs2(tmp, WdSaveFormat.wdFormatHTML);    //这个方法完成后会用word自动打开
                
                var myWebBrowser = new WebBrowser();
                myWebBrowser.Url = new Uri(tmp);
                
                myWebBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(delegate (object sender, WebBrowserDocumentCompletedEventArgs bdce)
                {
                    if (myWebBrowser.ReadyState == WebBrowserReadyState.Complete)
                    {
                        //把word宽度的磅值转换为浏览器的像素宽度
                        float w = WordApp.Selection.Range.PageSetup.PageWidth;
                        float m = WordApp.Selection.Range.PageSetup.LeftMargin + WordApp.Selection.Range.PageSetup.RightMargin;
                        int wwidth = (int)(792-m);
                        myWebBrowser.ScrollBarsEnabled = false;
                        myWebBrowser.Width = wwidth;
                        //MessageBox.Show(myWebBrowser.Width.ToString());
                        System.Drawing.Rectangle r = myWebBrowser.Document.Body.ScrollRectangle;
                        myWebBrowser.Height = r.Height;
                        var bitMapPic = new Bitmap(r.Width, r.Height);
                        myWebBrowser.DrawToBitmap(bitMapPic, r);
                        bitMapPic.Save(sfd.FileName,System.Drawing.Imaging.ImageFormat.Png);
                        WordApp.StatusBar = "页面保存图片成功!";
                        Process.Start(sfd.FileName); //打开图片.
                    }
                });
            }

        }

        private void ConverUnderShape()//转换所有文字下层的图形到上层
        {
            int i = 0;
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("转换衬于文字下图片");
                foreach (MSWord.Shape shape in WordApp.ActiveDocument.Shapes)
                {
                    if (shape.WrapFormat.Type == MSWord.WdWrapType.wdWrapBehind)
                    {
                        shape.WrapFormat.Type = MSWord.WdWrapType.wdWrapFront;
                        i++;
                    }
                }
            }
            catch(Exception ex)
            {
                WordApp.StatusBar = ex.Message;
            }
            finally
            {
                MessageBox.Show("共转换 " + i + "个衬于文字下方的图形为浮于文字上方!","小恐龙公文排版助手");
            }
        }

        private void SelectShape() //选择所有非嵌入图形
        {
            try
            {
                if (!WordApp.ActiveDocument.Equals(null))
                {
                    WordApp.ActiveDocument.Shapes.SelectAll();
                }
            }
            catch(Exception ex)
            {
                WordApp.StatusBar = ex.Message;
                return;
            }
        }
        /// <summary>
        /// 根据参数转换嵌入式图形
        /// </summary>
        /// <param name="type">要转换的图形环绕方式,默认1=浮于文字上方,2=四周型,3=上下型</param>
        void ConverSharp(int type=1)
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("转换图片环绕方式");
                int i = 0;
                foreach (MSWord.InlineShape shape in WordApp.ActiveDocument.InlineShapes)
                {
                    switch (type)
                    {
                        case 1:
                            shape.ConvertToShape();
                            i++;
                            break;
                        case 2:
                            shape.ConvertToShape().WrapFormat.Type = WdWrapType.wdWrapSquare;
                            i++;
                            break;
                        case 3:
                            shape.ConvertToShape().WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                            i++;
                            break;
                    }
                }
                undo.EndCustomRecord();
                WordApp.StatusBar = "已转换 " + i + "个图形!";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        void delPageHeader()//删除页眉
        {
            try
            {
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("删除页眉");
                Debug.WriteLine("首页不同:" + WordApp.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter);
                Debug.WriteLine("奇偶页不同:" + WordApp.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter);
                WordApp.ScreenUpdating = false;
                if (WordApp.ActiveWindow.ActivePane.View.SeekView != MSWord.WdSeekView.wdSeekCurrentPageHeader)  //打开当前页面页眉
                {
                    WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
                }
                WordApp.Selection.WholeStory();
                //WordApp.Selection.ClearCharacterAllFormatting();
                //WordApp.Selection.ClearCharacterDirectFormatting();
                //WordApp.Selection.ClearCharacterStyle();
                WordApp.Selection.Delete();
                WordApp.Selection.ClearFormatting();

                MSWord.Style style = WordApp.ActiveDocument.Styles.Add("公文:页眉", MSWord.WdStyleType.wdStyleTypeParagraph);
                style.set_BaseStyle(""); //样式基于无格式样式.

                WordApp.Selection.set_Style(style);

                WordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                WordApp.ScreenUpdating = true;
                undo.EndCustomRecord();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "删除页面横线失败", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                WordApp.ScreenUpdating = true;
            }
        }


        void splitDocBySection()  //根据分节符拆分文档
        {
            MSWord.Document olddoc = WordApp.ActiveDocument;
            if(olddoc.Sections.Count<=1)
            {
                MessageBox.Show("拆分文档前请插入分节符!","小恐龙公文排版助手:拆分文档",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            try
            {
                WordApp.CommandBars["Office Clipboard"].Visible = true; //启用office剪贴板
                WordApp.CommandBars["Office Clipboard"].Visible = false; //关闭剪贴贴面板并不会停用收集粘贴
            }
            catch(Exception ex)
            {
                MessageBox.Show("Office 功能不完整(比如精简版、绿色版）,拆分文档将无法继续!", "拆分文档遇到错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int j = olddoc.Sections.Count + 1;
            for (int i = 1; i <j ; i++)
            {
                olddoc.Sections[i].Range.Copy();   //复制每一节内容
                MSWord.Document newdoc = WordApp.Documents.Add(); //新建文档
                newdoc.Content.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting); //使用源格式粘贴到新建文档
                newdoc.PageSetup = olddoc.Sections[i].PageSetup;  //复制原文档页面设置
                WordReplace(false, "^b", "");  //清理新文档的分节符
            }
        }
        void ShowAllNumber()  //显示所有编号
        {
            MSWord.ListTemplate myTemplate = WordApp.ActiveDocument.ListTemplates.Add();  //添加编号
            myTemplate.ListLevels[1].TrailingCharacter = MSWord.WdTrailingCharacter.wdTrailingNone;//编号与文字之间无空格无Tab
            myTemplate.ListLevels[1].TextPosition = 0; //设置第二行缩进量
            myTemplate.ListLevels[1].NumberPosition = 0;

            myTemplate.ListLevels[1].StartAt = 1;//设置序列

           
            foreach (int myCode in Enum.GetValues(typeof(MSWord.WdListNumberStyle)))
            {
                string strName = Enum.GetName(typeof(MSWord.WdListNumberStyle), myCode);//获取名称
                string strVaule = myCode.ToString();
                WordApp.Selection.TypeText(strName + "   " + strVaule + "\r\n");
                try
                {
                    myTemplate.ListLevels[1].NumberStyle = (MSWord.WdListNumberStyle)myCode;
                    WordApp.Selection.Range.ListFormat.ApplyListTemplate(myTemplate); //应用新格式项目编号
                }
                catch
                { }

                WordApp.ActiveDocument.Content.ListFormat.ConvertNumbersToText();
                WordApp.Selection.TypeText("\r\n"); 
            }
        }
        void ShwoAllDialog() //列出所有内置对话框
        {
            foreach (int myCode in Enum.GetValues(typeof(MSWord.WdWordDialog)))
            {
                string strName = Enum.GetName(typeof(MSWord.WdWordDialog), myCode);//获取名称
                string strVaule = myCode.ToString();
                WordApp.Selection.TypeText(strName + "   " + strVaule + "\r\n");
            }
        }
        void ShowAllPanes() //列出所有内置面板
        {
            foreach (int myCode in Enum.GetValues(typeof(MSWord.WdTaskPanes)))
            {
                string strName = Enum.GetName(typeof(MSWord.WdTaskPanes), myCode);//获取名称
                string strVaule = myCode.ToString();
                WordApp.Selection.TypeText(strName + "   " + strVaule + "\r\n");
            }
        }

        //打开并运行一个VBA宏
        ///执行完毕后删除
        private void OpenVba()
        {
            try
            {
                string vbaFile = "";
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "VBA文件(*.vba,*.bas)|*.vba;*.bas|全部文件(*.*)|*.*";
                DialogResult dr = ofd.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    vbaFile = ofd.FileName;
                    Encoding encoding = GetTextFileEncodingType(vbaFile);
                    if (encoding == Encoding.Unicode || encoding == Encoding.UTF8)
                    {
                        DialogResult result = MessageBox.Show("您所选文件非ANSI编码,WordVBA不支持UTF8编码。\r\n您要用记事本打开并查看文件编码吗?\r\n单击确定将运行记事本程序,取消退出。", "读取VBA文件", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                        if (result == DialogResult.OK)
                        {
                            try { Process.Start("notepad", vbaFile); }
                            catch (Exception ex) { throw; }
                        }
                        return;
                    }

                    MSWord.Document doc = WordApp.ActiveDocument;
                    VBComponents vbc = doc.VBProject.VBComponents;
                    VBComponent vbm = vbc.Add(vbext_ComponentType.vbext_ct_StdModule);
                    vbm.Name = "小恐龙公文排版助手";
                    CodeModule module = vbm.CodeModule;
                    module.AddFromFile(vbaFile);

                    vbext_ProcKind ProcKind;
                    int startLine = module.CountOfDeclarationLines + 1;
                    string macroName = module.get_ProcOfLine(startLine, out ProcKind);
                    //MessageBox.Show(macroName);

                    System.Diagnostics.Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    MSWord.UndoRecord undo = WordApp.UndoRecord;
                    undo.StartCustomRecord("运行VBA:" + macroName);
                    WordApp.Run(macroName);
                    undo.EndCustomRecord();
                    vbc.Remove(vbm);
                    stopwatch.Stop();
                    WordApp.StatusBar = $"VBA文件执行共用时:{stopwatch.Elapsed.TotalSeconds:F2} 秒";

                    try
                    {
                        string fileNameNoExt = Path.GetFileNameWithoutExtension(vbaFile);
                        string item = fileNameNoExt + "|" + vbaFile;
                        if (Properties.Settings.Default.VBA历史 == null)
                        {
                            Properties.Settings.Default.VBA历史 = new System.Collections.Specialized.StringCollection();
                        }
                        if (!Properties.Settings.Default.VBA历史.Contains(item))
                        {
                            RibbonButton btnVBAlist;
                            btnVBAlist = this.Factory.CreateRibbonButton();
                            btnVBAlist.Label = fileNameNoExt;
                            btnVBAlist.ScreenTip = vbaFile;
                            btnVBAlist.ControlSize = MSOffice.RibbonControlSize.RibbonControlSizeRegular;
                            btnVBAlist.Click += BtnVBAlist_Click;
                            mnuScript.Visible = true;
                            mnuScript.Items.Add(btnVBAlist);
                            Properties.Settings.Default.VBA历史.Add(item);
                            Properties.Settings.Default.Save();
                            //Properties.Settings.Default.Upgrade();
                        }
                    }
                    catch (Exception ex)
                    {
                        string err = ex.ToString();
                        MessageBox.Show(err, "添加VBA历史记录时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
            catch (COMException ex)
            {
                string err = "";
                string helpURL = "";
                switch (ex.ErrorCode)
                {
                    case -2146822220:
                        err = "Word 未开启信任对VBA工程对象模型的访问。";
                        err+="\r\n\r\n点击[确定]获取更多信息！";
                        helpURL= "https://gw.xkonglong.com/#/vba";
                        break;
                    case -2147467259:
                        err = "WPS用户请先切换至【VB环境】,并在【宏安全性】选项中开启【信任对VBA工程对象模型的访问】";
                        err += "\r\n\r\n点击[确定]获取更多信息！";
                        helpURL= "https://gw.xkonglong.com/#/vba";
                        break;
                    case -2147352573:
                        err = ex.Message;
                        err += "\r\n\r\n点击[确定]获取更多信息！";
                        helpURL = "https://gw.xkonglong.com/#/vba";
                        break;
                    default:
                        err = "错误代码:" + ex.ErrorCode.ToString();
                        err += "\r\n错误信息:" + ex.Message;
                        Clipboard.SetText(err);
                        err += "\r\n\r\n[错误信息已复制到剪贴板]\r\n点击[确定]向作者反馈问题！";
                        helpURL = "mailto:xkonglong@qq.com?subject=小恐龙公文助手问题反馈";  //反馈问题url
                        break;
                }
                DialogResult dr = MessageBox.Show(err, "执行VBA代码出错", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dr == DialogResult.OK)
                {
                    Process.Start(helpURL);
                }
            }

            catch (Exception ex)
            {
                string err = ex.ToString();   //ex.Message;
                MessageBox.Show(err, "执行VBA代码出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void BtnVBA_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton btn = sender as RibbonButton;
            string file = btn.ScreenTip;
            RunVBA(file);
        }

        //执行vba文件
        private bool RunVBA(string vbaFile)
        {
            try
            {

                Encoding encoding = GetTextFileEncodingType(vbaFile);
                if (encoding == Encoding.Unicode || encoding == Encoding.UTF8)
                {
                    DialogResult result = MessageBox.Show("您所选文件非ANSI编码,WordVBA不支持UTF8编码。\r\n您要用记事本打开并查看文件编码吗?\r\n单击确定将运行记事本程序,取消退出。", "读取VBA文件", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        try { Process.Start("notepad", vbaFile); }
                        catch (Exception ex) { throw; }
                    }
                    return false;
                }

                MSWord.Document doc = WordApp.ActiveDocument;
                VBComponents vbc = doc.VBProject.VBComponents;
                VBComponent vbm = vbc.Add(vbext_ComponentType.vbext_ct_StdModule);
                vbm.Name = "小恐龙公文排版助手";
                CodeModule module = vbm.CodeModule;
                module.AddFromFile(vbaFile);

                vbext_ProcKind ProcKind;
                int startLine = module.CountOfDeclarationLines + 1;
                string macroName = module.get_ProcOfLine(startLine, out ProcKind);
                //MessageBox.Show(macroName);

                System.Diagnostics.Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                MSWord.UndoRecord undo = WordApp.UndoRecord;
                undo.StartCustomRecord("运行VBA:" + macroName);
                WordApp.Run(macroName);
                undo.EndCustomRecord();
                vbc.Remove(vbm);
                stopwatch.Stop();
                WordApp.StatusBar = $"VBA文件执行共用时:{stopwatch.Elapsed.TotalSeconds:F2} 秒";
                return true;
            }
            catch(Exception ex)
            {
                WordApp.StatusBar = ex.Message;
                return false;
            }
        }
        private void BtnVBAlist_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton btn = sender as RibbonButton;
            string file = btn.ScreenTip;
            RunVBA(file);
        }

        /// <summary>
        /// 获取文本文件的字符编码类型
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        static Encoding GetTextFileEncodingType(string fileName)
        {
            Encoding encoding = Encoding.Default;
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            BinaryReader binaryReader = new BinaryReader(fileStream, encoding);
            byte[] buffer = binaryReader.ReadBytes((int)fileStream.Length);
            binaryReader.Close();
            fileStream.Close();
            if (buffer.Length >= 3 && buffer[0] == 239 && buffer[1] == 187 && buffer[2] == 191)
            {
                encoding = Encoding.UTF8;
            }
            else if (buffer.Length >= 3 && buffer[0] == 254 && buffer[1] == 255 && buffer[2] == 0)
            {
                encoding = Encoding.BigEndianUnicode;
            }
            else if (buffer.Length >= 3 && buffer[0] == 255 && buffer[1] == 254 && buffer[2] == 65)
            {
                encoding = Encoding.Unicode;
            }
            else if (IsUTF8Bytes(buffer))
            {
                encoding = Encoding.UTF8;
            }
            return encoding;
        }

        /// <summary>
        /// 判断是否是不带 BOM 的 UTF8 格式
        /// BOM（Byte Order Mark），字节顺序标记，出现在文本文件头部，Unicode编码标准中用于标识文件是采用哪种格式的编码。
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static bool IsUTF8Bytes(byte[] data)
        {
            int charByteCounter = 1; //计算当前正分析的字符应还有的字节数 
            byte curByte; //当前分析的字节. 
            for (int i = 0; i < data.Length; i++)
            {
                curByte = data[i];
                if (charByteCounter == 1)
                {
                    if (curByte >= 0x80)
                    {
                        //判断当前 
                        while (((curByte <<= 1) & 0x80) != 0)
                        {
                            charByteCounter++;
                        }
                        //标记位首位若为非0 则至少以2个1开始 如:110XXXXX...........1111110X 
                        if (charByteCounter == 1 || charByteCounter > 6)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    //若是UTF-8 此时第一位必须为1 
                    if ((curByte & 0xC0) != 0x80)
                    {
                        return false;
                    }
                    charByteCounter--;
                }
            }
            if (charByteCounter > 1)
            {
                throw new Exception("非预期的byte格式");
            }
            return true;
        }

        private void clearVBA()
        {
            try
            {
                MSWord.Document doc = WordApp.ActiveDocument;
                VBComponents vbcs = doc.VBProject.VBComponents;
                string vbcName = "";

                foreach (VBComponent vbc in vbcs)
                {
                    vbcName += vbc.Name + "\r\n";
                    switch (vbc.Type)
                    {
                        case vbext_ComponentType.vbext_ct_StdModule:   //标准模块
                        case vbext_ComponentType.vbext_ct_ClassModule:  //类模块
                        case vbext_ComponentType.vbext_ct_MSForm:   //窗体

                            vbcs.Remove(vbc);
                            break;
                        case vbext_ComponentType.vbext_ct_ActiveXDesigner:   //控件
                        case vbext_ComponentType.vbext_ct_Document:   //文档
                            vbc.CodeModule.DeleteLines(1, vbc.CodeModule.CountOfLines);
                            break;
                    }
                }
                MessageBox.Show("当前文档中:\r\n" + vbcName + "已清理","清理VBA代码",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string err = ex.Message;
                MessageBox.Show(err, "清理VBA代码时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void LoadVBA(string vbaFolder="")
        {
            //读取所有vba脚本
            if (vbaFolder == "")
            {
                try
                {   //获取属性文件中的vba脚本路径
                    vbaFolder = Properties.Settings.Default.VBA路径;
                    if (vbaFolder == "")   //如果为空，默认为程序所在路径
                    {
                        string startPath = GetDllPath();
                        vbaFolder = Path.Combine(startPath, "VBA脚本");
                    }
                }
                catch (Exception ex)
                {
                    string err = ex.Message;
                    MessageBox.Show(err, "获取VBA脚本路径时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            if (Directory.Exists(vbaFolder))
            {
                foreach (string file in Directory.GetFiles(vbaFolder))
                {
                    string fileName = Path.GetFileName(file);
                    string fileNameNoExt = Path.GetFileNameWithoutExtension(file);
                    string fileExt = Path.GetExtension(file);
                    if (fileExt.ToLower() == ".vba")
                    {
                        RibbonButton btn;
                        btn = this.Factory.CreateRibbonButton();
                        btn.Label = fileNameNoExt;
                        btn.Tag = file;
                        btn.ScreenTip = file;
                        btn.Click += BtnVBAlist_Click;
                        mnuVBA.Items.Add(btn);
                    }
                }
            }
        }
    }
}
