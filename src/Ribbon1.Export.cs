using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        /// <summary>确保文档已保存,未保存时提示保存.返回 false 表示用户取消.</summary>
        private bool EnsureDocumentSaved(string title)
        {
            if (WordApp.ActiveDocument.Path.Length != 0) return true;
            var dr = MessageBox.Show("生成副本前请先保存本文档,要保存文档吗?", title,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr != DialogResult.Yes) return false;
            try
            {
                WordApp.ActiveDocument.Save();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n由于文档未保存,生成副本操作被取消.", "生成副本失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        private void SavePdf(string fileName)
        {
            WordApp.StatusBar = "正在生成 PDF 格式副本...";
            WordApp.ActiveDocument.ExportAsFixedFormat(fileName,
                MSWord.WdExportFormat.wdExportFormatPDF, true, 0, 0, 0, 0, 0, false, false,
                MSWord.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks);
            WordApp.StatusBar = "已生成 PDF 格式副本:" + fileName;
        }

        private void SaveDoc(string fileName, string oldFile)
        {
            WordApp.StatusBar = "正在生成 Word97-2003 格式副本...";
            WordApp.ActiveDocument.SaveAs2(fileName, MSWord.WdSaveFormat.wdFormatDocument97);
            WordApp.StatusBar = "已生成 Word97-2003 格式副本:" + fileName;
            WordApp.Documents.Open(oldFile);
            int w = WordApp.Windows.Count;
            if (w > 1) WordApp.Windows[w - 1].Activate();
        }

        private void btnSaveAsPDF_Click(object sender, RibbonControlEventArgs e)
        {
            if (!EnsureDocumentSaved("生成PDF副本提示")) return;
            try
            {
                string fp = WordApp.ActiveDocument.Path;
                string fn = Path.GetFileNameWithoutExtension(WordApp.ActiveDocument.Name);
                string outFile = Path.Combine(fp, fn + ".pdf");

                if (File.Exists(outFile))
                {
                    var drr = MessageBox.Show(
                        "文件 \r\n" + outFile + "\r\n已存在! " +
                        "\r\n点击\"是\"将覆盖同名文件.\r\n点击\"否\"将打开保存对话框更改路径或文件名.\r\n点击\"取消\"将取消生成副本.",
                        "生成副本时发现同名文件", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (drr == DialogResult.No)
                    {
                        using (var sfd = new SaveFileDialog())
                        {
                            sfd.AddExtension = true;
                            sfd.DefaultExt = "pdf";
                            sfd.Filter = "PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
                            sfd.InitialDirectory = fp;
                            sfd.FileName = outFile;
                            if (sfd.ShowDialog() == DialogResult.OK) SavePdf(sfd.FileName);
                        }
                    }
                    else if (drr == DialogResult.Yes) SavePdf(outFile);
                }
                else SavePdf(outFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "生成pdf副本时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSaveAsDoc_Click(object sender, RibbonControlEventArgs e)
        {
            if (!EnsureDocumentSaved("生成DOC副本提示")) return;
            try
            {
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                string oldFile = Path.Combine(fp, fn);
                string baseName = Path.GetFileNameWithoutExtension(fn);
                string outFile = Path.Combine(fp, baseName + ".doc");

                if (File.Exists(outFile))
                {
                    var drr = MessageBox.Show(
                        "文件 \r\n" + outFile + "\r\n已存在! " +
                        "\r\n点击\"是\"将覆盖同名文件.\r\n点击\"否\"将打开保存对话框更改路径或文件名.\r\n点击\"取消\"将取消生成副本.",
                        "生成副本时发现同名文件", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (drr == DialogResult.No)
                    {
                        using (var sfd = new SaveFileDialog())
                        {
                            sfd.AddExtension = true;
                            sfd.DefaultExt = "doc";
                            sfd.Filter = "Word97-2003格式文档(*.doc)|*.doc|所有文件 (*.*)|*.*";
                            sfd.InitialDirectory = fp;
                            sfd.FileName = outFile;
                            if (sfd.ShowDialog() == DialogResult.OK) SaveDoc(sfd.FileName, oldFile);
                        }
                    }
                    else if (drr == DialogResult.Yes) SaveDoc(outFile, oldFile);
                }
                else SaveDoc(outFile, oldFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source, "生成doc副本时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSaveAsPic_Click(object sender, RibbonControlEventArgs e)
        {
            if (!EnsureDocumentSaved("生成PNG副本提示")) return;
            try
            {
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                string oldFile = Path.Combine(fp, fn);
                MSWord.Pages pages = WordApp.ActiveDocument.ActiveWindow.ActivePane.Pages;

                for (int i = 1; i <= pages.Count; i++)
                {
                    WordApp.StatusBar = "正在将页面转换成图片 " + i + "...";
                    MSWord.Page page = pages[i];
                    using (var st = new MemoryStream((byte[])(page.EnhMetaFileBits)))
                    using (var img = Image.FromStream(st))
                    using (var bmp = img.Width > img.Height ? new Bitmap(img, 1754, 1240) : new Bitmap(img, 1240, 1754))
                    {
                        bmp.Save(string.Format(oldFile + "_{0}.png", i), ImageFormat.Png);
                    }
                }
                WordApp.StatusBar = "文字转图片已完成,请在打开的文件夹中查看";
                System.Diagnostics.Process.Start(fp);
            }
            catch (Exception ex)
            {
                WordApp.StatusBar = ex.Message;
            }
        }

        private void BtnSaveAsJPG_Click(object sender, RibbonControlEventArgs e)
        {
            float dpi = 150;
            try { dpi = Properties.Settings.Default.DPI; } catch { }

            if (!EnsureDocumentSaved("生成JPG副本提示")) return;
            try
            {
                string fp = WordApp.ActiveDocument.Path;
                string fn = WordApp.ActiveDocument.Name;
                string oldFile = Path.Combine(fp, fn);
                MSWord.Pages pages = WordApp.ActiveDocument.ActiveWindow.ActivePane.Pages;

                for (int i = 1; i <= pages.Count; i++)
                {
                    WordApp.StatusBar = "正在将页面转换成图片 " + i + "...";
                    MSWord.Page page = pages[i];
                    using (var st = new MemoryStream((byte[])(page.EnhMetaFileBits)))
                    using (var img = Image.FromStream(st))
                    {
                        int w = img.Width > img.Height ? 1754 : 1240;
                        int h = img.Width > img.Height ? 1240 : 1754;
                        using (var b = new Bitmap(w, h))
                        {
                            b.SetResolution(dpi, dpi);
                            using (var g = Graphics.FromImage(b))
                            {
                                g.Clear(Color.White);
                                g.DrawImage(img, 0, 0, w, h);
                            }
                            b.Save(string.Format(oldFile + "_{0}.jpg", i), ImageFormat.Jpeg);
                        }
                    }
                }
                WordApp.StatusBar = "文字转图片已完成,请在打开的文件夹中查看";
                System.Diagnostics.Process.Start(fp);
            }
            catch (Exception ex)
            {
                WordApp.StatusBar = ex.Message;
            }
        }

        private void btnSaveAsPicAt_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var sfd = new SaveFileDialog())
                {
                    sfd.AddExtension = true;
                    sfd.Title = "将当前页面保存为图片,您可以根据需求选择图片格式.";
                    sfd.FileName = DateTime.Now.ToString("yyyyMMdd-hhmmss");
                    sfd.DefaultExt = "png";
                    sfd.Filter = "A4尺寸网络交换图片(*.png)|*.png|原始尺寸网络交换图片(*.png)|*.png|原始尺寸有损压缩图片(*.jpg)|*.jpg|原始尺寸高精度文档图片(*.tiff)|*.tiff|所有文件(*.*)|*.*";
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    WordApp.StatusBar = "正在将当前页面保存成图片... ";
                    int i = WordApp.Selection.Information[MSWord.WdInformation.wdActiveEndPageNumber];
                    MSWord.Page page = WordApp.ActiveDocument.ActiveWindow.ActivePane.Pages[i];

                    using (var sr = new MemoryStream((byte[])(page.EnhMetaFileBits)))
                    using (var img = Image.FromStream(sr))
                    {
                        switch (sfd.FilterIndex)
                        {
                            case 1:
                                using (var png = img.Width > img.Height ? new Bitmap(img, 1754, 1240) : new Bitmap(img, 1240, 1754))
                                    png.Save(sfd.FileName, ImageFormat.Png);
                                break;
                            case 2:
                                img.Save(sfd.FileName, ImageFormat.Png);
                                break;
                            case 3:
                                using (var b = new Bitmap(img.Width, img.Height))
                                {
                                    b.SetResolution(img.HorizontalResolution, img.VerticalResolution);
                                    using (var g = Graphics.FromImage(b))
                                    {
                                        g.Clear(Color.White);
                                        g.DrawImageUnscaled(img, 0, 0);
                                    }
                                    b.Save(sfd.FileName, ImageFormat.Jpeg);
                                }
                                break;
                            case 4:
                                img.Save(sfd.FileName, ImageFormat.Tiff);
                                break;
                        }
                    }
                    WordApp.StatusBar = "页面保存图片成功!";
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                WordApp.StatusBar = "页面保存图片失败.";
            }
        }

        private void GlyBak_Click(object sender, RibbonControlEventArgs e)
        {
            switch (glyBak.SelectedItemIndex)
            {
                case 0: btnSaveAsPDF_Click(sender, e); break;
                case 1: btnSaveAsDoc_Click(sender, e); break;
                case 2: btnSaveAsPic_Click(sender, e); break;
                case 3: BtnSaveAsJPG_Click(sender, e); break;
            }
        }

        private void GlyBak_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            // WPS 下如需动态图标可在此处理
        }
    }
}