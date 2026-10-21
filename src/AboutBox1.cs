using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Diagnostics;

namespace 公文排版
{
    partial class AboutBox1 : Form
    {
        private readonly ToolTip _tip = new ToolTip();
        private string _downloadUrl;
        private bool _updateChecked;

        public AboutBox1()
        {
            InitializeComponent();
            string name = Assembly.GetExecutingAssembly().GetName().Name;
            string ver = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = "关于" + name + ver;
            this.Height = this.MinimumSize.Height;
        }

        private void LoadFont(ComboBox combo)
        {
            if (combo == null) return;
            try
            {
                combo.BeginUpdate();
                combo.Items.Clear();
                using (var ifc = new System.Drawing.Text.InstalledFontCollection())
                {
                    foreach (FontFamily ff in ifc.Families)
                        combo.Items.Add(ff.Name);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            finally
            {
                combo.EndUpdate();
            }
        }

        private void AboutBox1_Load(object sender, EventArgs e)
        {
            try
            {
                string oVer = "Windows " + Environment.OSVersion.Version;
                oVer += Environment.Is64BitOperatingSystem ? " 64位" : " 32位";

                string wVer = Application.ProductName + "(" + Application.ProductVersion + ")";
                wVer += Environment.Is64BitProcess ? " 64位" : " 32位";

                string wPath = Application.StartupPath;
                string mPath = Assembly.GetExecutingAssembly().Location;

                this.txtInfo.Text = this.Text + "\r\n" + oVer + "\r\n" + wVer + "\r\n" + wPath + "\r\n" + mPath;

                tableLayoutPanel1.Dock = DockStyle.Fill;

                LoadFont(cobLetter);

                // TableLayoutPanel 双缓冲，减少刷新闪烁
                typeof(TableLayoutPanel)
                    .GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.SetValue(tableLayoutPanel1, true, null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("AboutBox1_Load: " + ex.Message);
                // 不再 throw，窗体仍可打开
            }
        }

        internal sealed class UpdateResult
        {
            public bool Success;
            public Version NewVersion;
            public string DownloadUrl;
            public string Message;
        }

        /// <summary>
        /// 通过指定的URL获取版本更新信息. 返回 null 表示网络失败.
        /// </summary>
        public static UpdateResult CheckUpdate(string url)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(new Uri(url));
                request.KeepAlive = false;
                request.AllowAutoRedirect = true;
                request.UserAgent = "xkonglong WordAddin Check Update Client " +
                                     Assembly.GetExecutingAssembly().GetName().Version;
                request.Timeout = 10000;
                request.ReadWriteTimeout = 10000;
                request.Method = "GET";

                using (var res = (HttpWebResponse)request.GetResponse())
                {
                    if (res.StatusCode != HttpStatusCode.OK && res.StatusCode != HttpStatusCode.PartialContent)
                    {
                        return new UpdateResult { Success = false, Message = "网络错误代码:" + (int)res.StatusCode };
                    }

                    using (var stream = res.GetResponseStream())
                    using (var reader = new StreamReader(stream))
                    {
                        string text = reader.ReadToEnd().Trim();
                        if (string.IsNullOrEmpty(text))
                            return new UpdateResult { Success = false, Message = "版本信息为空" };

                        var result = new UpdateResult { Success = true };
                        if (text.Contains("|"))
                        {
                            string[] parts = text.Split(new[] { '|' }, 2);
                            try { result.NewVersion = new Version(parts[0].Trim()); }
                            catch { return new UpdateResult { Success = false, Message = "未知的版本信息!" }; }
                            result.DownloadUrl = parts.Length > 1 ? parts[1].Trim() : null;
                        }
                        else
                        {
                            try { result.NewVersion = new Version(text); }
                            catch { return new UpdateResult { Success = false, Message = "未知的版本信息!" }; }
                        }
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("CheckUpdate: " + ex.Message);
                return new UpdateResult { Success = false, Message = "获取信息出错,请检查网络!" };
            }
        }

        private void LnkUpdate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!_updateChecked)
            {
                _updateChecked = true;
                var result = CheckUpdate("https://xkonglong.com/down/update.txt");
                if (result == null || !result.Success)
                {
                    lnkUpdate.Text = result?.Message ?? "检查更新失败";
                    lnkUpdate.Enabled = false;
                    return;
                }

                Version current = Assembly.GetExecutingAssembly().GetName().Version;
                if (result.NewVersion.CompareTo(current) > 0)
                {
                    _downloadUrl = result.DownloadUrl ?? "https://xkonglong.com";
                    lnkUpdate.Text = "最新版本:" + result.NewVersion + " 下载可用";
                }
                else
                {
                    lnkUpdate.Text = "您已是最新版本!";
                    lnkUpdate.Enabled = false;
                }
                return;
            }

            if (!string.IsNullOrEmpty(_downloadUrl))
            {
                try
                {
                    Process.Start(_downloadUrl);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\r\n请手动打开浏览器输入网址访问:\r\n" + _downloadUrl);
                }
            }
        }

        private void TxtInfo_Click(object sender, EventArgs e)
        {
            if (!(sender is TextBox tbox)) return;
            if (string.IsNullOrEmpty(txtInfo.Text)) return;
            try
            {
                Clipboard.SetText(txtInfo.Text);
                _tip.Show("系统信息已复制", tbox, 1000);
            }
            catch (Exception ex)
            {
                _tip.Show("系统信息复制时出错!\r\n" + ex.Message, tbox, 1000);
            }
        }

        private void ShowAdvancedPanel(bool show)
        {
            if (show)
            {
                this.Height = 400;
                panel1.Visible = false;
                lblTip.Visible = true;
                lblTip.Text = "用户设置，双击恢复默认.";
                GetFont();
                btnCustom.Text = "<<高级功能";
                btnSave.Visible = true;
                tableLayoutPanel1.Dock = DockStyle.Fill;
                tableLayoutPanel1.Visible = true;
            }
            else
            {
                lblTip.Visible = false;
                btnSave.Visible = false;
                this.Height = MinimumSize.Height;
                panel1.Visible = true;
                tableLayoutPanel1.Visible = false;
                btnCustom.Text = ">>高级功能";
            }
        }

        private void BtnCustom_Click(object sender, EventArgs e)
        {
            ShowAdvancedPanel(panel1.Visible);
        }

        private void SetFont(object sender, EventArgs e)
        {
            if (!(sender is Button btn)) return;
            try
            {
                using (var fd = new FontDialog())
                {
                    fd.Font = btn.Font;
                    fd.AllowSimulations = true;
                    fd.ShowEffects = false;
                    fd.MinSize = 9;
                    fd.FontMustExist = true;

                    if (fd.ShowDialog(this) == DialogResult.OK)
                    {
                        btn.Text = fd.Font.Name;
                        btn.Font = fd.Font;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n当前用户权限无法访问该字体,请为所有用户安装该字体。",
                    "打开字体对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private static void ApplyFont(Button btn, Font font)
        {
            if (btn == null || font == null) return;
            btn.Font = font;
            btn.Text = font.Name;
        }

        private void GetFont()
        {
            try
            {
                var s = Properties.Settings.Default;
                ApplyFont(button1, s.公文标题);
                ApplyFont(button2, s.一级标题);
                ApplyFont(button3, s.二级标题);
                ApplyFont(button4, s.三级标题);
                ApplyFont(button5, s.四级标题);
                ApplyFont(button6, s.正文);
                ApplyFont(button7, s.版头);
                ApplyFont(button8, s.签发人姓名);
                ApplyFont(button9, s.版记);
                ApplyFont(button10, s.页码);

                string letterFont = s.英文字体;
                if (!string.IsNullOrEmpty(letterFont))
                    cobLetter.Text = letterFont;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "读取公文格式配置文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SaveConfig()
        {
            try
            {
                var s = Properties.Settings.Default;
                s.公文标题 = button1.Font;
                s.一级标题 = button2.Font;
                s.二级标题 = button3.Font;
                s.三级标题 = button4.Font;
                s.四级标题 = button5.Font;
                s.正文 = button6.Font;
                s.版头 = button7.Font;
                s.签发人姓名 = button8.Font;
                s.版记 = button9.Font;
                s.页码 = button10.Font;

                string letterFont = cobLetter.Text?.Trim();
                if (string.IsNullOrEmpty(letterFont))
                {
                    MessageBox.Show("未指定默认西文字体!", this.Text);
                    return;
                }
                s.英文字体 = letterFont;

                s.Save();

                MessageBox.Show(
                    "用户自定义配置文件已保存！\r\n 用户配置文件在Word更新后会被重置.您可通过双击赞赏码修改默认配置文件.",
                    this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

                ShowAdvancedPanel(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "保存公文格式配置文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void AboutBox1_HelpButtonClicked(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                Process.Start(Properties.Settings.Default.帮助_高级功能);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "打开在线帮助时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveConfig();
        }

        private void logoPictureBox_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string asmPath = Assembly.GetExecutingAssembly().Location;
                string configPath = asmPath + ".config";
                string baseDir = Path.GetDirectoryName(asmPath);
                string editorPath = Path.Combine(baseDir ?? string.Empty, "Resources", "ConfigEditor.exe");

                if (!File.Exists(editorPath))
                {
                    MessageBox.Show("未找到配置编辑器:" + editorPath, this.Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Process.Start(editorPath, "\"" + configPath + "\"");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n\r\n" + ex.Source,
                    "打开默认配置文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void lblTip_DoubleClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定恢复默认字体设置吗？", "确认",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                Properties.Settings.Default.Reset();
                Properties.Settings.Default.Reload();
                GetFont();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "重置设置时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}