using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Configuration; //test


namespace 公文排版
{
    partial class AboutBox1 : Form
    {
        public AboutBox1()
        {
            InitializeComponent();
            string name = Assembly.GetExecutingAssembly().GetName().Name;
            string ver = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = "关于" + name + ver;
            this.Height = this.MinimumSize.Height;
        }
        // 组合框载入字体的实现方法
        private void LoadFont(object sender)
        {
            try
            {
                //载入字体
                ((ComboBox)sender).Items.Clear();
                System.Drawing.Text.InstalledFontCollection ifc = new System.Drawing.Text.InstalledFontCollection();
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
                Debug.WriteLine(ex.Message);
            }
        }

        private void AboutBox1_Load(object sender, EventArgs e)
        {
            try
            {
                //操作系统信息
                string oVer = "Windows " + System.Environment.OSVersion.Version.ToString();
                if (Environment.Is64BitOperatingSystem)
                {
                    oVer += " 64位";
                }
                else
                {
                    oVer += " 32位";
                }
                //Office信息
                string wVer = Application.ProductName + "(" + Application.ProductVersion + ")";
                if (Environment.Is64BitProcess)
                {
                    wVer += " 64位";
                }
                else
                {
                    wVer += " 32位";
                }
                string wPath = Application.StartupPath;   //office运行路径
                string mPath = Assembly.GetExecutingAssembly().CodeBase;
                this.txtInfo.Text = this.Text + "\r\n" + oVer + "\r\n" + wVer + "\r\n" + wPath + "\r\n" + mPath;
                tableLayoutPanel1.Dock = DockStyle.Fill;
                button1.DataBindings.Add("Font", Properties.Settings.Default, "公文标题",true,DataSourceUpdateMode.OnPropertyChanged);
                button2.DataBindings.Add("Font", Properties.Settings.Default, "一级标题", true, DataSourceUpdateMode.OnPropertyChanged);
                button3.DataBindings.Add("Font", Properties.Settings.Default, "二级标题", true, DataSourceUpdateMode.OnPropertyChanged);
                button4.DataBindings.Add("Font", Properties.Settings.Default, "三级标题", true, DataSourceUpdateMode.OnPropertyChanged);
                button5.DataBindings.Add("Font", Properties.Settings.Default, "四级标题", true, DataSourceUpdateMode.OnPropertyChanged);
                button6.DataBindings.Add("Font", Properties.Settings.Default, "正文", true, DataSourceUpdateMode.OnPropertyChanged);
                button7.DataBindings.Add("Font", Properties.Settings.Default, "版头", true, DataSourceUpdateMode.OnPropertyChanged);
                button8.DataBindings.Add("Font", Properties.Settings.Default, "签发人姓名", true, DataSourceUpdateMode.OnPropertyChanged);
                button9.DataBindings.Add("Font", Properties.Settings.Default, "版记", true, DataSourceUpdateMode.OnPropertyChanged);
                button10.DataBindings.Add("Font", Properties.Settings.Default, "页码", true, DataSourceUpdateMode.OnPropertyChanged);
                
                LoadFont(cobLetter);

                // 当TableLayoutPanel控件中的需要更新的Label过多的时候，刷新Label的时候会出现闪烁问题，主要解决办法就是增加双缓冲.
                //但是直接设置窗体双缓冲效果并不好,需要如下代码:
                tableLayoutPanel1.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | 
                    System.Reflection.BindingFlags.NonPublic).SetValue(tableLayoutPanel1, true, null);  

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }


        /// <summary>
        /// 通过指定的URL获取版本更新信息.
        /// </summary>
        /// <param name="url">版本信息文件url地址</param>
        /// <returns>string</returns>
        public static string CheckUpdate(string url)
        {
            string rXml = "未获取!";
            try
            {
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(new Uri(url));
                myHttpWebRequest.KeepAlive = false;
                myHttpWebRequest.AllowAutoRedirect = false;
                myHttpWebRequest.UserAgent = "xkonglong WordAddin Check Update Client"+ Assembly.GetExecutingAssembly().GetName().Version.ToString(); ;
                myHttpWebRequest.Timeout = 10000;
                myHttpWebRequest.ContentType = "application/x-www-form-urlencoded;charset=utf-8";

                using (HttpWebResponse res = (HttpWebResponse)myHttpWebRequest.GetResponse())
                {
                    if (res.StatusCode == HttpStatusCode.OK || res.StatusCode == HttpStatusCode.PartialContent)//返回为200或206
                    {
                        //string dd = res.ContentEncoding;
                        System.IO.Stream strem = res.GetResponseStream();
                        System.IO.StreamReader r = new System.IO.StreamReader(strem);
                        rXml = r.ReadToEnd();
                    }
                    else   //其他则直接返回 http 代码
                    {
                        rXml = "网络错误代码:"+((int)res.StatusCode).ToString();
                    }
                    return rXml;
                }
            }
            catch 
            {
                rXml = "获取信息出错,请检查网络!";
                return rXml;
                throw;
            }
        }

        private void LnkUpdate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (lnkUpdate.Text.Contains("检查更新"))
            {
                string cu = CheckUpdate("https://xkonglong.com/down/update.txt");
                if (cu.Length== 0) { return; }  //如果获取信息为空,退出
                //如果信息包含|,说明有下载地址,分割字符串
                if (cu.Contains("|"))
                {
                    string[] ver = cu.Split('|');
                    try
                    {
                        //比较版本
                        Version over = Assembly.GetExecutingAssembly().GetName().Version; 
                            Debug.WriteLine(over);
                        Version nver = new Version(ver[0]);
                        if ((nver.CompareTo(over)) > 0)
                        {
                            lnkUpdate.Text = "最新版本:" + ver[0] + "下载可用";
                            lnkUpdate.Tag = ver[1];
                            return;
                        }
                        else
                        {
                            lnkUpdate.Text = "您已是最新版本!";
                            lnkUpdate.Enabled = false; //禁止频繁更新,减少网站流量冲击
                        }
                    }
                    catch
                    {
                        lnkUpdate.Text = "未知的版本信息!";
                        lnkUpdate.Enabled = false; //禁止频繁更新,减少网站流量冲击
                        return;
                    }
                }
                else
                {
                    lnkUpdate.Text = "最新版本:" + cu;
                    lnkUpdate.Enabled = false; //禁止频繁更新,减少网站流量冲击
                }
            }
            else if(lnkUpdate.Text.Contains("下载可用"))
            {
                if (lnkUpdate.Tag == "")
                {
                    lnkUpdate.Tag = "https://xkonglong.com";
                }
                try
                {
                    Process.Start(lnkUpdate.Tag.ToString());
                }
                catch(Win32Exception ex)
                {
                    MessageBox.Show(ex.Message + "\r\n请手动打开浏览器输入网址访问:\r\n " + lnkUpdate.Tag);
                }                
            }
        }

        private void TxtInfo_Click(object sender, EventArgs e)//单击文本框复制内容
        {
            TextBox tbox = sender as TextBox;
            try
            {
                ToolTip tp = new ToolTip();
                txtInfo.SelectAll();  //没用,只是对用户的视觉效果
                Clipboard.SetText(txtInfo.Text);
                //tp.Show("系统信息已复制", (TextBox)sender, 1000);
                tp.Show("系统信息已复制", tbox, 1000);
            }
            catch(Exception ex)
            {
                ToolTip tp = new ToolTip();
                //tp.Show("系统信息复制时出错!\r\n"+ex.Message, (TextBox)sender, 1000);
                tp.Show("系统信息复制时出错!\r\n" + ex.Message, tbox, 1000);
                throw;
            }
        }

        private void BtnCustom_Click(object sender, EventArgs e)//高级功能按钮
        {
            //切换显示关于/高级功能
            if (panel1.Visible==true)
            {
                
                this.Height = 400;
                panel1.Visible = false;
                

                lblTip.Visible = true;
                lblTip.Text = "用户设置，双击恢复默认.";

                GetFont();  //读取字体设置
                btnCustom.Text = "<<高级功能";
                btnSave.Visible = true;
                tableLayoutPanel1.Dock = DockStyle.Fill;
                tableLayoutPanel1.Visible = true;
                return;
            }
            else
            {
                lblTip.Visible = false;
                btnSave.Visible = false;
                this.Height = MinimumSize.Height;
                panel1.Visible = true;
                tableLayoutPanel1.Visible = false;
                btnCustom.Text = ">>高级功能";
                
                return;
            }
        }
        private void SetFont(object sender,EventArgs e)  //单击按钮 设置字体
        {
            try
            {
                Button btn = sender as Button;
                FontDialog fd = new FontDialog() ;
                //fd.Font = ((Button)sender).Font;
                fd.Font = btn.Font;
                fd.AllowSimulations = true;
                fd.ShowEffects = false;
                fd.MinSize = 9;
                fd.FontMustExist = true;

                DialogResult dr = fd.ShowDialog(this);
                if (dr == DialogResult.OK)
                {
                    btn.Text = fd.Font.Name;
                    btn.Font = fd.Font;
                }
                Debug.WriteLine("Size:{0}", fd.Font.Size);
                Debug.WriteLine("SizePoint:{0}", fd.Font.SizeInPoints);
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n 当前用户权限无法访问该字体,请为所有用户安装该字体。";
                MessageBox.Show(err, "打开字体对话框时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                throw;
            }
        }

        private void GetFont()//获取配置文件字体
        {
            try
            {
                button1.Font = Properties.Settings.Default.公文标题;
                button1.Text = Properties.Settings.Default.公文标题.Name;

                button2.Font = Properties.Settings.Default.一级标题;
                button2.Text = Properties.Settings.Default.一级标题.Name;

                button3.Font = Properties.Settings.Default.二级标题;
                button3.Text = Properties.Settings.Default.二级标题.Name;

                button4.Font = Properties.Settings.Default.三级标题;
                button4.Text = Properties.Settings.Default.三级标题.Name;

                button5.Font = Properties.Settings.Default.四级标题;
                button5.Text = Properties.Settings.Default.四级标题.Name;

                button6.Font = Properties.Settings.Default.正文;
                button6.Text = Properties.Settings.Default.正文.Name;

                button7.Font = Properties.Settings.Default.版头;
                button7.Text = Properties.Settings.Default.版头.Name;

                button8.Font = Properties.Settings.Default.签发人姓名;
                button8.Text = Properties.Settings.Default.签发人姓名.Name;

                button9.Font = Properties.Settings.Default.版记;
                button9.Text = Properties.Settings.Default.版记.Name;

                button10.Font = Properties.Settings.Default.页码;
                button10.Text = Properties.Settings.Default.页码.Name;

                string letterFont = Properties.Settings.Default.英文字体;
                if (letterFont.Length > 0)
                {
                    cobLetter.Text = letterFont;
                }
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "读取公文格式配置文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                throw;
            }
        }


        private void SaveConfig()//保存配置
        {
            try
            {
                //保存自定义字体格式
                Properties.Settings.Default.公文标题 = button1.Font;
                Properties.Settings.Default.一级标题 = button2.Font;
                Properties.Settings.Default.二级标题 = button3.Font;
                Properties.Settings.Default.三级标题 = button4.Font;
                Properties.Settings.Default.四级标题 = button5.Font;
                Properties.Settings.Default.正文 = button6.Font;
                Properties.Settings.Default.版头 = button7.Font;
                Properties.Settings.Default.签发人姓名 = button8.Font;
                Properties.Settings.Default.版记 = button9.Font;
                Properties.Settings.Default.页码 = button10.Font;

                //保存西文字体
                if (cobLetter.SelectedIndex < 1 || cobLetter.Text == "")
                {
                    MessageBox.Show("未指定默认西文字体!", this.Text);
                    return;
                }
                else
                {
                    Properties.Settings.Default.英文字体 = cobLetter.SelectedItem.ToString();
                }

                //保存配置文件
                Properties.Settings.Default.Save();
                string tip = "用户自定义配置文件已保存！\r\n 用户配置文件在Word更新后会被重置.您可通过双击赞赏码修改默认配置文件.";
                    //+"\r\n   \r\n注意:自定义格式需要关闭所有word程序后生效。";
                MessageBox.Show(tip, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

                //面板恢复到关于状态
                lblTip.Visible = false;
                btnSave.Visible = false;
                this.Height = MinimumSize.Height;
                panel1.Visible = true;
                tableLayoutPanel1.Visible = false;
                btnCustom.Text = ">>高级功能";
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "保存公文格式配置文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }


        }



        private void AboutBox1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            try
            {

                Process.Start(Properties.Settings.Default.帮助_高级功能);
               // this.Close();
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n\r\n" + ex.Source;
                MessageBox.Show(err, "打开在线帮助时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }




        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveConfig();
        }


        //隐藏的技能
        private void logoPictureBox_DoubleClick(object sender, EventArgs e)
        {
            string mPath = Assembly.GetExecutingAssembly().CodeBase+".config";
            mPath = mPath.Remove(0, 8);// 移除file:///
            string exe=mPath.Substring(0, mPath.Length-15);
            exe =exe+  "Resources\\ConfigEditor.exe";
            Debug.WriteLine("配置文件路径:" + mPath);
            Debug.WriteLine("修改器路径:" + exe);
            //try
            //{
                Process.Start(exe, "\"" + mPath + "\"");
                this.Close();
            //}
            //catch (Exception ex)
            //{
            //    string err = ex.Message + "\r\n\r\n" + ex.Source;
            //    MessageBox.Show(err, "打开默认配置文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //}
            //    //MessageBox.Show("修改默认配置文件有风险,请注意备份.","公文排版助手",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                //try
                //{
                //    Process.Start("NotePad.exe", mPath);
                //    this.Close();
                //}
                //catch (Exception ex)
                //{
                //    string err = ex.Message + "\r\n\r\n" + ex.Source;
                //    MessageBox.Show(err, "打开config文件时出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //}

            }

        private void lblTip_DoubleClick(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
        }

        private void AboutBox1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        //  以下代码是直接写配置文件的, 但是读取有问题, 插件载入时也需要读取配置才行. 暂存.
        /*

        static void ReadAppSettings()
        {
            try
            {

                // Get the configuration file.
                System.Configuration.Configuration config =
                    ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                // Get the appSettings section.
                System.Configuration.AppSettingsSection appSettings =
                    (System.Configuration.AppSettingsSection)config.GetSection("appSettings");

                // Get the auxiliary file name.
                Debug.WriteLine("Auxiliary file: {0}", config.AppSettings.File);

                // Get the settings collection (key/value pairs).
                if (appSettings.Settings.Count != 0)
                {
                    foreach (string key in appSettings.Settings.AllKeys)
                    {
                        string value = appSettings.Settings[key].Value;
                        Debug.WriteLine("Key: {0} Value: {1}", key, value);
                    }
                }
                else
                {
                    Debug.WriteLine("The appSettings section is empty. Write first.");
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("Exception raised: {0}", e.Message);
            }
        }
        private void WriteAppSettings()
        {
            try
            {
                // Get the application configuration file.
                System.Configuration.Configuration config =
                   ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                // Create a unique key/value pair to add to 
                // the appSettings section.
                string keyName = "公文标题字体test:";
                string value = this.button1.Font.ToString();

                // Add the key/value pair to the appSettings 
                // section.
                // config.AppSettings.Settings.Add(keyName, value);
                System.Configuration.AppSettingsSection appSettings = config.AppSettings;
                appSettings.Settings.Add(keyName, value);

                // Save the configuration file.
                config.Save(ConfigurationSaveMode.Modified);

                // Force a reload in memory of the changed section.
                // This to read the section with the
                // updated values.
                ConfigurationManager.RefreshSection("appSettings");

                Debug.WriteLine(
                    "Added the following Key: {0} Value: {1} .", keyName, value);
            }
            catch (Exception e)
            {
                Debug.WriteLine("Exception raised: {0}",
                    e.Message);
            }
        }
        */


    }
}
