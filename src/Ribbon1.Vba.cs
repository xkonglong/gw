using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Vbe.Interop;
using MSWord = Microsoft.Office.Interop.Word;
using MSOffice = Microsoft.Office.Core;

namespace 公文排版
{
    public partial class Ribbon1
    {
        private void OpenVba()
        {
            try
            {
                string vbaFile;
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Filter = "VBA文件(*.vba,*.bas)|*.vba;*.bas|全部文件(*.*)|*.*";
                    if (ofd.ShowDialog() != DialogResult.OK) return;
                    vbaFile = ofd.FileName;
                }

                if (!CheckVbaEncoding(vbaFile)) return;

                MSWord.Document doc = WordApp.ActiveDocument;
                VBComponents vbc = doc.VBProject.VBComponents;
                VBComponent vbm = vbc.Add(vbext_ComponentType.vbext_ct_StdModule);
                vbm.Name = "小恐龙公文排版助手";
                CodeModule module = vbm.CodeModule;
                module.AddFromFile(vbaFile);

                vbext_ProcKind ProcKind;
                int startLine = module.CountOfDeclarationLines + 1;
                string macroName = module.get_ProcOfLine(startLine, out ProcKind);

                var stopwatch = new Stopwatch();
                stopwatch.Start();
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("运行VBA:" + macroName);
                try { WordApp.Run(macroName); }
                finally { undo.EndCustomRecord(); vbc.Remove(vbm); }
                stopwatch.Stop();
                WordApp.StatusBar = $"VBA文件执行共用时:{stopwatch.Elapsed.TotalSeconds:F2} 秒";

                AddVbaHistory(vbaFile);
            }
            catch (COMException ex)
            {
                HandleVbaComException(ex);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "执行VBA代码出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnVBA_Click(object sender, RibbonControlEventArgs e)
        {
            //if (sender is RibbonButton btn) RunVBA(btn.ScreenTip);
            OpenVba();
        }

        private bool RunVBA(string vbaFile)
        {
            try
            {
                if (!CheckVbaEncoding(vbaFile)) return false;

                MSWord.Document doc = WordApp.ActiveDocument;
                VBComponents vbc = doc.VBProject.VBComponents;
                VBComponent vbm = vbc.Add(vbext_ComponentType.vbext_ct_StdModule);
                vbm.Name = "小恐龙公文排版助手";
                CodeModule module = vbm.CodeModule;
                module.AddFromFile(vbaFile);

                vbext_ProcKind ProcKind;
                int startLine = module.CountOfDeclarationLines + 1;
                string macroName = module.get_ProcOfLine(startLine, out ProcKind);

                var stopwatch = new Stopwatch();
                stopwatch.Start();
                var undo = WordApp.UndoRecord;
                undo.StartCustomRecord("运行VBA:" + macroName);
                try { WordApp.Run(macroName); }
                finally { undo.EndCustomRecord(); vbc.Remove(vbm); }
                stopwatch.Stop();
                WordApp.StatusBar = $"VBA文件执行共用时:{stopwatch.Elapsed.TotalSeconds:F2} 秒";
                return true;
            }
            catch (Exception ex)
            {
                WordApp.StatusBar = ex.Message;
                return false;
            }
        }

        private void BtnVBAlist_Click(object sender, RibbonControlEventArgs e)
        {
            if (sender is RibbonButton btn) RunVBA(btn.ScreenTip);
        }

        private bool CheckVbaEncoding(string vbaFile)
        {
            Encoding encoding = GetTextFileEncodingType(vbaFile);
            if (encoding == Encoding.Unicode || encoding == Encoding.UTF8)
            {
                var result = MessageBox.Show(
                    "您所选文件非ANSI编码,WordVBA不支持UTF8编码。\r\n您要用记事本打开并查看文件编码吗?\r\n单击确定将运行记事本程序,取消退出。",
                    "读取VBA文件", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.OK)
                {
                    try { Process.Start("notepad", vbaFile); }
                    catch (Exception ex) { Debug.WriteLine(ex.Message); }
                }
                return false;
            }
            return true;
        }

        private void AddVbaHistory(string vbaFile)
        {
            try
            {
                string fileNameNoExt = Path.GetFileNameWithoutExtension(vbaFile);
                string item = fileNameNoExt + "|" + vbaFile;
                if (Properties.Settings.Default.VBA历史 == null)
                    Properties.Settings.Default.VBA历史 = new StringCollection();
                if (Properties.Settings.Default.VBA历史.Contains(item)) return;

                var btnVBAlist = this.Factory.CreateRibbonButton();
                btnVBAlist.Label = fileNameNoExt;
                btnVBAlist.ScreenTip = vbaFile;
                btnVBAlist.ControlSize = MSOffice.RibbonControlSize.RibbonControlSizeRegular;
                btnVBAlist.Click += BtnVBAlist_Click;
                mnuScript.Visible = true;
                mnuScript.Items.Add(btnVBAlist);
                Properties.Settings.Default.VBA历史.Add(item);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "添加VBA历史记录时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void HandleVbaComException(COMException ex)
        {
            string err = "";
            string helpURL;
            switch (ex.ErrorCode)
            {
                case -2146822220:
                    err = "Word 未开启信任对VBA工程对象模型的访问。\r\n\r\n点击[确定]获取更多信息！";
                    helpURL = "https://gw.xkonglong.com/#/vba";
                    break;
                case -2147467259:
                    err = "WPS用户请先切换至【VB环境】,并在【宏安全性】选项中开启【信任对VBA工程对象模型的访问】\r\n\r\n点击[确定]获取更多信息！";
                    helpURL = "https://gw.xkonglong.com/#/vba";
                    break;
                case -2147352573:
                    err = ex.Message + "\r\n\r\n点击[确定]获取更多信息！";
                    helpURL = "https://gw.xkonglong.com/#/vba";
                    break;
                default:
                    err = "错误代码:" + ex.ErrorCode + "\r\n错误信息:" + ex.Message;
                    Clipboard.SetText(err);
                    err += "\r\n\r\n[错误信息已复制到剪贴板]\r\n点击[确定]向作者反馈问题！";
                    helpURL = "mailto:xkonglong@qq.com?subject=小恐龙公文助手问题反馈";
                    break;
            }
            if (MessageBox.Show(err, "执行VBA代码出错", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                Process.Start(helpURL);
        }

        private static Encoding GetTextFileEncodingType(string fileName)
        {
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                using (var br = new BinaryReader(fs, Encoding.Default))
                {
                    byte[] buffer = br.ReadBytes((int)fs.Length);
                    if (buffer.Length >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
                        return Encoding.UTF8;
                    if (buffer.Length >= 3 && buffer[0] == 0xFE && buffer[1] == 0xFF && buffer[2] == 0x00)
                        return Encoding.BigEndianUnicode;
                    if (buffer.Length >= 3 && buffer[0] == 0xFF && buffer[1] == 0xFE && buffer[2] == 0x41)
                        return Encoding.Unicode;
                    if (IsUTF8Bytes(buffer)) return Encoding.UTF8;
                    return Encoding.Default;
                }
            }
            catch { return Encoding.Default; }
        }

        private static bool IsUTF8Bytes(byte[] data)
        {
            int charByteCounter = 1;
            for (int i = 0; i < data.Length; i++)
            {
                byte curByte = data[i];
                if (charByteCounter == 1)
                {
                    if (curByte >= 0x80)
                    {
                        while (((curByte <<= 1) & 0x80) != 0) charByteCounter++;
                        if (charByteCounter == 1 || charByteCounter > 6) return false;
                    }
                }
                else
                {
                    if ((curByte & 0xC0) != 0x80) return false;
                    charByteCounter--;
                }
            }
            return charByteCounter <= 1;
        }

        private void clearVBA()
        {
            try
            {
                MSWord.Document doc = WordApp.ActiveDocument;
                VBComponents vbcs = doc.VBProject.VBComponents;
                var names = new StringBuilder();

                for (int i = vbcs.Count; i >= 1; i--)
                {
                    VBComponent vbc = vbcs.Item(i);
                    names.AppendLine(vbc.Name);
                    switch (vbc.Type)
                    {
                        case vbext_ComponentType.vbext_ct_StdModule:
                        case vbext_ComponentType.vbext_ct_ClassModule:
                        case vbext_ComponentType.vbext_ct_MSForm:
                            vbcs.Remove(vbc);
                            break;
                        case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                        case vbext_ComponentType.vbext_ct_Document:
                            vbc.CodeModule.DeleteLines(1, vbc.CodeModule.CountOfLines);
                            break;
                    }
                }
                MessageBox.Show("当前文档中:\r\n" + names + "已清理", "清理VBA代码",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "清理VBA代码时出错",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void btnClearVBA_Click(object sender, RibbonControlEventArgs e)
        {
            clearVBA();
        }
        private void LoadVBA(string vbaFolder = "")
        {
            if (string.IsNullOrEmpty(vbaFolder))
            {
                try
                {
                    vbaFolder = Properties.Settings.Default.VBA路径;
                    if (string.IsNullOrEmpty(vbaFolder))
                        vbaFolder = Path.Combine(GetDllPath() ?? string.Empty, "VBA脚本");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "获取VBA脚本路径时出错",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }

            mnuVBA.Items.Clear();
            if (!Directory.Exists(vbaFolder)) return;

            foreach (string file in Directory.GetFiles(vbaFolder))
            {
                if (Path.GetExtension(file).ToLower() != ".vba") continue;
                var btn = this.Factory.CreateRibbonButton();
                btn.Label = Path.GetFileNameWithoutExtension(file);
                btn.Tag = file;
                btn.ScreenTip = file;
                btn.Click += BtnVBAlist_Click;
                mnuVBA.Items.Add(btn);
            }
        }
    }
}