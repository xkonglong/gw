using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace 公文排版
{
    public partial class Ribbon1
    {
        private void DocPassWord(string passWord)
        {
            try { WordApp.ActiveDocument.Password = passWord; }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnProtectDoc_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                WordApp.TaskPanes[MSWord.WdTaskPanes.wdTaskPaneDocumentProtection].Visible = true;
                if (WordApp.ActiveDocument.ProtectionType == MSWord.WdProtectionType.wdNoProtection)
                {
                    WordApp.Dialogs[MSWord.WdWordDialog.wdDialogToolsProtectDocument].Show();
                }
                btnProtectDoc.Checked = false;
            }
            catch (Exception ex)
            {
                WordApp.StatusBar = ex.Message;
            }
        }

        private void btnSetPW_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                WordApp.Dialogs[MSWord.WdWordDialog.wdDialogToolsOptionsSecurity].Show();
            }
            catch (Exception ex)
            {
                if (IsWps)
                {
                    MessageBox.Show("对不起,您的WPS暂时还不兼容这一调用内置对话框的方法!", "打开内置对话框时出错");
                    btnSetPW.Enabled = false;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void btnClaerPW_Click(object sender, RibbonControlEventArgs e)
        {
            var dr = MessageBox.Show("你确定要清除本文档的\"打开密码\"吗?\r\n 点击确定将直接清除密码.",
                "文档加密", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                DocPassWord("");
                WordApp.StatusBar = "警告:文档打开密码已被清除!";
            }
        }

        private void btnClearProtectPW_Click(object sender, RibbonControlEventArgs e)
        {
            var dr = MessageBox.Show("你确定要清除本文档的\"修改密码\"吗?\r\n 点击确定将直接清除密码.",
                "文档加密", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                WordApp.ActiveDocument.WritePassword = "";
                WordApp.StatusBar = "警告:文档修改密码已被清除!";
            }
        }
    }
}