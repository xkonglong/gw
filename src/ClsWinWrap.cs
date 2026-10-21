using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 公文排版
{
    public class ClsWinWrap : IWin32Window
    {
        private IntPtr m_Handle;

        //构造函数，参数是父窗口的句柄
        public ClsWinWrap(IntPtr handle)
        {
            this.m_Handle = handle;
        }

        //构造函数，参数是父窗口的句柄
        public ClsWinWrap(int handle)
        {
            this.m_Handle = new IntPtr(handle);
        }

        public IntPtr Handle
        {
            get { return m_Handle; }
        }

        //打开窗体，Show参数使用该类自身
        public void Show(Form frm)
        {
            frm.Show(this);
        }

        //模式打开窗体，Show参数使用该类自身
        public DialogResult ShowDialog(Form frm)
        {
            return frm.ShowDialog(this);
        }
    }
}
