using System;
using System.Windows.Forms;

namespace 公文排版
{
    public class ClsWinWrap : IWin32Window
    {
        private readonly IntPtr m_Handle;

        public ClsWinWrap(IntPtr handle)
        {
            this.m_Handle = handle;
        }

        public ClsWinWrap(int handle)
        {
            this.m_Handle = new IntPtr(handle);
        }

        public IntPtr Handle
        {
            get { return m_Handle; }
        }

        public void Show(Form frm)
        {
            if (frm == null) throw new ArgumentNullException(nameof(frm));
            frm.Show(this);
        }

        public DialogResult ShowDialog(Form frm)
        {
            if (frm == null) throw new ArgumentNullException(nameof(frm));
            return frm.ShowDialog(this);
        }
    }
}