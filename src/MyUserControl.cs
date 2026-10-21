using System;
using System.Windows.Forms;
using System.Diagnostics;

namespace 公文排版
{
    public partial class MyUserControl : UserControl
    {
        public MyUserControl()
        {
            InitializeComponent();
        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {
        }

        private void Button_Click(object sender, EventArgs e)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                if (wordApp == null) return;
                wordApp.Selection.TypeText(((Button)sender).Text);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Button_Click: " + ex.Message);
            }
        }
    }
}