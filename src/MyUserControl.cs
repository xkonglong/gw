using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 公文排版
{
    public partial class MyUserControl : UserControl
    {
        Microsoft.Office.Interop.Word.Application WordApp;
        
        public MyUserControl()
        {
            InitializeComponent();
        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {
            //string[] syml = { "㎡","m³","‰", "℃", "°", "Ω", "Φ","φ",
            //    "∵", "∴",
            //    "+", "-", "×", "÷", "≤", "≥", "≠", "≈",
            //    "≌", "±", "∑", "∏", "∈",  "√", "✘" };
            //for(int i = 0; i < syml.Length; i++)
            //{
            //    //Label label = new Label();
            //    //label.Text = syml[i];
            //    //flowLayoutPanel1.Controls.Add(label);
            //    Button button = new Button();
            //    button.Text = syml[i];
            //    button.Size = new Size(32, 32);
            //    button.FlatStyle = FlatStyle.Flat;
            //    button.Click += Button_Click;

            //    flowLayoutPanel1.Controls.Add(button);
            //}
            WordApp = Globals.ThisAddIn.Application;
        }

        private void Button_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.TypeText(((Button)sender).Text);
            
        }


    }
}
