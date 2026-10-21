namespace 公文排版
{
    partial class FrmFontChangeDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.cobNumber = new System.Windows.Forms.ComboBox();
            this.chkNumber = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkLetter = new System.Windows.Forms.CheckBox();
            this.cobLetter = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkboxFont = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cobFont1 = new System.Windows.Forms.ComboBox();
            this.cobFont2 = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnOK);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 229);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(398, 43);
            this.panel1.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(311, 9);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 27);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Enabled = false;
            this.btnOK.Location = new System.Drawing.Point(210, 9);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 27);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // cobNumber
            // 
            this.cobNumber.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cobNumber.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cobNumber.Enabled = false;
            this.cobNumber.FormattingEnabled = true;
            this.cobNumber.Location = new System.Drawing.Point(198, 27);
            this.cobNumber.Name = "cobNumber";
            this.cobNumber.Size = new System.Drawing.Size(153, 20);
            this.cobNumber.TabIndex = 1;
            this.cobNumber.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.CobNumber_HelpRequested);
            // 
            // chkNumber
            // 
            this.chkNumber.AutoSize = true;
            this.chkNumber.Location = new System.Drawing.Point(20, 31);
            this.chkNumber.Name = "chkNumber";
            this.chkNumber.Size = new System.Drawing.Size(162, 16);
            this.chkNumber.TabIndex = 2;
            this.chkNumber.Text = "更改数字(比如&123)的字体";
            this.chkNumber.UseVisualStyleBackColor = true;
            this.chkNumber.CheckedChanged += new System.EventHandler(this.ChkNumber_CheckedChanged);
            this.chkNumber.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.ChkNumber_HelpRequested);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 18);
            this.label1.Margin = new System.Windows.Forms.Padding(0, 10, 10, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(281, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "使用前，请勾选要更改的类型，然后选择指定字体。";
            // 
            // chkLetter
            // 
            this.chkLetter.AutoSize = true;
            this.chkLetter.Location = new System.Drawing.Point(20, 68);
            this.chkLetter.Name = "chkLetter";
            this.chkLetter.Size = new System.Drawing.Size(162, 16);
            this.chkLetter.TabIndex = 4;
            this.chkLetter.Text = "更改字母(比如&abc)的字体";
            this.chkLetter.UseVisualStyleBackColor = true;
            this.chkLetter.CheckedChanged += new System.EventHandler(this.ChkLetter_CheckedChanged);
            this.chkLetter.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.ChkLetter_HelpRequested);
            // 
            // cobLetter
            // 
            this.cobLetter.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cobLetter.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cobLetter.Enabled = false;
            this.cobLetter.FormattingEnabled = true;
            this.cobLetter.ItemHeight = 12;
            this.cobLetter.Location = new System.Drawing.Point(198, 66);
            this.cobLetter.Name = "cobLetter";
            this.cobLetter.Size = new System.Drawing.Size(153, 20);
            this.cobLetter.TabIndex = 3;
            this.cobLetter.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.CobLetter_HelpRequested);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.groupBox2);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.groupBox1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(398, 229);
            this.panel2.TabIndex = 4;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkboxFont);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.cobFont1);
            this.groupBox2.Controls.Add(this.cobFont2);
            this.groupBox2.Location = new System.Drawing.Point(12, 43);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(379, 66);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "groupBox2";
            // 
            // chkboxFont
            // 
            this.chkboxFont.AutoSize = true;
            this.chkboxFont.Location = new System.Drawing.Point(6, 0);
            this.chkboxFont.Name = "chkboxFont";
            this.chkboxFont.Size = new System.Drawing.Size(126, 16);
            this.chkboxFont.TabIndex = 11;
            this.chkboxFont.Text = "按指定字体替换(&S)";
            this.chkboxFont.UseVisualStyleBackColor = true;
            this.chkboxFont.CheckedChanged += new System.EventHandler(this.ChkboxFont_CheckedChanged);
            this.chkboxFont.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.ChkboxFont_HelpRequested);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(173, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(23, 12);
            this.label2.TabIndex = 12;
            this.label2.Text = "==>";
            // 
            // cobFont1
            // 
            this.cobFont1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cobFont1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cobFont1.Enabled = false;
            this.cobFont1.FormattingEnabled = true;
            this.cobFont1.Location = new System.Drawing.Point(16, 31);
            this.cobFont1.Name = "cobFont1";
            this.cobFont1.Size = new System.Drawing.Size(153, 20);
            this.cobFont1.TabIndex = 10;
            this.cobFont1.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.CobFont1_HelpRequested);
            // 
            // cobFont2
            // 
            this.cobFont2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cobFont2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cobFont2.Enabled = false;
            this.cobFont2.FormattingEnabled = true;
            this.cobFont2.Location = new System.Drawing.Point(198, 31);
            this.cobFont2.Name = "cobFont2";
            this.cobFont2.Size = new System.Drawing.Size(153, 20);
            this.cobFont2.TabIndex = 9;
            this.cobFont2.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.CobFont2_HelpRequested);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cobLetter);
            this.groupBox1.Controls.Add(this.chkLetter);
            this.groupBox1.Controls.Add(this.cobNumber);
            this.groupBox1.Controls.Add(this.chkNumber);
            this.groupBox1.Location = new System.Drawing.Point(13, 115);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(379, 108);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "按类型更改字体";
            // 
            // frmFontChangeDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(398, 272);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFontChangeDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "字体更改对话框";
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.FrmFontChangeDialog_HelpButtonClicked);
            this.Load += new System.EventHandler(this.FrmFontChangeDialog_Load);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ComboBox cobNumber;
        private System.Windows.Forms.CheckBox chkNumber;
        private System.Windows.Forms.CheckBox chkLetter;
        private System.Windows.Forms.ComboBox cobLetter;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkboxFont;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cobFont1;
        private System.Windows.Forms.ComboBox cobFont2;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}