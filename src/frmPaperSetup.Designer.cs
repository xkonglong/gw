namespace 公文排版
{
    partial class frmPaperSetup
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
            this.components = new System.ComponentModel.Container();
            this.chkIncludeFonts = new System.Windows.Forms.CheckBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.nudFoot = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.nudTop = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.nudBottom = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.nudLeft = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.nudRight = new System.Windows.Forms.NumericUpDown();
            this.chkShowHideFont = new System.Windows.Forms.CheckBox();
            this.groupBoxGW = new System.Windows.Forms.GroupBox();
            this.picGW = new System.Windows.Forms.PictureBox();
            this.chkLineSpacing = new System.Windows.Forms.CheckBox();
            this.nudLineSpacing = new System.Windows.Forms.NumericUpDown();
            this.nudLetters = new System.Windows.Forms.NumericUpDown();
            this.chkGW = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.nudLine = new System.Windows.Forms.NumericUpDown();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.picShowHiddenText = new System.Windows.Forms.PictureBox();
            this.picSaveFont = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFoot)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudTop)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudBottom)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLeft)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudRight)).BeginInit();
            this.groupBoxGW.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picGW)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLineSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLetters)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLine)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picShowHiddenText)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picSaveFont)).BeginInit();
            this.SuspendLayout();
            // 
            // chkIncludeFonts
            // 
            this.chkIncludeFonts.AutoSize = true;
            this.chkIncludeFonts.Location = new System.Drawing.Point(30, 341);
            this.chkIncludeFonts.Name = "chkIncludeFonts";
            this.chkIncludeFonts.Size = new System.Drawing.Size(144, 16);
            this.chkIncludeFonts.TabIndex = 10;
            this.chkIncludeFonts.Text = "文档中嵌入不常用字体";
            this.chkIncludeFonts.UseVisualStyleBackColor = true;
            // 
            // linkLabel1
            // 
            this.linkLabel1.LinkArea = new System.Windows.Forms.LinkArea(40, 7);
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.AlwaysUnderline;
            this.linkLabel1.Location = new System.Drawing.Point(12, 9);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Padding = new System.Windows.Forms.Padding(1);
            this.linkLabel1.Size = new System.Drawing.Size(363, 40);
            this.linkLabel1.TabIndex = 28;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "提示：本设置针对该文档全局，且会更改纸张尺寸为A4大小。\r\n若有特殊需求，请打开页面设置对话框自行设置。";
            this.linkLabel1.UseCompatibleTextRendering = true;
            this.linkLabel1.Click += new System.EventHandler(this.LinkLabel1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.nudFoot);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.nudTop);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.nudBottom);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.nudLeft);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.nudRight);
            this.groupBox1.Location = new System.Drawing.Point(12, 52);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(363, 190);
            this.groupBox1.TabIndex = 26;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "页边距";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(74, 153);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(91, 15);
            this.label7.TabIndex = 7;
            this.label7.Text = "页脚距边界(mm)";
            // 
            // nudFoot
            // 
            this.nudFoot.DecimalPlaces = 1;
            this.nudFoot.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudFoot.Location = new System.Drawing.Point(169, 151);
            this.nudFoot.Name = "nudFoot";
            this.nudFoot.Size = new System.Drawing.Size(58, 21);
            this.nudFoot.TabIndex = 5;
            this.nudFoot.Value = new decimal(new int[] {
            247,
            0,
            0,
            65536});
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(90, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "上边距(mm)";
            // 
            // nudTop
            // 
            this.nudTop.DecimalPlaces = 1;
            this.nudTop.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudTop.Location = new System.Drawing.Point(169, 27);
            this.nudTop.Name = "nudTop";
            this.nudTop.Size = new System.Drawing.Size(58, 21);
            this.nudTop.TabIndex = 1;
            this.nudTop.Value = new decimal(new int[] {
            37,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(90, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "下边距(mm)";
            // 
            // nudBottom
            // 
            this.nudBottom.DecimalPlaces = 1;
            this.nudBottom.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudBottom.Location = new System.Drawing.Point(169, 113);
            this.nudBottom.Name = "nudBottom";
            this.nudBottom.Size = new System.Drawing.Size(58, 21);
            this.nudBottom.TabIndex = 2;
            this.nudBottom.Value = new decimal(new int[] {
            35,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(24, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 15);
            this.label3.TabIndex = 4;
            this.label3.Text = "左边距(mm)";
            // 
            // nudLeft
            // 
            this.nudLeft.DecimalPlaces = 1;
            this.nudLeft.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudLeft.Location = new System.Drawing.Point(107, 69);
            this.nudLeft.Name = "nudLeft";
            this.nudLeft.Size = new System.Drawing.Size(58, 21);
            this.nudLeft.TabIndex = 3;
            this.nudLeft.Value = new decimal(new int[] {
            28,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(195, 71);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "右边距(mm)";
            // 
            // nudRight
            // 
            this.nudRight.DecimalPlaces = 1;
            this.nudRight.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudRight.Location = new System.Drawing.Point(278, 69);
            this.nudRight.Name = "nudRight";
            this.nudRight.Size = new System.Drawing.Size(58, 21);
            this.nudRight.TabIndex = 4;
            this.nudRight.Value = new decimal(new int[] {
            26,
            0,
            0,
            0});
            // 
            // chkShowHideFont
            // 
            this.chkShowHideFont.AutoSize = true;
            this.chkShowHideFont.Checked = true;
            this.chkShowHideFont.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShowHideFont.Location = new System.Drawing.Point(30, 365);
            this.chkShowHideFont.Name = "chkShowHideFont";
            this.chkShowHideFont.Size = new System.Drawing.Size(228, 16);
            this.chkShowHideFont.TabIndex = 11;
            this.chkShowHideFont.Text = "显示隐藏文字但不打印(勾选直接生效)";
            this.chkShowHideFont.UseVisualStyleBackColor = true;
            this.chkShowHideFont.CheckedChanged += new System.EventHandler(this.ChkShowHideFont_CheckedChanged);
            // 
            // groupBoxGW
            // 
            this.groupBoxGW.Controls.Add(this.picGW);
            this.groupBoxGW.Controls.Add(this.chkLineSpacing);
            this.groupBoxGW.Controls.Add(this.nudLineSpacing);
            this.groupBoxGW.Controls.Add(this.nudLetters);
            this.groupBoxGW.Controls.Add(this.chkGW);
            this.groupBoxGW.Controls.Add(this.label6);
            this.groupBoxGW.Controls.Add(this.nudLine);
            this.groupBoxGW.Location = new System.Drawing.Point(12, 248);
            this.groupBoxGW.Name = "groupBoxGW";
            this.groupBoxGW.Size = new System.Drawing.Size(363, 85);
            this.groupBoxGW.TabIndex = 33;
            this.groupBoxGW.TabStop = false;
            this.groupBoxGW.Text = "间距";
            // 
            // picGW
            // 
            this.picGW.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.picGW.Cursor = System.Windows.Forms.Cursors.Hand;
            this.picGW.Image = global::公文排版.Properties.Resources.Info;
            this.picGW.Location = new System.Drawing.Point(329, 23);
            this.picGW.Margin = new System.Windows.Forms.Padding(0);
            this.picGW.Name = "picGW";
            this.picGW.Size = new System.Drawing.Size(16, 16);
            this.picGW.TabIndex = 40;
            this.picGW.TabStop = false;
            this.toolTip1.SetToolTip(this.picGW, "开启此选项后,可以精确控制每页行数和每行字数.\r\n注意: 页面默认字体将被调整为三号正文. \r\n为了避免排版错乱,请先设置页面后录入并排版。");
            this.picGW.Click += new System.EventHandler(this.ShowWordGW);
            // 
            // chkLineSpacing
            // 
            this.chkLineSpacing.AutoSize = true;
            this.chkLineSpacing.Checked = true;
            this.chkLineSpacing.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLineSpacing.Location = new System.Drawing.Point(18, 55);
            this.chkLineSpacing.Name = "chkLineSpacing";
            this.chkLineSpacing.Size = new System.Drawing.Size(162, 16);
            this.chkLineSpacing.TabIndex = 8;
            this.chkLineSpacing.Text = "设置行间距为固定值(磅):";
            this.chkLineSpacing.UseVisualStyleBackColor = true;
            this.chkLineSpacing.CheckedChanged += new System.EventHandler(this.ChkLineSpacing_CheckedChanged);
            // 
            // nudLineSpacing
            // 
            this.nudLineSpacing.DecimalPlaces = 2;
            this.nudLineSpacing.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudLineSpacing.Location = new System.Drawing.Point(187, 50);
            this.nudLineSpacing.Name = "nudLineSpacing";
            this.nudLineSpacing.Size = new System.Drawing.Size(74, 21);
            this.nudLineSpacing.TabIndex = 9;
            this.nudLineSpacing.Value = new decimal(new int[] {
            28,
            0,
            0,
            0});
            // 
            // nudLetters
            // 
            this.nudLetters.Location = new System.Drawing.Point(261, 20);
            this.nudLetters.Name = "nudLetters";
            this.nudLetters.Size = new System.Drawing.Size(58, 21);
            this.nudLetters.TabIndex = 7;
            this.nudLetters.Value = new decimal(new int[] {
            28,
            0,
            0,
            0});
            // 
            // chkGW
            // 
            this.chkGW.AutoSize = true;
            this.chkGW.Location = new System.Drawing.Point(18, 24);
            this.chkGW.Name = "chkGW";
            this.chkGW.Size = new System.Drawing.Size(120, 16);
            this.chkGW.TabIndex = 37;
            this.chkGW.Text = "精确控制每页行数";
            this.chkGW.UseVisualStyleBackColor = true;
            this.chkGW.CheckedChanged += new System.EventHandler(this.ChkGW_CheckedChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(202, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 36;
            this.label6.Text = "每行字数";
            // 
            // nudLine
            // 
            this.nudLine.Location = new System.Drawing.Point(138, 20);
            this.nudLine.Name = "nudLine";
            this.nudLine.Size = new System.Drawing.Size(58, 21);
            this.nudLine.TabIndex = 6;
            this.nudLine.Value = new decimal(new int[] {
            22,
            0,
            0,
            0});
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnOK);
            this.panel1.Controls.Add(this.btnApply);
            this.panel1.Controls.Add(this.btnReset);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 399);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(387, 46);
            this.panel1.TabIndex = 34;
            // 
            // btnOK
            // 
            this.btnOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOK.Location = new System.Drawing.Point(206, 11);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(78, 23);
            this.btnOK.TabIndex = 12;
            this.btnOK.Text = "确定(&O)";
            this.btnOK.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(300, 11);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 23);
            this.btnApply.TabIndex = 13;
            this.btnApply.Text = "应用(&A)";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(12, 11);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 99;
            this.btnReset.Tag = "";
            this.btnReset.Text = "重置(&R)";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.BtnReset_Click);
            // 
            // toolTip1
            // 
            this.toolTip1.AutoPopDelay = 50000;
            this.toolTip1.InitialDelay = 500;
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ReshowDelay = 100;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.toolTip1.ToolTipTitle = "小恐龙温馨提示:";
            // 
            // picShowHiddenText
            // 
            this.picShowHiddenText.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.picShowHiddenText.Cursor = System.Windows.Forms.Cursors.Hand;
            this.picShowHiddenText.Image = global::公文排版.Properties.Resources.Info;
            this.picShowHiddenText.Location = new System.Drawing.Point(278, 365);
            this.picShowHiddenText.Margin = new System.Windows.Forms.Padding(0);
            this.picShowHiddenText.Name = "picShowHiddenText";
            this.picShowHiddenText.Size = new System.Drawing.Size(16, 16);
            this.picShowHiddenText.TabIndex = 39;
            this.picShowHiddenText.TabStop = false;
            this.toolTip1.SetToolTip(this.picShowHiddenText, "勾选此选项,将会开启显示隐藏文字,并取消打印隐藏文字选项.\r\n此选项即时生效,但不会保存为默认设置.. \r\n若取消后,仍显示隐藏文件,请在Word选项中取消勾选 " +
        "\"显示所有格式标记\".");
            this.picShowHiddenText.Click += new System.EventHandler(this.ShowHiddenText);
            // 
            // picSaveFont
            // 
            this.picSaveFont.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.picSaveFont.Cursor = System.Windows.Forms.Cursors.Hand;
            this.picSaveFont.Image = global::公文排版.Properties.Resources.Info;
            this.picSaveFont.Location = new System.Drawing.Point(192, 341);
            this.picSaveFont.Margin = new System.Windows.Forms.Padding(0);
            this.picSaveFont.Name = "picSaveFont";
            this.picSaveFont.Size = new System.Drawing.Size(16, 16);
            this.picSaveFont.TabIndex = 38;
            this.picSaveFont.TabStop = false;
            this.toolTip1.SetToolTip(this.picSaveFont, "嵌入不常用字体,比如\"小标宋\",这并不会大幅增加文档大小.");
            this.picSaveFont.Click += new System.EventHandler(this.SaveFontInDoc);
            // 
            // frmPaperSetup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(387, 445);
            this.Controls.Add(this.picShowHiddenText);
            this.Controls.Add(this.picSaveFont);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.groupBoxGW);
            this.Controls.Add(this.chkShowHideFont);
            this.Controls.Add(this.chkIncludeFonts);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.HelpButton = true;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmPaperSetup";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "小恐龙公文排版助手-页面设置";
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.FrmPaperSetup_HelpButtonClicked);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmPaperSetup_FormClosed);
            this.Load += new System.EventHandler(this.FrmPaperSetup_Load);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FrmPaperSetup_KeyPress);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudFoot)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudTop)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudBottom)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLeft)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudRight)).EndInit();
            this.groupBoxGW.ResumeLayout(false);
            this.groupBoxGW.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picGW)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLineSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLetters)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLine)).EndInit();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picShowHiddenText)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picSaveFont)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkIncludeFonts;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nudTop;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown nudBottom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown nudLeft;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown nudRight;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown nudFoot;
        private System.Windows.Forms.CheckBox chkShowHideFont;
        private System.Windows.Forms.GroupBox groupBoxGW;
        private System.Windows.Forms.CheckBox chkLineSpacing;
        private System.Windows.Forms.NumericUpDown nudLineSpacing;
        private System.Windows.Forms.NumericUpDown nudLetters;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown nudLine;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.CheckBox chkGW;
        private System.Windows.Forms.PictureBox picSaveFont;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.PictureBox picShowHiddenText;
        private System.Windows.Forms.PictureBox picGW;
    }
}