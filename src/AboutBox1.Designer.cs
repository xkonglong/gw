namespace 公文排版
{
    partial class AboutBox1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutBox1));
            this.txtInfo = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.logoPictureBox = new System.Windows.Forms.PictureBox();
            this.lblText = new System.Windows.Forms.Label();
            this.lnkUpdate = new System.Windows.Forms.LinkLabel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.lblLetter = new System.Windows.Forms.Label();
            this.lblPageNumber = new System.Windows.Forms.Label();
            this.button10 = new System.Windows.Forms.Button();
            this.lblFormatSong = new System.Windows.Forms.Label();
            this.button9 = new System.Windows.Forms.Button();
            this.lblFormatKai = new System.Windows.Forms.Label();
            this.button8 = new System.Windows.Forms.Button();
            this.lblFormatHei = new System.Windows.Forms.Label();
            this.button7 = new System.Windows.Forms.Button();
            this.lblFormatNormal = new System.Windows.Forms.Label();
            this.button6 = new System.Windows.Forms.Button();
            this.lblFormatTitle4 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.lblFormatTitle3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.lblFormatTitle2 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.lblFormatTitle1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.lblFormatTitle = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.cobLetter = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnCustom = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblTip = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtInfo
            // 
            this.txtInfo.Location = new System.Drawing.Point(247, 99);
            this.txtInfo.Margin = new System.Windows.Forms.Padding(6, 3, 3, 3);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.ReadOnly = true;
            this.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtInfo.Size = new System.Drawing.Size(282, 138);
            this.txtInfo.TabIndex = 30;
            this.txtInfo.TabStop = false;
            this.txtInfo.Text = "说明";
            this.txtInfo.Click += new System.EventHandler(this.TxtInfo_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnOK.Location = new System.Drawing.Point(447, 5);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 23);
            this.btnOK.TabIndex = 31;
            this.btnOK.Text = "关闭(&C)";
            // 
            // logoPictureBox
            // 
            this.logoPictureBox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.logoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("logoPictureBox.Image")));
            this.logoPictureBox.Location = new System.Drawing.Point(0, 3);
            this.logoPictureBox.Name = "logoPictureBox";
            this.logoPictureBox.Size = new System.Drawing.Size(244, 253);
            this.logoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.logoPictureBox.TabIndex = 26;
            this.logoPictureBox.TabStop = false;
            this.logoPictureBox.DoubleClick += new System.EventHandler(this.logoPictureBox_DoubleClick);
            // 
            // lblText
            // 
            this.lblText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblText.Location = new System.Drawing.Point(247, 3);
            this.lblText.Name = "lblText";
            this.lblText.Size = new System.Drawing.Size(282, 93);
            this.lblText.TabIndex = 33;
            this.lblText.Text = "           小恐龙工作室荣誉出品\r\n\r\n  小恐龙公文排版插件适用于 Word20010-2021和365版本,WPS Office2013-2019。" +
    " \r\n\r\n本插件不含任何恶意代码，完全免费，接受捐赠！";
            // 
            // lnkUpdate
            // 
            this.lnkUpdate.AutoSize = true;
            this.lnkUpdate.Location = new System.Drawing.Point(250, 242);
            this.lnkUpdate.Name = "lnkUpdate";
            this.lnkUpdate.Size = new System.Drawing.Size(71, 12);
            this.lnkUpdate.TabIndex = 34;
            this.lnkUpdate.TabStop = true;
            this.lnkUpdate.Text = "检查更新...";
            this.lnkUpdate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LnkUpdate_LinkClicked);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoScroll = true;
            this.tableLayoutPanel1.AutoScrollMargin = new System.Drawing.Size(10, 0);
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 44.44444F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 55.55556F));
            this.tableLayoutPanel1.Controls.Add(this.lblLetter, 0, 10);
            this.tableLayoutPanel1.Controls.Add(this.lblPageNumber, 0, 9);
            this.tableLayoutPanel1.Controls.Add(this.button10, 1, 9);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatSong, 0, 8);
            this.tableLayoutPanel1.Controls.Add(this.button9, 1, 8);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatKai, 0, 7);
            this.tableLayoutPanel1.Controls.Add(this.button8, 1, 7);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatHei, 0, 6);
            this.tableLayoutPanel1.Controls.Add(this.button7, 1, 6);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatNormal, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.button6, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatTitle4, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.button5, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatTitle3, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.button4, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatTitle2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.button3, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatTitle1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button2, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.lblFormatTitle, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.button1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.cobLetter, 1, 10);
            this.tableLayoutPanel1.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(9, 315);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 11;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(529, 307);
            this.tableLayoutPanel1.TabIndex = 36;
            this.tableLayoutPanel1.Visible = false;
            // 
            // lblLetter
            // 
            this.lblLetter.AutoSize = true;
            this.lblLetter.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblLetter.Location = new System.Drawing.Point(4, 285);
            this.lblLetter.Name = "lblLetter";
            this.lblLetter.Size = new System.Drawing.Size(227, 21);
            this.lblLetter.TabIndex = 20;
            this.lblLetter.Text = "西文字体(公文页面设置时生效)";
            this.lblLetter.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblPageNumber
            // 
            this.lblPageNumber.AutoSize = true;
            this.lblPageNumber.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblPageNumber.Location = new System.Drawing.Point(4, 257);
            this.lblPageNumber.Name = "lblPageNumber";
            this.lblPageNumber.Size = new System.Drawing.Size(227, 27);
            this.lblPageNumber.TabIndex = 18;
            this.lblPageNumber.Text = "页码";
            this.lblPageNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button10
            // 
            this.button10.AutoSize = true;
            this.button10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button10.Location = new System.Drawing.Point(235, 257);
            this.button10.Margin = new System.Windows.Forms.Padding(0);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(293, 27);
            this.button10.TabIndex = 19;
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatSong
            // 
            this.lblFormatSong.AutoSize = true;
            this.lblFormatSong.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatSong.Location = new System.Drawing.Point(4, 229);
            this.lblFormatSong.Name = "lblFormatSong";
            this.lblFormatSong.Size = new System.Drawing.Size(227, 27);
            this.lblFormatSong.TabIndex = 16;
            this.lblFormatSong.Text = "四号仿宋(版记)";
            this.lblFormatSong.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button9
            // 
            this.button9.AutoSize = true;
            this.button9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button9.Location = new System.Drawing.Point(235, 229);
            this.button9.Margin = new System.Windows.Forms.Padding(0);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(293, 27);
            this.button9.TabIndex = 17;
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatKai
            // 
            this.lblFormatKai.AutoSize = true;
            this.lblFormatKai.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatKai.Location = new System.Drawing.Point(4, 201);
            this.lblFormatKai.Name = "lblFormatKai";
            this.lblFormatKai.Size = new System.Drawing.Size(227, 27);
            this.lblFormatKai.TabIndex = 14;
            this.lblFormatKai.Text = "三号楷体(签发人姓名)";
            this.lblFormatKai.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button8
            // 
            this.button8.AutoSize = true;
            this.button8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button8.Location = new System.Drawing.Point(235, 201);
            this.button8.Margin = new System.Windows.Forms.Padding(0);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(293, 27);
            this.button8.TabIndex = 15;
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatHei
            // 
            this.lblFormatHei.AutoSize = true;
            this.lblFormatHei.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatHei.Location = new System.Drawing.Point(4, 173);
            this.lblFormatHei.Name = "lblFormatHei";
            this.lblFormatHei.Size = new System.Drawing.Size(227, 27);
            this.lblFormatHei.TabIndex = 12;
            this.lblFormatHei.Text = "三号黑体(版头)";
            this.lblFormatHei.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button7
            // 
            this.button7.AutoSize = true;
            this.button7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button7.Location = new System.Drawing.Point(235, 173);
            this.button7.Margin = new System.Windows.Forms.Padding(0);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(293, 27);
            this.button7.TabIndex = 13;
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatNormal
            // 
            this.lblFormatNormal.AutoSize = true;
            this.lblFormatNormal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatNormal.Location = new System.Drawing.Point(4, 145);
            this.lblFormatNormal.Name = "lblFormatNormal";
            this.lblFormatNormal.Size = new System.Drawing.Size(227, 27);
            this.lblFormatNormal.TabIndex = 10;
            this.lblFormatNormal.Text = "正文";
            this.lblFormatNormal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button6
            // 
            this.button6.AutoSize = true;
            this.button6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button6.Location = new System.Drawing.Point(235, 145);
            this.button6.Margin = new System.Windows.Forms.Padding(0);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(293, 27);
            this.button6.TabIndex = 11;
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatTitle4
            // 
            this.lblFormatTitle4.AutoSize = true;
            this.lblFormatTitle4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatTitle4.Location = new System.Drawing.Point(4, 117);
            this.lblFormatTitle4.Name = "lblFormatTitle4";
            this.lblFormatTitle4.Size = new System.Drawing.Size(227, 27);
            this.lblFormatTitle4.TabIndex = 8;
            this.lblFormatTitle4.Text = "标题（1）";
            this.lblFormatTitle4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button5
            // 
            this.button5.AutoSize = true;
            this.button5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button5.Location = new System.Drawing.Point(235, 117);
            this.button5.Margin = new System.Windows.Forms.Padding(0);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(293, 27);
            this.button5.TabIndex = 9;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatTitle3
            // 
            this.lblFormatTitle3.AutoSize = true;
            this.lblFormatTitle3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatTitle3.Location = new System.Drawing.Point(4, 89);
            this.lblFormatTitle3.Name = "lblFormatTitle3";
            this.lblFormatTitle3.Size = new System.Drawing.Size(227, 27);
            this.lblFormatTitle3.TabIndex = 6;
            this.lblFormatTitle3.Text = "标题1.";
            this.lblFormatTitle3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button4
            // 
            this.button4.AutoSize = true;
            this.button4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button4.Location = new System.Drawing.Point(235, 89);
            this.button4.Margin = new System.Windows.Forms.Padding(0);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(293, 27);
            this.button4.TabIndex = 7;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatTitle2
            // 
            this.lblFormatTitle2.AutoSize = true;
            this.lblFormatTitle2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatTitle2.Location = new System.Drawing.Point(4, 61);
            this.lblFormatTitle2.Name = "lblFormatTitle2";
            this.lblFormatTitle2.Size = new System.Drawing.Size(227, 27);
            this.lblFormatTitle2.TabIndex = 4;
            this.lblFormatTitle2.Text = "标题（一）";
            this.lblFormatTitle2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button3
            // 
            this.button3.AutoSize = true;
            this.button3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button3.Location = new System.Drawing.Point(235, 61);
            this.button3.Margin = new System.Windows.Forms.Padding(0);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(293, 27);
            this.button3.TabIndex = 5;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatTitle1
            // 
            this.lblFormatTitle1.AutoSize = true;
            this.lblFormatTitle1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatTitle1.Location = new System.Drawing.Point(4, 33);
            this.lblFormatTitle1.Name = "lblFormatTitle1";
            this.lblFormatTitle1.Size = new System.Drawing.Size(227, 27);
            this.lblFormatTitle1.TabIndex = 2;
            this.lblFormatTitle1.Text = "标题一、";
            this.lblFormatTitle1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button2
            // 
            this.button2.AutoSize = true;
            this.button2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button2.Location = new System.Drawing.Point(235, 33);
            this.button2.Margin = new System.Windows.Forms.Padding(0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(293, 27);
            this.button2.TabIndex = 3;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.SetFont);
            // 
            // lblFormatTitle
            // 
            this.lblFormatTitle.AutoSize = true;
            this.lblFormatTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFormatTitle.Location = new System.Drawing.Point(4, 1);
            this.lblFormatTitle.Name = "lblFormatTitle";
            this.lblFormatTitle.Size = new System.Drawing.Size(227, 31);
            this.lblFormatTitle.TabIndex = 0;
            this.lblFormatTitle.Text = "公文标题";
            this.lblFormatTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblFormatTitle.UseMnemonic = false;
            // 
            // button1
            // 
            this.button1.AutoSize = true;
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(235, 1);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(293, 31);
            this.button1.TabIndex = 1;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.SetFont);
            // 
            // cobLetter
            // 
            this.cobLetter.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cobLetter.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cobLetter.BackColor = System.Drawing.SystemColors.Control;
            this.cobLetter.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cobLetter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cobLetter.FormattingEnabled = true;
            this.cobLetter.ItemHeight = 12;
            this.cobLetter.Location = new System.Drawing.Point(235, 285);
            this.cobLetter.Margin = new System.Windows.Forms.Padding(0);
            this.cobLetter.Name = "cobLetter";
            this.cobLetter.Size = new System.Drawing.Size(293, 20);
            this.cobLetter.Sorted = true;
            this.cobLetter.TabIndex = 21;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.logoPictureBox);
            this.panel1.Controls.Add(this.txtInfo);
            this.panel1.Controls.Add(this.lblText);
            this.panel1.Controls.Add(this.lnkUpdate);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(9, 9);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(532, 270);
            this.panel1.TabIndex = 37;
            // 
            // btnCustom
            // 
            this.btnCustom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCustom.Location = new System.Drawing.Point(343, 5);
            this.btnCustom.Name = "btnCustom";
            this.btnCustom.Size = new System.Drawing.Size(80, 23);
            this.btnCustom.TabIndex = 39;
            this.btnCustom.Text = ">>高级功能";
            this.btnCustom.UseVisualStyleBackColor = true;
            this.btnCustom.Click += new System.EventHandler(this.BtnCustom_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnSave);
            this.panel2.Controls.Add(this.lblTip);
            this.panel2.Controls.Add(this.btnCustom);
            this.panel2.Controls.Add(this.btnOK);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(9, 636);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(532, 34);
            this.panel2.TabIndex = 40;
            // 
            // btnSave
            // 
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(236, 5);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(80, 23);
            this.btnSave.TabIndex = 41;
            this.btnSave.Text = "保存设置";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Visible = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblTip
            // 
            this.lblTip.AutoSize = true;
            this.lblTip.ForeColor = System.Drawing.Color.Blue;
            this.lblTip.Location = new System.Drawing.Point(4, 10);
            this.lblTip.Name = "lblTip";
            this.lblTip.Size = new System.Drawing.Size(35, 12);
            this.lblTip.TabIndex = 40;
            this.lblTip.Text = "说明:";
            this.lblTip.Visible = false;
            this.lblTip.DoubleClick += new System.EventHandler(this.lblTip_DoubleClick);
            // 
            // AboutBox1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(550, 679);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(566, 356);
            this.Name = "AboutBox1";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "关于";
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.AboutBox1_HelpButtonClicked);
            this.Load += new System.EventHandler(this.AboutBox1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.PictureBox logoPictureBox;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.LinkLabel lnkUpdate;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label lblFormatSong;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Label lblFormatKai;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Label lblFormatHei;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Label lblFormatNormal;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Label lblFormatTitle4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label lblFormatTitle3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label lblFormatTitle2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label lblFormatTitle1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label lblFormatTitle;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnCustom;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblTip;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblPageNumber;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Label lblLetter;
        private System.Windows.Forms.ComboBox cobLetter;
    }
}
