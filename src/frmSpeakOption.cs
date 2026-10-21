using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.AudioFormat;
using System.Diagnostics;

namespace 公文排版
{
    public partial class frmSpeakOption : Form
    {
        SpeechSynthesizer synth;
        Microsoft.Office.Interop.Word.Application WordApp;
        private Point mPoint;
        public frmSpeakOption()
        {
            InitializeComponent();
            
        }

        private void FrmSpeakOption_Leave(object sender, EventArgs e)
        {

        }


        private void FrmSpeakOption_Load(object sender, EventArgs e)
        {
            this.Location = MousePosition;
            WordApp = Globals.ThisAddIn.Application;

        }




        private void BtnPlay_Click(object sender, EventArgs e)
        {
            synth.SpeakAsyncCancelAll();
            trackBar1.Enabled = false;
            cobVoiceName.Enabled = false;

            if (WordApp.Selection.Text.Length > 1)
            {
                synth.SpeakAsync(WordApp.Selection.Text);
            }
            else
            {
                //WordApp.Selection.WholeStory();
                //WordApp.Selection.ExtendMode = true;
                //WordApp.Selection.Extend();//单词
                //WordApp.Selection.Extend();//句子
                //WordApp.Selection.Extend();//段落
                WordApp.Selection.MoveEnd(Microsoft.Office.Interop.Word.WdUnits.wdParagraph);
                //WordApp.Selection.ExtendMode = false;
                synth.SpeakAsync(WordApp.Selection.Text);
            }



        }



        private void BtnStop_Click(object sender, EventArgs e)
        {
            synth.SpeakAsyncCancelAll();
            trackBar1.Enabled = true;
            cobVoiceName.Enabled = true;
        }

        private void TrackBar1_ValueChanged(object sender, EventArgs e)
        {
            synth.Rate = trackBar1.Value;
        }

        private void CobVoiceName_SelectedIndexChanged(object sender, EventArgs e)
        {
            synth.SelectVoice(cobVoiceName.SelectedItem.ToString());
        }

        private void FrmSpeakOption_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void FrmSpeakOption_Shown(object sender, EventArgs e)
        {
            try
            {
                //throw new Exception();

                synth = new SpeechSynthesizer();
                synth.SpeakCompleted += Synth_SpeakCompleted;
                cobVoiceName.Items.Clear();


                Debug.WriteLine(synth.GetInstalledVoices().Count);

                foreach (InstalledVoice voice in synth.GetInstalledVoices(new System.Globalization.CultureInfo("zh-CN")))
                {
                    VoiceInfo info = voice.VoiceInfo;

                    cobVoiceName.Items.Add(info.Name);

                }
                if (cobVoiceName.Items.Count > 0)
                {
                    cobVoiceName.SelectedIndex = 0;
                }
                trackBar1.Value = synth.Rate;

            }
            catch
            {
                Globals.Ribbons.Ribbon1.btnSpeak.Enabled = false;
                this.Dispose();
            }

        }

        private void Synth_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            trackBar1.Enabled = true;
            cobVoiceName.Enabled = true;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            synth.Dispose();
            this.Close();
        }

        private void FrmSpeakOption_MouseDown(object sender, MouseEventArgs e)
        {
            mPoint = new Point(e.X, e.Y);
        }

        private void FrmSpeakOption_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Location = new Point(this.Location.X + e.X - mPoint.X, this.Location.Y + e.Y - mPoint.Y);
            }
        }
    }
    
}
