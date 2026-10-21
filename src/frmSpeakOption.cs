using System;
using System.Drawing;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Diagnostics;

namespace 公文排版
{
    public partial class frmSpeakOption : Form
    {
        private SpeechSynthesizer synth;
        private Microsoft.Office.Interop.Word.Application WordApp;
        private Point mPoint;

        public frmSpeakOption()
        {
            InitializeComponent();
        }

        private void FrmSpeakOption_Load(object sender, EventArgs e)
        {
            this.Location = MousePosition;
            WordApp = Globals.ThisAddIn.Application;
        }

        private void BtnPlay_Click(object sender, EventArgs e)
        {
            if (synth == null || WordApp == null) return;

            synth.SpeakAsyncCancelAll();
            trackBar1.Enabled = false;
            cobVoiceName.Enabled = false;

            try
            {
                if (WordApp.Selection.Text.Length > 1)
                {
                    synth.SpeakAsync(WordApp.Selection.Text);
                }
                else
                {
                    WordApp.Selection.MoveEnd(Microsoft.Office.Interop.Word.WdUnits.wdParagraph);
                    synth.SpeakAsync(WordApp.Selection.Text);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("BtnPlay: " + ex.Message);
                trackBar1.Enabled = true;
                cobVoiceName.Enabled = true;
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            if (synth == null) return;
            synth.SpeakAsyncCancelAll();
            trackBar1.Enabled = true;
            cobVoiceName.Enabled = true;
        }

        private void TrackBar1_ValueChanged(object sender, EventArgs e)
        {
            if (synth != null) synth.Rate = trackBar1.Value;
        }

        private void CobVoiceName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (synth == null) return;
            if (cobVoiceName.SelectedItem != null)
                synth.SelectVoice(cobVoiceName.SelectedItem.ToString());
        }

        private void FrmSpeakOption_Shown(object sender, EventArgs e)
        {
            try
            {
                synth = new SpeechSynthesizer();
                synth.SpeakCompleted += Synth_SpeakCompleted;
                cobVoiceName.Items.Clear();

                foreach (InstalledVoice voice in synth.GetInstalledVoices(new System.Globalization.CultureInfo("zh-CN")))
                {
                    cobVoiceName.Items.Add(voice.VoiceInfo.Name);
                }
                if (cobVoiceName.Items.Count > 0)
                    cobVoiceName.SelectedIndex = 0;
                trackBar1.Value = synth.Rate;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("FrmSpeakOption_Shown: " + ex.Message);
                if (Globals.Ribbons.Ribbon1 != null)
                    Globals.Ribbons.Ribbon1.btnSpeak.Enabled = false;
                DisposeSynth();
                this.Close();
            }
        }

        private void Synth_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            if (this.IsDisposed) return;
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() =>
                {
                    trackBar1.Enabled = true;
                    cobVoiceName.Enabled = true;
                }));
            }
            else
            {
                trackBar1.Enabled = true;
                cobVoiceName.Enabled = true;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DisposeSynth();
            this.Close();
        }

        private void DisposeSynth()
        {
            if (synth == null) return;
            try
            {
                synth.SpeakAsyncCancelAll();
                synth.SpeakCompleted -= Synth_SpeakCompleted;
                synth.Dispose();
            }
            catch (Exception ex) { Debug.WriteLine(ex.Message); }
            finally { synth = null; }
        }

        private void FrmSpeakOption_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisposeSynth();
        }

        private void FrmSpeakOption_MouseDown(object sender, MouseEventArgs e)
        {
            mPoint = new Point(e.X, e.Y);
        }

        private void FrmSpeakOption_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Location = new Point(this.Location.X + e.X - mPoint.X,
                                          this.Location.Y + e.Y - mPoint.Y);
            }
        }
    }
}