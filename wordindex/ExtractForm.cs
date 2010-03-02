using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Word2003Tools4Dominique
{
    public partial class ExtractForm : Form
    {
        string _TextToParse;
        public ExtractForm()
        {
            InitializeComponent();
        }
        public ExtractForm(string Text)
        {
            InitializeComponent();
            _TextToParse = Text;
            labelMsg.Visible = false;
            /*
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 100;
             * */
        }

        private void ExtractForm_Load(object sender, EventArgs e)
        {
            if (!File.Exists(Util.GetPatternsFilePath())) Util.GenerateDefaultPatternsFile();
            LoadPatterns();
            comboBoxPattern.SelectedIndex = 0;
            labelComment.Text = "";
        }
        void LoadPatterns()
        {
            comboBoxPattern.Items.Clear();
            List<string> filterlist = Util.GetPatterns();
            foreach (string line in filterlist)
            {
                comboBoxPattern.Items.Add(line);
            }
        }

        private void linkLabelRegularExpression_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Util.BrowseUrl(Util.REGULAR_EXPRESSION_HELP_URL1);
        }

        private void buttonExtract_Click(object sender, EventArgs e)
        {
            labelMsg.Visible = true;
            labelMsg.Update();

            string pattern=textBoxPatternToExtract.Text;

            List<string> wordListResult=new List<string>();
            if (!String.IsNullOrEmpty(pattern))
            {
                bool ignoreCase = checkBoxIgnoreCase.Checked;
                Commands.ExtractPattern(_TextToParse, pattern, ignoreCase, ref wordListResult);
                string resultFilePath = Util.GetFinalResultFilePath();
                Util.WriteToFile(resultFilePath, wordListResult, Msg.MSG_NOTHING_FOUND_OR_NO_TEXT);
                Util.SortUniq(resultFilePath, Util.GetSortedUniqResultFilePath());
                Util.DisplayFile(Util.GetSortedUniqResultFilePath());
            }
            labelMsg.Visible = false;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBoxPattern_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text=comboBoxPattern.Text;
            textBoxPatternToExtract.Text = text;
            labelComment.Text = text;
            int idx = text.IndexOf(Util.COMMENT_MARKER);
            if (idx > 0)
            {
                labelComment.Text = text.Substring(0, idx);
                textBoxPatternToExtract.Text = text.Substring(idx + Util.COMMENT_MARKER.Length, text.Length - (idx + Util.COMMENT_MARKER.Length));
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Util.BrowseUrl(Util.REGULAR_EXPRESSION_HELP_URL2);
        }

        private void buttonPatterns_Click(object sender, EventArgs e)
        {
            Util.DisplayFile(Util.GetPatternsFilePath());
        }

    }
}
