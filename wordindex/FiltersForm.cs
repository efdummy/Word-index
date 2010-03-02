using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace Word2003Tools4Dominique
{
    public partial class FiltersForm : Form
    {
        bool DIRTY;
        public int MinValue;
        public FiltersForm(int min)
        {
            InitializeComponent();
            DIRTY = false;
            listBoxPatterns.Sorted = true;
            MinValue = min;
            EnableDisableUI();
        }
        
        const string REGEXP_HELP_TITLE = "Help on regular expressions";
        const string REGEXP_HELP = "je             (je)\naccueill* (all words starting with accueill) \nl[ea]        (le and la)\n?-?          (all three letters words with a midlle separator)\n[IVXLC]+ (Roman digits)";
        const string REGEXP_TOOTLIP = "Type here your regular expression to filter words";
        const string MSG_INCORRECT_REGEXP = "The regular expression is'nt correct. Check the syntax.";

        private void FiltersForm_Load(object sender, EventArgs e)
        {
            ToolTip help = new ToolTip();
            help.SetToolTip(textBoxPattern, REGEXP_TOOTLIP);
            LoadFilters();
            for (int i = 1; i < 10; i++)
                comboBoxSize.Items.Add(i.ToString());
            comboBoxSize.SelectedIndex = MinValue - 1;
        }
        void LoadFilters()
        {
            listBoxPatterns.Items.Clear();
            List<string> filterlist = Util.GetFilters();
            foreach (string line in filterlist)
            {
                listBoxPatterns.Items.Add(line);
            }
        }

        private void linkLabelRegularExpression_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Util.BrowseUrl(Util.REGULAR_EXPRESSION_HELP_URL1);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            MessageBox.Show(REGEXP_HELP, REGEXP_HELP_TITLE);
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            string FilePath = Util.GetFiltersFilePath();
            StreamWriter SW;
            SW = File.CreateText(FilePath);
            foreach(string line in listBoxPatterns.Items)
                SW.WriteLine(line);
            SW.Close();
        }
        private void EnableDisableUI()
        {
            switch(textBoxPattern.Text.Length)
            {
                case 0:
                    buttonAdd.Enabled = false;
                    break;
                default:
                    buttonAdd.Enabled = true;
                    buttonDelete.Enabled = true;
                    break;
            }

            if (listBoxPatterns.SelectedItem != null)
                buttonDelete.Enabled = true;
            else
                buttonDelete.Enabled = false;

            switch (DIRTY)
            {
                case true:
                    buttonSave.Enabled = true;
                    break;
                default:
                    buttonSave.Enabled = false;
                    break;
            }
        }

        private void textBoxPattern_TextChanged(object sender, EventArgs e)
        {
            EnableDisableUI();
        }

        private void listBoxPatterns_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxPattern.Text = listBoxPatterns.Text;
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            DIRTY = true;

            try
            {
                Regex regExp = new Regex(textBoxPattern.Text, RegexOptions.Compiled);
                listBoxPatterns.Items.Add(textBoxPattern.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(MSG_INCORRECT_REGEXP+" : "+ex.Message);
            }

            EnableDisableUI();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            DIRTY = true;
            if (listBoxPatterns.SelectedItem != null) listBoxPatterns.Items.Remove(listBoxPatterns.SelectedItem);
            EnableDisableUI();
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            Util.GenerateDefaultFiltersFile();
            LoadFilters();
            textBoxPattern.Text = "";
            comboBoxSize.SelectedIndex = MinValue - 1;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Util.BrowseUrl(Util.REGULAR_EXPRESSION_HELP_URL2);
        }

        private void comboBoxSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            MinValue = int.Parse(comboBoxSize.Text);
        }
    }
}
