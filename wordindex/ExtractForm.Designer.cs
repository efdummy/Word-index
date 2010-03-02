namespace Word2003Tools4Dominique
{
    partial class ExtractForm
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
            this.comboBoxPattern = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonExtract = new System.Windows.Forms.Button();
            this.linkLabelRegularExpression = new System.Windows.Forms.LinkLabel();
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelComment = new System.Windows.Forms.Label();
            this.labelMsg = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxIgnoreCase = new System.Windows.Forms.CheckBox();
            this.textBoxPatternToExtract = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonPatterns = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBoxPattern
            // 
            this.comboBoxPattern.FormattingEnabled = true;
            this.comboBoxPattern.Location = new System.Drawing.Point(141, 78);
            this.comboBoxPattern.Name = "comboBoxPattern";
            this.comboBoxPattern.Size = new System.Drawing.Size(502, 24);
            this.comboBoxPattern.TabIndex = 0;
            this.comboBoxPattern.SelectedIndexChanged += new System.EventHandler(this.comboBoxPattern_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Pattern to extract :";
            // 
            // buttonExtract
            // 
            this.buttonExtract.Location = new System.Drawing.Point(649, 44);
            this.buttonExtract.Name = "buttonExtract";
            this.buttonExtract.Size = new System.Drawing.Size(104, 23);
            this.buttonExtract.TabIndex = 2;
            this.buttonExtract.Text = "Extract...";
            this.buttonExtract.UseVisualStyleBackColor = true;
            this.buttonExtract.Click += new System.EventHandler(this.buttonExtract_Click);
            // 
            // linkLabelRegularExpression
            // 
            this.linkLabelRegularExpression.AutoSize = true;
            this.linkLabelRegularExpression.Location = new System.Drawing.Point(299, 230);
            this.linkLabelRegularExpression.Name = "linkLabelRegularExpression";
            this.linkLabelRegularExpression.Size = new System.Drawing.Size(54, 17);
            this.linkLabelRegularExpression.TabIndex = 5;
            this.linkLabelRegularExpression.TabStop = true;
            this.linkLabelRegularExpression.Text = "Google";
            this.linkLabelRegularExpression.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRegularExpression_LinkClicked);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(649, 227);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(104, 23);
            this.buttonClose.TabIndex = 6;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelComment
            // 
            this.labelComment.Location = new System.Drawing.Point(138, 105);
            this.labelComment.Name = "labelComment";
            this.labelComment.Size = new System.Drawing.Size(505, 116);
            this.labelComment.TabIndex = 7;
            this.labelComment.Text = "Comment";
            // 
            // labelMsg
            // 
            this.labelMsg.AutoSize = true;
            this.labelMsg.Location = new System.Drawing.Point(358, 21);
            this.labelMsg.Name = "labelMsg";
            this.labelMsg.Size = new System.Drawing.Size(72, 17);
            this.labelMsg.TabIndex = 8;
            this.labelMsg.Text = "Working...";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(360, 230);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(65, 17);
            this.linkLabel1.TabIndex = 9;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Microsoft";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 230);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(242, 17);
            this.label2.TabIndex = 10;
            this.label2.Text = "Help on regular expression patterns :";
            // 
            // checkBoxIgnoreCase
            // 
            this.checkBoxIgnoreCase.AutoSize = true;
            this.checkBoxIgnoreCase.Location = new System.Drawing.Point(649, 21);
            this.checkBoxIgnoreCase.Name = "checkBoxIgnoreCase";
            this.checkBoxIgnoreCase.Size = new System.Drawing.Size(104, 21);
            this.checkBoxIgnoreCase.TabIndex = 11;
            this.checkBoxIgnoreCase.Text = "Ignore case";
            this.checkBoxIgnoreCase.UseVisualStyleBackColor = true;
            // 
            // textBoxPatternToExtract
            // 
            this.textBoxPatternToExtract.Location = new System.Drawing.Point(141, 44);
            this.textBoxPatternToExtract.Name = "textBoxPatternToExtract";
            this.textBoxPatternToExtract.Size = new System.Drawing.Size(502, 22);
            this.textBoxPatternToExtract.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 81);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(115, 17);
            this.label3.TabIndex = 13;
            this.label3.Text = "Pattern sample : ";
            // 
            // buttonPatterns
            // 
            this.buttonPatterns.Location = new System.Drawing.Point(649, 78);
            this.buttonPatterns.Name = "buttonPatterns";
            this.buttonPatterns.Size = new System.Drawing.Size(104, 23);
            this.buttonPatterns.TabIndex = 14;
            this.buttonPatterns.Text = "Edit samples";
            this.buttonPatterns.UseVisualStyleBackColor = true;
            this.buttonPatterns.Click += new System.EventHandler(this.buttonPatterns_Click);
            // 
            // ExtractForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 275);
            this.Controls.Add(this.buttonPatterns);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxPatternToExtract);
            this.Controls.Add(this.checkBoxIgnoreCase);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.labelMsg);
            this.Controls.Add(this.labelComment);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.linkLabelRegularExpression);
            this.Controls.Add(this.buttonExtract);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxPattern);
            this.Name = "ExtractForm";
            this.Text = "Type or select the pattern to extract";
            this.Load += new System.EventHandler(this.ExtractForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxPattern;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonExtract;
        private System.Windows.Forms.LinkLabel linkLabelRegularExpression;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelComment;
        private System.Windows.Forms.Label labelMsg;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBoxIgnoreCase;
        private System.Windows.Forms.TextBox textBoxPatternToExtract;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonPatterns;
    }
}