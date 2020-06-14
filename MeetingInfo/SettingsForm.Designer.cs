namespace MeetingInfo
{
    partial class SettingsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.LabelLanguage = new System.Windows.Forms.Label();
            this.TextBoxLanguage = new System.Windows.Forms.TextBox();
            this.ButtonSave = new System.Windows.Forms.Button();
            this.ButtonGetLanguage = new System.Windows.Forms.Button();
            this.CheckBoxMeetingAcceptButton = new System.Windows.Forms.CheckBox();
            this.LabelMeetingAcceptButton = new System.Windows.Forms.Label();
            this.LabelRibbonMaxWidth = new System.Windows.Forms.Label();
            this.NumericUpDownRibbonMaxWidth = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.NumericUpDownRibbonMaxWidth)).BeginInit();
            this.SuspendLayout();
            // 
            // LabelLanguage
            // 
            resources.ApplyResources(this.LabelLanguage, "LabelLanguage");
            this.LabelLanguage.Name = "LabelLanguage";
            // 
            // TextBoxLanguage
            // 
            resources.ApplyResources(this.TextBoxLanguage, "TextBoxLanguage");
            this.TextBoxLanguage.Name = "TextBoxLanguage";
            // 
            // ButtonSave
            // 
            resources.ApplyResources(this.ButtonSave, "ButtonSave");
            this.ButtonSave.Name = "ButtonSave";
            this.ButtonSave.UseVisualStyleBackColor = true;
            this.ButtonSave.Click += new System.EventHandler(this.ButtonSave_Click);
            // 
            // ButtonGetLanguage
            // 
            resources.ApplyResources(this.ButtonGetLanguage, "ButtonGetLanguage");
            this.ButtonGetLanguage.Name = "ButtonGetLanguage";
            this.ButtonGetLanguage.UseVisualStyleBackColor = true;
            this.ButtonGetLanguage.Click += new System.EventHandler(this.ButtonGetLanguage_Click);
            // 
            // CheckBoxMeetingAcceptButton
            // 
            resources.ApplyResources(this.CheckBoxMeetingAcceptButton, "CheckBoxMeetingAcceptButton");
            this.CheckBoxMeetingAcceptButton.Name = "CheckBoxMeetingAcceptButton";
            this.CheckBoxMeetingAcceptButton.UseVisualStyleBackColor = true;
            // 
            // LabelMeetingAcceptButton
            // 
            resources.ApplyResources(this.LabelMeetingAcceptButton, "LabelMeetingAcceptButton");
            this.LabelMeetingAcceptButton.Name = "LabelMeetingAcceptButton";
            // 
            // LabelRibbonMaxWidth
            // 
            resources.ApplyResources(this.LabelRibbonMaxWidth, "LabelRibbonMaxWidth");
            this.LabelRibbonMaxWidth.Name = "LabelRibbonMaxWidth";
            // 
            // NumericUpDownRibbonMaxWidth
            // 
            resources.ApplyResources(this.NumericUpDownRibbonMaxWidth, "NumericUpDownRibbonMaxWidth");
            this.NumericUpDownRibbonMaxWidth.Maximum = new decimal(new int[] {
            1024,
            0,
            0,
            0});
            this.NumericUpDownRibbonMaxWidth.Name = "NumericUpDownRibbonMaxWidth";
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.ButtonSave;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.NumericUpDownRibbonMaxWidth);
            this.Controls.Add(this.LabelRibbonMaxWidth);
            this.Controls.Add(this.LabelMeetingAcceptButton);
            this.Controls.Add(this.CheckBoxMeetingAcceptButton);
            this.Controls.Add(this.ButtonGetLanguage);
            this.Controls.Add(this.ButtonSave);
            this.Controls.Add(this.TextBoxLanguage);
            this.Controls.Add(this.LabelLanguage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "SettingsForm";
            ((System.ComponentModel.ISupportInitialize)(this.NumericUpDownRibbonMaxWidth)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LabelLanguage;
        private System.Windows.Forms.Button ButtonSave;
        public System.Windows.Forms.TextBox TextBoxLanguage;
        private System.Windows.Forms.Button ButtonGetLanguage;
        private System.Windows.Forms.Label LabelMeetingAcceptButton;
        public System.Windows.Forms.CheckBox CheckBoxMeetingAcceptButton;
        private System.Windows.Forms.Label LabelRibbonMaxWidth;
        public System.Windows.Forms.NumericUpDown NumericUpDownRibbonMaxWidth;
    }
}