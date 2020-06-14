namespace MeetingInfo
{
    partial class CreditsForm
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
            this.CreditsViewer = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // CreditsViewer
            // 
            this.CreditsViewer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreditsViewer.IsWebBrowserContextMenuEnabled = false;
            this.CreditsViewer.Location = new System.Drawing.Point(0, 0);
            this.CreditsViewer.MinimumSize = new System.Drawing.Size(20, 20);
            this.CreditsViewer.Name = "CreditsViewer";
            this.CreditsViewer.ScrollBarsEnabled = false;
            this.CreditsViewer.Size = new System.Drawing.Size(334, 211);
            this.CreditsViewer.TabIndex = 0;
            this.CreditsViewer.WebBrowserShortcutsEnabled = false;
            this.CreditsViewer.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.CreditsViewer_DocumentCompleted);
            this.CreditsViewer.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.CreditsViewer_Navigating);
            // 
            // CreditsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(334, 211);
            this.Controls.Add(this.CreditsViewer);
            this.Name = "CreditsForm";
            this.ShowIcon = false;
            this.Text = "Credits";
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.WebBrowser CreditsViewer;
    }
}