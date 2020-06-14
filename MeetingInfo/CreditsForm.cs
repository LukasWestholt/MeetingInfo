using System.Diagnostics;
using System.Windows.Forms;

namespace MeetingInfo
{
    public partial class CreditsForm : Form
    {
        public CreditsForm()
        {
            InitializeComponent();
        }

        private void CreditsViewer_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            Debug.WriteLine(e.Url.ToString());
            if (!e.Url.ToString().StartsWith("about:"))
            {
                e.Cancel = true;
                System.Diagnostics.Process.Start(e.Url.ToString());
            }
        }

        // https://stackoverflow.com/a/5312580/8980073
        private void CreditsViewer_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            CreditsViewer.Document.Click += new HtmlElementEventHandler(Document_Click);
        }

        void Document_Click(object sender, HtmlElementEventArgs e)
        {
            HtmlElement ele = CreditsViewer.Document.GetElementFromPoint(e.MousePosition);
            while (ele != null)
            {
                if (ele.TagName.ToLower() == "a")
                {
                    ele.SetAttribute("target", "_self");
                }
                ele = ele.Parent;
            }
        }
    }
}
