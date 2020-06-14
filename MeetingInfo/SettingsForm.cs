using System;
using System.Windows.Forms;

namespace MeetingInfo
{
    public partial class SettingsForm : Form
    {
        private SettingsWrapper _settings;
        private System.Globalization.CultureInfo _cultureInfo;
        public SettingsForm(SettingsWrapper settings, System.Globalization.CultureInfo cultureInfo)
        {
            _cultureInfo = cultureInfo;
            _settings = settings;
            System.Threading.Thread.CurrentThread.CurrentUICulture = _cultureInfo;
            InitializeComponent();
        }

        private void ButtonSave_Click(object sender, EventArgs e)
        {
            _settings.SetAcceptButton(CheckBoxMeetingAcceptButton.Checked);
            _settings.SetRibbonMaxWidth(Decimal.ToInt32(NumericUpDownRibbonMaxWidth.Value));
            bool success = _settings.SetLanguage(TextBoxLanguage.Text);
            if (success) Close();
        }

        private void ButtonGetLanguage_Click(object sender, EventArgs e)
        {
            TextBoxLanguage.Text = System.Globalization.CultureInfo.InstalledUICulture.Name;
        }
    }
}
