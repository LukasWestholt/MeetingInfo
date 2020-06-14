namespace MeetingInfo
{
    public class SettingsWrapper
    {
        public SettingsWrapper()
        {
            Language = string.IsNullOrEmpty(Properties.Settings.Default.Language) ? System.Globalization.CultureInfo.InstalledUICulture.Name : Properties.Settings.Default.Language;
            AcceptButton = Properties.Settings.Default.AcceptButton;
            RibbonMaxWidth = CheckWidth(Properties.Settings.Default.RibbonMaxWidth == 0 ? 240 : Properties.Settings.Default.RibbonMaxWidth);
            // TODO get window width calculator 0-1024 chars
            // Explorer.Width
            // https://docs.microsoft.com/de-de/dotnet/api/microsoft.office.interop.outlook._explorer.width?view=outlook-pia#Microsoft_Office_Interop_Outlook__Explorer_Width
        }

        public string Language { get; private set; }
        public bool AcceptButton { get; private set; }

        public int RibbonMaxWidth { get; private set; }

        public bool SetLanguage(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            if (Language == value) return true;
            try
            {
                new System.Globalization.CultureInfo(value);
            }
            catch (System.Globalization.CultureNotFoundException e)
            {
                // do NOT translate this text
                if (System.Windows.Forms.MessageBox.Show(e.Message + "\n\r\n\rDo you want to open table of language tags in browser?",
                    "Language key not found", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Error) == System.Windows.Forms.DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("https://docs.microsoft.com/openspecs/windows_protocols/ms-lcid/a9eac961-e77d-41a6-90a5-ce1a8b0cdb9c");
                }
                return false;
            }
            Language = value;
            Globals.MeetingInfoMain.SetCultureInfo(Language);
            Properties.Settings.Default.Language = Language;
            SaveUserSettings();
            return true;
        }

        public bool SetAcceptButton(bool new_state)
        {
            if (AcceptButton == new_state) return true;
            AcceptButton = new_state;
            Properties.Settings.Default.AcceptButton = new_state;
            SaveUserSettings();
            return true;
        }

        public bool SetRibbonMaxWidth(int value)
        {
            if (RibbonMaxWidth == value) return true;
            value = CheckWidth(value);
            RibbonMaxWidth = value;
            Properties.Settings.Default.RibbonMaxWidth = value;
            SaveUserSettings();
            return true;
        }

        private int CheckWidth(int i)
        {
            if (i < 0) return 0;
            if (i > 1024) return 1024;
            return i;
        }

        public void SaveUserSettings()
        {
            Properties.Settings.Default.Save();
        }
    }
}
