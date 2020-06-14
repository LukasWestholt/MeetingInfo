using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Globalization;
using System.Linq;

namespace MeetingInfo
{
    public partial class MeetingInfoMain
    {
        // TODO meeting in mailinglist no ribbon problem
        // TODO BUILD https://docs.microsoft.com/de-de/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2019
        // TODO Im Kalender von Jahr zu Jahr "springen"

        private readonly Ribbon _ribbon = new Ribbon();

        private Dictionary<int, ElementWrapper> labels = new Dictionary<int, ElementWrapper>();

        // <-- must be document global (reason: garbage collector)
        private Inspectors inspectors;
        private Explorers explorers;
        private CreditsForm creditsForm;
        private SettingsForm settingsForm;
        // -->

        // https://github.com/LukasWestholt/MeetingInfo/tree/master/MeetingInfo
        public const string ADD_IN_NAME = "MeetingInfo";
        public const string VERSION = "1.0";

        public const string DEFAULT_TEXT_LABEL = "NULL";
        public const string DEFAULT_TEXT_SCREENTIP = "NULL";
        public const string DEFAULT_TEXT_SUPERTEXT = "NULL";
        public const bool DEFAULT_STATE_VISIBLE = true;
        private const int DEBUG = 2;
        private readonly string[] DEBUG_TEXT = { "ERROR", "INFO", "DEBUG", "DEBUG+" };

        private const string STRING_SEPERATOR = "; ";
        private const string STRING_EXTENDER = "…"; // three dots away

        private readonly SettingsWrapper setting = new SettingsWrapper();
        private readonly System.Resources.ResourceManager resmgr = new System.Resources.ResourceManager("MeetingInfo.Properties.Resources", typeof(MeetingInfoMain).Assembly);
        private CultureInfo ci;

        public Dictionary<Inspector, InspectorWrapper> InspectorWrappers { get; } = new Dictionary<Inspector, InspectorWrapper>();
        public Dictionary<Explorer, ExplorerWrapper> ExplorerWrappers { get; } = new Dictionary<Explorer, ExplorerWrapper>();

        public void SetCultureInfo(string new_ci)
        {
            ci = new CultureInfo(new_ci);
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _ribbon;
        }

        public string I18n(string s)
        {
            string str = resmgr.GetString(s, ci);
            if (str == null) return s.ToLower().Replace("_", " ");
            return str;
        }

        public bool Event(Object selObject)
        {
            bool result = CheckObject(selObject);
            if (!result)
            {
                EverythingOnNull();
            }
            return result;
        }

        public bool CheckObject(Object selObject)
        {
            if (DebugCheck(3))
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(resmgr.GetStream("fgth_welcome", ci));
                player.Play();
            }

            // selObject=AppointmentItem => Kalenderliste + Doppel-klick auf Kalendereinträge
            // selObject=MeetingItem => Meetings in Mailliste + Doppel-klick auf Mails

            /*
            * The AppointmentItem object represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.
            * The AppointmentItem object includes methods that perform actions such as responding to or forwarding meeting requests, 
            * and properties that specify meeting details such as the location and time.
            * */

            if (selObject is MeetingItem)
            {
                DPrint(I18n("RECOGNIZED") + " - MeetingItem", 2);
                selObject = (selObject as MeetingItem).GetAssociatedAppointment(false);
            }

            if (!(selObject is AppointmentItem) || (selObject as AppointmentItem) == null || (selObject as AppointmentItem).EntryID == null) {
                if (selObject == null) DPrint("selObject " + I18n("IS_NULL"), 0);
                return false;
            }
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem#properties
            AppointmentItem apptItem = (selObject as AppointmentItem);
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.meetingstatus
            bool isMeeting = (MeetingStatusToInt(apptItem.MeetingStatus) != 0);
            DPrint(I18n("RECOGNIZED") + " - AppointmentItem: isMeeting=" + isMeeting.ToString(), 2);

            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.organizer
            string Organizer_Name;
            string Organizer_Full = null;
            if (apptItem.Organizer == null && MeetingStatusToInt(apptItem.MeetingStatus) == 1)  // my own non-sended meeting, in this case GetOrganizer() is obsolete.
            {
                Organizer_Name = apptItem.SendUsingAccount.UserName;
                DPrint("apptItem.SendUsingAccount.DisplayName=\"" + apptItem.SendUsingAccount.DisplayName + "\"", 2);
            }
            else
            {
                // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.getorganizer
                string[] result = AdressEntryNameExtract(apptItem.GetOrganizer());
                if (result != null)
                {
                    Organizer_Name = result[0];
                    Organizer_Full = result[1];
                }
                else
                {
                    Organizer_Name = apptItem.Organizer;
                }
            }
            DPrint("apptItem.MeetingStatus=\"" + apptItem.MeetingStatus + "\"", 2);
            if (apptItem.GetOrganizer().Name != apptItem.Organizer) DPrint("DIFF: apptItem.GetOrganizer().Name=\"" + apptItem.GetOrganizer().Name + "\" / apptItem.Organizer=\"" + apptItem.Organizer + "\"", 0);
            Organizer_Name = Organizer_Name ?? "[" + I18n("ORGANIZER") + " " + I18n("IS_EMPTY") + "]";
            Organizer_Full = Organizer_Full ?? Organizer_Name;

            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.subject
            string Subject = apptItem.Subject;

            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem.recipients
            string[] output = Recipients(apptItem.Recipients, (apptItem.GetOrganizer() != null ? apptItem.GetOrganizer().ID : ""));
            //string[] output = Recipients(apptItem.Recipients, "");
            // TODO testing is this without organizer?? testing with somebody
            string RequiredAttendees_Names = output[0];
            string OptionalAttendees_Names = output[1];
            string RequiredAttendees_Full = output[2];
            string OptionalAttendees_Full = output[3];
            bool Acceptable = MeetingStatusToInt(apptItem.MeetingStatus) == 3;
            int Importance = MessgageImportanceToInt(apptItem.Importance);
            Subject = Subject ?? "[" + I18n("SUBJECT") + " " + I18n("IS_EMPTY") + "]";

            SetElements(new string[] { Subject, Organizer_Name, RequiredAttendees_Names, OptionalAttendees_Names }, 0);
            SetElements(new string[] { I18n("SUBJECT"), I18n("ORGANIZER"), I18n("REQUIRED_ATTENDEES"), I18n("OPTIONAL_ATTENDEES") }, 1);
            SetElements(new string[] { Subject, Organizer_Full, RequiredAttendees_Full, OptionalAttendees_Full }, 2);
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.importance
            CheckImage((System.Drawing.Bitmap)resmgr.GetObject("important", ci), 0, Importance == 2);
            SetElement(Acceptable && setting.AcceptButton ? I18n("ACCEPT") : null, 4, 0);
            SetElement(Acceptable && setting.AcceptButton ? I18n("ACCEPT") : null, 4, 1);
            SetElement(Acceptable && setting.AcceptButton ? I18n("ACCEPT") : null, 4, 2);
            SetElement(Acceptable && setting.AcceptButton ? apptItem : null, 4);
            return true;
        }

        private void EverythingOnNull()
        {
            if (_ribbon.IsLoaded()) // don't run this if the ribbon is not loaded yet, only after OnRibbonLoaded.
            {
                SetElements(new string[] { null, null, null, null }, 0);
                SetElements(new string[] { null, null, null, null }, 1);
                SetElements(new string[] { null, null, null, null }, 2);
                CheckImage(null, 0, false);
                SetElement(null, 4, 0);
                SetElement(null, 4, 1);
                SetElement(null, 4, 2);
                SetElement(null, 4);
            }
            else
            {
                DPrint("ribbon is not loaded yet", 2);
            }
        }

        private string[] AdressEntryNameExtract(AddressEntry addressEntry)
        {
            if (addressEntry == null) return null;
            
            string Name;
            ContactItem contact = GetContact(addressEntry);

            if (contact != null && !String.IsNullOrEmpty(contact.FullName))
            {
                Name = contact.FullName;
            } else
            {
                Name = addressEntry.Name;
            }
            if (String.IsNullOrEmpty(Name)) return null;
            return new[] {Name, Name + (!String.IsNullOrEmpty(addressEntry.Address) && !Name.ToLower().Contains(addressEntry.Address.ToLower()) ? " (" + addressEntry.Address + ")" : "")};
        }

        private ContactItem GetContact(AddressEntry addressEntry)
        {
            try
            {
                addressEntry.GetContact();
            }
            catch (System.Exception e)
            {
                DPrint(I18n("OUTDATED") + " AddressEntry info " + e.Message, 0);
                return null;
            }
            return addressEntry.GetContact();
        }

        private string[] Recipients(Recipients recipients, string organisator_ID)
        {
            // RequiredAttendees_Names, OptionalAttendees_Names, RequiredAttendees_Full, OptionalAttendees_Full
            string[] output = new string[] { "", "", "", "" };
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem.recipients
            foreach (Recipient recipient in recipients)
            {
                AddressEntry adressEntry = recipient.AddressEntry;
                // TODO Testing with somebody
                //if(recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5FFD0003"))
                //https://stackoverflow.com/questions/49927811/outlook-appointment-meeting-organizer-information-inconsistent
                if (adressEntry.ID != organisator_ID)
                {
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.resolve
                    bool resolved = recipient.Resolve();

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.addressentry.name
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.meetingresponsestatus
                    // TODO Mail-Adresse https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.address

                    string[] result = AdressEntryNameExtract(adressEntry);
                    string NamesText = (resolved & !String.IsNullOrEmpty(result[0]) ? result[0] : recipient.Name) + 
                        (!String.IsNullOrEmpty(MeetingResponseToStr(recipient.MeetingResponseStatus)) ? " (" + MeetingResponseToStr(recipient.MeetingResponseStatus) + ")" : "");
                    string FullText = (resolved & !String.IsNullOrEmpty(result[1]) ? result[1] : recipient.Name + " [" + I18n("NOT_FOUND_IN") + " " + I18n("ADDRESS_BOOK") + "]") +
                            (!String.IsNullOrEmpty(MeetingResponseToStr(recipient.MeetingResponseStatus)) ? " (" + MeetingResponseToStr(recipient.MeetingResponseStatus) + ")" : "");

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.type
                    int x;
                    if (recipient.Type != 2)
                    {
                        if (recipient.Type != 1) DPrint("olOrganizer or olResource detected", 1);   // olOrganizer or olResource
                        x = 0;  // olRequired
                    }
                    else  // olOptional
                    {
                        x = 1;
                    }
                    output[x] += NamesText + STRING_SEPERATOR;
                    output[x+2] += FullText + STRING_SEPERATOR;
                }
                else
                {
                    DPrint("The organisator was found in the recipients", 2);
                }
            }
            for (int i = 0; i < output.Length; i++)
            {
                if (output[i].Length > 1) output[i] = output[i].Remove(output[i].Length - STRING_SEPERATOR.Length);
            }
            return output;
        }

        private void SetElement(string text, int label_int, int type)
        {
            if (!String.IsNullOrEmpty(text))
            {
                if (text.Length > 1024) text = text.Substring(0, 1024-STRING_EXTENDER.Length) + STRING_EXTENDER;
                // max 1024 characters on Label and Screentip. https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/d104fcb2-6177-4eb9-a400-0a5f8ddcd539
                if (type == 0)
                {
                    if (text.Length > setting.RibbonMaxWidth) text = text.Substring(0, setting.RibbonMaxWidth - STRING_EXTENDER.Length) + STRING_EXTENDER;
                    labels[label_int].Label = text;
                    labels[label_int].Visible = true;
                }
                if (type == 1) labels[label_int].Screentip = text;
                if (type == 2) labels[label_int].Supertip = text;
            }
            else
            {
                if (type == 0)
                {
                    labels[label_int].Label = "";
                    labels[label_int].Visible = false;
                }
                if (type == 1) labels[label_int].Screentip = "";
                if (type == 2) labels[label_int].Supertip = "";
            }
        }

        private void SetElements(string[] texts, int type)
        {
            foreach (KeyValuePair<int, ElementWrapper> entry in labels)
            {
                if (texts.Length < entry.Key+1) break;  // non-entries in texts not edit
                if (!String.IsNullOrEmpty(texts[entry.Key]))
                {
                    string text = texts[entry.Key];
                    if (text.Length > 1024) text = text.Substring(0, 1024 - STRING_EXTENDER.Length) + STRING_EXTENDER;
                    // max 1024 characters on Label and Screentip. https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/d104fcb2-6177-4eb9-a400-0a5f8ddcd539
                    if (type == 0)
                    {
                        if (text.Length > setting.RibbonMaxWidth) text = text.Substring(0, setting.RibbonMaxWidth - STRING_EXTENDER.Length) + STRING_EXTENDER;
                        entry.Value.Label = text;
                        entry.Value.Visible = true;
                    }
                    if (type == 1) entry.Value.Screentip = text;
                    if (type == 2) entry.Value.Supertip = text;
                }
                else
                {
                    if (type == 0)
                    {
                        entry.Value.Label = "";
                        entry.Value.Visible = false;
                    }
                    if (type == 1) entry.Value.Screentip = "";
                    if (type == 2) entry.Value.Supertip = "";
                }
            }
        }

        private void SetElement(AppointmentItem apptItem, int label_int)
        {
            labels[label_int].AppointmentItem = apptItem;
        }

        private void CheckImage(System.Drawing.Bitmap img, int label_int, bool show_img)
        {
            if (show_img)
            {
                labels[label_int].Image = img;
            }
            else
            {
                labels[label_int].Image = null;
            }
        }

        public void OnAction(IRibbonControl ribbonUI)
        {
            if (ribbonUI.Id.ToLower() == "buttoncredits")
            {
                creditsForm = new CreditsForm();
                creditsForm.CreditsViewer.DocumentText = (string)resmgr.GetObject("Credits", ci);
                creditsForm.Show();
            } else if (ribbonUI.Id.ToLower() == "buttonsettings")
            {
                settingsForm = new SettingsForm(setting, ci);
                settingsForm.TextBoxLanguage.Text = setting.Language;
                settingsForm.CheckBoxMeetingAcceptButton.Checked = setting.AcceptButton;
                settingsForm.NumericUpDownRibbonMaxWidth.Value = setting.RibbonMaxWidth;
                settingsForm.Show();
            } else if (new []{"label11", "label22", "label33", "label44" }.Contains(ribbonUI.Id.ToLower()))
            {
                // nothing
            } else if (new[] { "directaccept1", "directaccept2", "directaccept3", "directaccept4" }.Contains(ribbonUI.Id.ToLower()))
            {
                // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.respond
                // https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-automatically-accept-a-meeting-request
                MeetingItem meetItem = labels[4].AppointmentItem.Respond(OlMeetingResponse.olMeetingAccepted, true);
                DPrint(meetItem, 2);
                meetItem.Send();
            }
        }
        private string MeetingResponseToStr(OlResponseStatus response)
        {
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.olresponsestatus
            switch (response)
            {
                case OlResponseStatus.olResponseNone:
                    return "";
                case OlResponseStatus.olResponseOrganized:
                    return "";
                case OlResponseStatus.olResponseTentative:
                    return I18n("RESPONSE") + ": " + I18n("TENTATIVELY") + I18n("ACCEPTED");
                case OlResponseStatus.olResponseAccepted:
                    return I18n("RESPONSE") + ": " + I18n("ACCEPTED");
                case OlResponseStatus.olResponseDeclined:
                    return I18n("RESPONSE") + ": " + I18n("DECLINED");
                case OlResponseStatus.olResponseNotResponded:
                    return I18n("RESPONSE") + ": ?";
            }
            DPrint("no recognized OlResponseStatus", 0);
            return I18n("RESPONSE") + ": ?";
        }

        private int MeetingStatusToInt(OlMeetingStatus status)
        {
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.olmeetingstatus
            switch (status)
            {
                case OlMeetingStatus.olMeeting:
                    return 1;
                case OlMeetingStatus.olMeetingCanceled:
                    return 5;
                case OlMeetingStatus.olMeetingReceived:
                    return 3;
                case OlMeetingStatus.olMeetingReceivedAndCanceled:
                    return 7;
                case OlMeetingStatus.olNonMeeting:
                    return 0;
            }
            DPrint("no recognized OlMeetingStatus", 0);
            return 0;
        }

        private int MessgageImportanceToInt(OlImportance importance)
        {
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.olresponsestatus
            switch (importance)
            {
                case OlImportance.olImportanceHigh:
                    return 2;
                case OlImportance.olImportanceLow:
                    return 0;
                case OlImportance.olImportanceNormal:
                    return 1;
            }
            DPrint("no recognized OlImportance", 0);
            return 1;
        }

        private void Inspectors_NewInspector(Inspector ins)
        {
            InspectorWrappers.Add(ins, new InspectorWrapper(ins));
        }

        private void Explorers_NewExplorer(Explorer exp)
        {
            ExplorerWrappers.Add(exp, new ExplorerWrapper(exp));
        }

        private void DPrint(object obj, int debugMode = 1)
        {
            String str;
            if (obj == null)
            {
                str = "NULL";
            } else
            {
                str = obj.ToString();
            }
            if (DebugCheck(debugMode)) System.Diagnostics.Debug.WriteLine("[" + ADD_IN_NAME + "] " + I18n(DEBUG_TEXT[debugMode]) + ": " + str);
        }

        private bool DebugCheck(int debugMode)
        {
            // true when debugMode is lower/equal then DEBUG
            return (debugMode <= DEBUG);
        }

        private void MeetingInfoMain_Startup(object sender, System.EventArgs e)
        {
            ci = new CultureInfo(setting.Language);

            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            Explorers_NewExplorer(this.Application.ActiveExplorer());
            explorers = this.Application.Explorers;
            explorers.NewExplorer +=
                new ExplorersEvents_NewExplorerEventHandler(Explorers_NewExplorer);
            
            labels = new Dictionary<int, ElementWrapper>()
            {
                { 0, _ribbon.Label1 },
                { 1, _ribbon.Label2 },
                { 2, _ribbon.Label3 },
                { 3, _ribbon.Label4 },
                { 4, _ribbon.DirectAccept },
            };

            DPrint("Assembly full name:\n   " + typeof(MeetingInfoMain).Assembly.FullName, 2);
            DPrint("Assembly qualified name:\n   " + typeof(MeetingInfoMain).AssemblyQualifiedName, 2);
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(MeetingInfoMain_Startup);
        }

        #endregion
    }
}
