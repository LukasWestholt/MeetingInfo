using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Globalization;

namespace MeetingInfo
{
    public partial class MeetingInfoMain
    {

        // <div>Icons erstellt von <a href="https://smashicons.com/" title="Smashicons">Smashicons</a> from <a href="https://www.flaticon.com/de/" title="Flaticon">www.flaticon.com</a></div>
        // <a target="_blank" href="https://icons8.de/icons/set/high-priority">Hohe Priorität icon</a> icon by <a target="_blank" href="https://icons8.de">Icons8</a>

        // TODO settings width, language, credits

        // TODO size large?

        // TODO meeting in mailinglist no ribbon problem

        // TODO BUILD https://docs.microsoft.com/de-de/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2019

        private readonly Ribbon _ribbon = new Ribbon();
        
        private Dictionary<int, ElementWrapper> labels = new Dictionary<int, ElementWrapper>();

        private Inspectors inspectors;
        private Explorers explorers;
        private readonly long last_exec = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();

        // https://github.com/LukasWestholt/MeetingInfo/tree/master/MeetingInfo
        public const string ADD_IN_NAME = "MeetingInfo";
        public const string VERSION = "1.0";

        public const string DEFAULT_TEXT_LABEL = "NULL";
        public const string DEFAULT_TEXT_SCREENTIP = "NULL";
        public const bool DEFAULT_STATE_VISIBLE = true;
        public const string LANG = "de-DE";
        private const int DEBUG = 1;
        private readonly string[] DEBUG_TEXT = {"ERROR", "INFO", "DEBUG"};


        private readonly System.Resources.ResourceManager resmgr = new System.Resources.ResourceManager("MeetingInfo.Properties.Resources", typeof(MeetingInfoMain).Assembly);
        private readonly CultureInfo ci = new CultureInfo(LANG);

        public Dictionary<Inspector, InspectorWrapper> InspectorWrappers { get; } = new Dictionary<Inspector, InspectorWrapper>();
        public Dictionary<Explorer, ExplorerWrapper> ExplorerWrappers { get; } = new Dictionary<Explorer, ExplorerWrapper>();

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _ribbon;
        }

        private string I18n(string s)
        {
            return resmgr.GetString(s, ci);
        }

        public void CheckObject(Object selObject)
        {
            if (debugCheck(2))
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(resmgr.GetStream("fgth_welcome", ci));
                player.Play();
            }

            // TODO merging AppointmentItem and MeetingItem??
            if (selObject == null)
            {
                DPrint("selObject " + I18n("IS_NULL"), 0);
            }
            else if (selObject is AppointmentItem) // Kalenderliste + Doppel-klick auf Kalendereinträge
            {   
                DPrint(I18n("RECOGNIZED") + " - AppointmentItem", 2);
                // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem#properties
                AppointmentItem apptItem = (selObject as AppointmentItem);
                if (apptItem != null && apptItem.EntryID != null)
                {
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.organizer
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.getorganizer
                    string Organizer_Name;
                    AddressEntry Organizer_AdressEntry = apptItem.GetOrganizer();
                    if (Organizer_AdressEntry != null && !String.IsNullOrEmpty(Organizer_AdressEntry.Name))
                    {
                        Organizer_Name = Organizer_AdressEntry.Name + " [" + I18n("ADDRESS_BOOK") + "]";
                        // Organizer_Name = Organizer_AdressEntry.GetContact().FullName + " [Adressbuch]";
                        // Organizer_AdressEntry.GetContact().Display();
                        //Organizer_AdressEntry.Details();
                        // TODO Anzeige von Namen + Mail-Adresse von Organizer
                    }
                    else
                    {
                        DPrint("GetOrganizer " + I18n("IS_EMPTY"), 1);
                        Organizer_Name = apptItem.Organizer ?? "[" + I18n("ORGANIZER") + " " + I18n("IS_EMPTY") + "]";
                    }
                    if (Organizer_AdressEntry.Name != apptItem.Organizer) DPrint(Organizer_AdressEntry.Name + " / " + apptItem.Organizer, 2);
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.subject
                    string Subject = apptItem.Subject ?? "[" + I18n("SUBJECT") + " " + I18n("IS_EMPTY") + "]";

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentitem.meetingstatus
                    // bool isMeeting = (MeetingStatusToInt(apptItem.MeetingStatus) != 0);

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem.recipients
                    string[] output = Recipients(apptItem.Recipients, (Organizer_AdressEntry != null ? Organizer_AdressEntry.ID : ""));
                    string RequiredAttendees = output[0]; // TODO is this without organizer?? testing with somebody
                    string OptionalAttendees = output[1];
                    
                    //max 1024 characters.
                    SetElement(new string[] {Subject, Organizer_Name, RequiredAttendees, OptionalAttendees }, 0);
                    SetElement(new string[] { I18n("SUBJECT"), I18n("ORGANIZER"), I18n("REQUIRED_ATTENDEES"), I18n("OPTIONAL_ATTENDEES") }, 1);
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem.importance
                    CheckImage((System.Drawing.Bitmap)resmgr.GetObject("important", ci), 0, MessgageImportanceToInt(apptItem.Importance) == 2);
                    
                }
            }
            else if (selObject is MeetingItem) // Meetings in Mailliste + Doppel-klick auf Mails
            {
                DPrint(I18n("RECOGNIZED") + " - MeetingItem", 2);
                // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem#properties
                MeetingItem meetItem = (selObject as MeetingItem);
                if (meetItem != null && meetItem.EntryID != null)
                {
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.sendername
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.senderemailaddress
                    string Organizer_Name;
                    if (meetItem.SenderName != meetItem.SenderEmailAddress)
                    {
                        Organizer_Name = meetItem.SenderName + "/" + meetItem.SenderEmailAddress;
                    }
                    else
                    {
                        Organizer_Name = meetItem.SenderName;
                    }
                    // TODO why this and not above

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.subject
                    string Subject = meetItem.Subject ?? "[" + I18n("SUBJECT") + " " + I18n("IS_EMPTY") + "]";

                    //https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.getassociatedappointment
                    DPrint(meetItem.GetAssociatedAppointment(false).RequiredAttendees, 2);
                    // TODO testing this and check results

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.recipients
                    string[] output = Recipients(meetItem.Recipients, "");
                    string RequiredAttendees = output[0];
                    string OptionalAttendees = output[1];

                    //max 1024 characters. TODO checking this out
                    SetElement(new string[] { Subject, Organizer_Name, RequiredAttendees, OptionalAttendees }, 0);
                    SetElement(new string[] { I18n("SUBJECT"), I18n("ORGANIZER"), I18n("REQUIRED_ATTENDEES"), I18n("OPTIONAL_ATTENDEES") }, 1);
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.meetingitem.importance
                    CheckImage((System.Drawing.Bitmap)resmgr.GetObject("important", ci), 0, MessgageImportanceToInt(meetItem.Importance) == 2);
                    

                    // TODO WORK ON THIS NEW FEATURE
                    /*if (!meetItem.Sent)
                    {
                        // https://docs.microsoft.com/de-de/office/vba/api/outlook.account
                        
                        Accounts accounts = Application.Session.Accounts;
                        foreach (Account account in accounts)
                        {
                            // get selected account
                        }
                        meetItem.SendUsingAccount("<selected account>");
                    }*/
                }
            }
            else
            {
                // TODO is this working and needed? first start?
                // TODO double exec??
                if (last_exec + 500 < new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds())
                {
                    SetElement(new string[] { null, null, null, null }, 0);
                    SetElement(new string[] { null, null, null, null }, 1);
                    CheckImage(null, 0, false);
                }
            }
        }

        private string[] Recipients(Recipients recipients, string organisator_ID)
        {
            string RequiredAttendees = "";
            string OptionalAttendees = "";
            const string STRING_SEPERATOR = "; ";
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.appointmentItem.recipients
            foreach (Recipient recipient in recipients)
            {
                // TODO Testing with somebody
                if (recipient.AddressEntry.ID != organisator_ID)
                {
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.resolve
                    bool resolved = recipient.Resolve();

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.addressentry.name
                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.meetingresponsestatus
                    // TODO Mail-Adresse https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.address

                    string text = (resolved ? recipient.AddressEntry.Name + " [" + I18n("ADDRESS_BOOK") +"]" : recipient.Name) +
                            (!String.IsNullOrEmpty(MeetingResponseToStr(recipient.MeetingResponseStatus)) ? " (" + MeetingResponseToStr(recipient.MeetingResponseStatus) + ")" : "");

                    // https://docs.microsoft.com/de-de/office/vba/api/outlook.recipient.type
                    if (recipient.Type == 1)
                    {
                        // olRequired
                        RequiredAttendees += text + STRING_SEPERATOR;
                    }
                    else if (recipient.Type == 2)
                    {
                        // olOptional
                        OptionalAttendees += text + STRING_SEPERATOR;
                    }
                    else
                    {
                        // olOrganizer or olResource
                        RequiredAttendees += text + STRING_SEPERATOR;
                        DPrint("olOrganizer or olResource detected", 1);
                    }
                }
                else
                {
                    DPrint("NOT WRITE CAUSE ORGA", 0);
                }
            }
            if (RequiredAttendees.Length > 1) RequiredAttendees = RequiredAttendees.Remove(RequiredAttendees.Length - STRING_SEPERATOR.Length);
            if (OptionalAttendees.Length > 1) OptionalAttendees = OptionalAttendees.Remove(OptionalAttendees.Length - STRING_SEPERATOR.Length);
            return new string[] { RequiredAttendees, OptionalAttendees };
        }

        private void SetElement(string text, int label_int, int type)
        {
            if (!String.IsNullOrEmpty(text))
            {
                if (type == 0)
                {
                    labels[label_int].Label = text;
                    labels[label_int].Visible = true;
                }
                if (type == 1) labels[label_int].Screentip = text;
            }
            else
            {
                if (type == 0)
                {
                    labels[label_int].Label = "";
                    labels[label_int].Visible = false;
                }
                if (type == 1) labels[label_int].Screentip = "";
            }
        }

        private void SetElement(string[] texts, int type)
        {
            foreach (KeyValuePair<int, ElementWrapper> entry in labels)
            {
                if (!String.IsNullOrEmpty(texts[entry.Key]))
                {
                    if (type == 0)
                    {
                        entry.Value.Label = texts[entry.Key];
                        entry.Value.Visible = true;
                    }
                    if (type == 1) entry.Value.Screentip = texts[entry.Key];
                }
                else
                {
                    if (type == 0)
                    {
                        entry.Value.Label = "";
                        entry.Value.Visible = false;
                    }
                    if (type == 1) entry.Value.Screentip = "";
                }
            }
        }

        private void SetElementImage(System.Drawing.Bitmap img, int label_int)
        {
            labels[label_int].Image = img;
        }

        private void CheckImage(System.Drawing.Bitmap img, int label_int, bool show_img)
        {
            if (show_img)
            {
                SetElementImage(img, label_int);
            }
            else
            {
                SetElementImage(null, label_int);
            }
        }

        private string MeetingResponseToStr(OlResponseStatus response)
        {
            // https://docs.microsoft.com/de-de/office/vba/api/outlook.olresponsestatus
            switch (response) {
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

        private void DPrint(string str, int debugMode)
        {
            if (debugCheck(debugMode)) System.Diagnostics.Debug.WriteLine("[" + ADD_IN_NAME + "] " + I18n(DEBUG_TEXT[debugMode]) + ": " + str);
        }

        private bool debugCheck(int debugMode)
        {
            // true when debugMode is lower/equal then DEBUG
            return (debugMode <= DEBUG);
        }

        private void MeetingInfoMain_Startup(object sender, System.EventArgs e)
        {
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
                { 3, _ribbon.Label4 }
            };

            DPrint("Assembly full name:\n   " + typeof(MeetingInfoMain).Assembly.FullName, 2);
            DPrint("Assembly qualified name:\n   " + typeof(MeetingInfoMain).AssemblyQualifiedName, 2);
        }

        private void MeetingInfoMain_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(MeetingInfoMain_Startup);
            this.Shutdown += new System.EventHandler(MeetingInfoMain_Shutdown);
        }

        #endregion
    }
}
