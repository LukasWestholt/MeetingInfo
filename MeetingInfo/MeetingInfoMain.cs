using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace MeetingInfo
{
    public partial class MeetingInfoMain
    {

        private readonly Ribbon _ribbon = new Ribbon();

        private Inspectors inspectors;
        private Explorers explorers;
        private long last_exec = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();

        public const string ADD_IN_NAME = "MeetingInfo";
        public const string VERSION = "1.0";

        public Dictionary<Inspector, InspectorWrapper> InspectorWrappers { get; } = new Dictionary<Inspector, InspectorWrapper>();
        public Dictionary<Explorer, ExplorerWrapper> ExplorerWrappers { get; } = new Dictionary<Explorer, ExplorerWrapper>();

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _ribbon;
        }

        private void Inspectors_NewInspector(Inspector ins)
        {
            InspectorWrappers.Add(ins, new InspectorWrapper(ins));
        }

        private void Explorers_NewExplorer(Explorer exp)
        {
            ExplorerWrappers.Add(exp, new ExplorerWrapper(exp));
        }

        public void CheckObject(Object selObject)
        {
            //System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"fgth_welcome.wav");
            //player.Play();

            if (selObject == null)
            {
                System.Diagnostics.Debug.WriteLine("[" + ADD_IN_NAME + "] ERROR: selObject war null");
                return;
            }

            if (selObject is AppointmentItem) // Kalenderliste + Doppel-klick auf Kalendereinträge
            {
                System.Diagnostics.Debug.WriteLine("erkannt - AppointmentItem");
                AppointmentItem apptItem = (selObject as AppointmentItem);
                if (apptItem != null && apptItem.EntryID != null)
                {
                    String OptionalAttendees = apptItem.OptionalAttendees;
                    String RequiredAttendees = apptItem.RequiredAttendees;
                    String Organizer = apptItem.Organizer;
                    SetLabel(apptItem.Subject, 1); //max 1024 characters.
                    SetLabel(Organizer, 2);
                    return;
                }
            }
            else if (selObject is MeetingItem) // Meetings in Mailliste + Doppel-klick auf Mails
            {
                System.Diagnostics.Debug.WriteLine("erkannt - MeetingItem");
                MeetingItem meetItem = (selObject as MeetingItem);
                if (meetItem != null && meetItem.EntryID != null)
                {
                    String RequiredAttendees = meetItem.Recipients[1].Address;
                    String Organizer = meetItem.SenderEmailAddress;
                    SetLabel(meetItem.Subject, 1); //max 1024 characters.
                    SetLabel(Organizer, 2);
                    return;
                }
            }
            if (last_exec + 500 < new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds())
            {
                SetLabel(null, 1);
                SetLabel(null, 2);
            }
        }

        private void SetLabel(string text, int label_int)
        {
            if (!String.IsNullOrEmpty(text))
            {
                switch (label_int)
                {
                    case 1:
                        _ribbon.Label1 = text;
                        break;
                    case 2:
                        _ribbon.Label2 = text;
                        break;
                }
                
            }
            else
            {
                switch (label_int)
                {
                    case 1:
                        _ribbon.Label1 = "no data found";
                        break;
                    case 2:
                        _ribbon.Label2 = "no data found";
                        break;
                }
            }
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
