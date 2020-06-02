using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace MeetingInfo
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbonUI;

        public ElementWrapper Label1;
        public ElementWrapper Label2;
        public ElementWrapper Label3;
        public ElementWrapper Label4;
        private readonly Dictionary<string[], ElementWrapper> labels;

        public Ribbon()
        {
            Label1 = new ElementWrapper(this);
            Label2 = new ElementWrapper(this);
            Label3 = new ElementWrapper(this);
            Label4 = new ElementWrapper(this);

            labels = new Dictionary<string[], ElementWrapper>()
            {
                { new []{ "label1", "label2", "label3", "label4" }, Label1 },
                { new []{ "label11", "label22", "label33", "label44" }, Label2 },
                { new []{ "label111", "label222", "label333", "label444" }, Label3 },
                { new []{ "label1111", "label2222", "label3333", "label4444" }, Label4 }
            };
        }
 
        public void Invalidate()
        {
            // you can tell Outlook to update the entire ribbon
            if (this._ribbonUI != null)
            {
                this._ribbonUI.Invalidate();
            } else
            {  // ERROR HANDLING
                ErrorMessage("Invalidate");
            }
        }

        public void Invalidate(string a)
        {
            // or you can tell Outlook to update a single tab/group/control
            if (this._ribbonUI != null) {
                this._ribbonUI.InvalidateControl(a);
            } else
            {
                ErrorMessage("Invalidate");
            }
        }

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            string ribbonXML = String.Empty;
            if (ribbonID == "Microsoft.Outlook.MeetingRequest.Read" || ribbonID == "Microsoft.Outlook.Appointment")
            {
                ribbonXML = GetResourceText("MeetingInfo.Ribbon.xml");
            } else if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                ribbonXML = GetResourceText("MeetingInfo.Ribbon_explorer.xml");
            }
            return ribbonXML;
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void OnRibbonLoaded(Office.IRibbonUI ribbonUI)
        {
            this._ribbonUI = ribbonUI;
        }

        // TODO Button
        public void OnButtonTest(Office.IRibbonControl ribbonUI)
        {
            System.Windows.Forms.MessageBox.Show("Hello!");
        }
        public string OnGetLabel(Office.IRibbonControl ribbonUI)
        {
            ElementWrapper elem = GetElement(ribbonUI);
            return elem != null ? elem.Label: MeetingInfoMain.DEFAULT_TEXT_LABEL;
        }

        public string OnGetScreentip(Office.IRibbonControl ribbonUI)
        {
            ElementWrapper elem = GetElement(ribbonUI);
            return elem != null ? elem.Screentip: MeetingInfoMain.DEFAULT_TEXT_SCREENTIP;
        }

        public bool OnGetVisible(Office.IRibbonControl ribbonUI)
        {
            ElementWrapper elem = GetElement(ribbonUI);
            return elem != null ? elem.Visible: MeetingInfoMain.DEFAULT_STATE_VISIBLE;
        }
        public System.Drawing.Bitmap OnGetImage(Office.IRibbonControl ribbonUI)
        {
            ElementWrapper elem = GetElement(ribbonUI);
            return elem.Image;
        }

        public bool OnGetShowImage(Office.IRibbonControl ribbonUI)
        {
            ElementWrapper elem = GetElement(ribbonUI);
            return elem != null && elem.Image != null;
        }

        #endregion

        #region Hilfsprogramme

        private ElementWrapper GetElement(Office.IRibbonControl ribbonUI)
        {
            foreach (KeyValuePair<string[], ElementWrapper> entry in labels)
            {
                if (entry.Key.Contains(ribbonUI.Id.ToLower()))
                {
                    return entry.Value;
                }
            }
            ErrorMessage("Unknown Element: " + ribbonUI.Id.ToLower());
            return null;
        }
        private static void ErrorMessage(string errortext)
        {
            System.Windows.Forms.MessageBox.Show("Error on [" + errortext + "] - Please deactivate this add-in (" + MeetingInfoMain.ADD_IN_NAME + ")");
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
