using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace MeetingInfo
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbonUI;

        public Ribbon()
        {
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

        private string _label1 = string.Empty;
        public string Label1
        {
            get
            {
                if (this._label1 != string.Empty)
                    return this._label1;
                else
                    return "Default-text";
            }
            set
            {
                if (this._label1 != value)
                {
                    this._label1 = value;
                    this.Invalidate();
                }
            }
        }

        private string _label2 = string.Empty;
        public string Label2
        {
            get
            {
                if (this._label2 != string.Empty)
                    return this._label2;
                else
                    return "Default-text";
            }
            set
            {
                if (this._label2 != value)
                {
                    this._label2 = value;
                    this.Invalidate();
                }
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

        public void ButtonTest(Office.IRibbonControl ribbonUI)
        {
            System.Windows.Forms.MessageBox.Show("Hello!");
        }
        public string GetLabel(Office.IRibbonControl ribbonUI)
        {
            string[] first_labels = { "label1", "label2", "label3", "label4" };
            string[] second_labels = { "label11", "label22", "label33", "label44" };
            if (first_labels.Contains(ribbonUI.Id.ToLower()))
            {
                return this.Label1;
            } else if (second_labels.Contains(ribbonUI.Id.ToLower()))
            {
                return this.Label2;
            } else 
            {
                return "default";
            }
        }
        #endregion

        #region Hilfsprogramme
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
