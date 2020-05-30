using System;
using Microsoft.Office.Interop.Outlook;
namespace MeetingInfo
{
    public class ExplorerWrapper
    {
        private Explorer explorer;
        private long last_exec = 0;

        public ExplorerWrapper(Explorer Explorer)
        {
            explorer = Explorer;

            ((ExplorerEvents_Event)explorer).Close +=
                new ExplorerEvents_CloseEventHandler(ExplorerWrapper_Close);

            ((ExplorerEvents_10_Event)explorer).Activate += new ExplorerEvents_10_ActivateEventHandler(ExplorerWrapper_Activate);

            explorer.SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
        }

        void Explorer_SelectionChange()
        {
            if (explorer.Selection.Count == 1 && last_exec + 250 < new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds())
            {
                Object selObject = explorer.Selection[1];
                Globals.ThisAddIn.CheckObject(selObject);
                last_exec = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();
            }
        }

        void ExplorerWrapper_Activate()
        {
            Explorer_SelectionChange();
        }

        void ExplorerWrapper_Close()
        {
            Globals.ThisAddIn.ExplorerWrappers.Remove(explorer);
            ((ExplorerEvents_Event)explorer).Close -=
                new ExplorerEvents_CloseEventHandler(ExplorerWrapper_Close);
            ((ExplorerEvents_10_Event)explorer).Activate -=
                new ExplorerEvents_10_ActivateEventHandler(ExplorerWrapper_Activate);
            explorer.SelectionChange -= new ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
            explorer = null;
        }
    }
}
