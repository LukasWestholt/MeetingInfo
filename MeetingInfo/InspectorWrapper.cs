using System;
using Microsoft.Office.Interop.Outlook;
namespace MeetingInfo
{
    public class InspectorWrapper
    {
        private Inspector inspector;
        private long last_exec = 0;

        public InspectorWrapper(Inspector Inspector)
        {
            inspector = Inspector;

            ((InspectorEvents_Event)inspector).Close +=
                new InspectorEvents_CloseEventHandler(InspectorWrapper_Close);

            ((InspectorEvents_10_Event)inspector).Activate += new InspectorEvents_10_ActivateEventHandler(InspectorWrapper_Activate);

            Globals.MeetingInfoMain.Event(inspector.CurrentItem);
            last_exec = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();
        }

        void InspectorWrapper_Activate()
        {
            if (last_exec + 500 < new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds())
            {
                Globals.MeetingInfoMain.Event(inspector.CurrentItem);
                last_exec = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();
            }
        }

        void InspectorWrapper_Close()
        {
            Globals.MeetingInfoMain.InspectorWrappers.Remove(inspector);
            ((InspectorEvents_Event)inspector).Close -=
                new InspectorEvents_CloseEventHandler(InspectorWrapper_Close);
            ((InspectorEvents_10_Event)inspector).Activate -=
                new InspectorEvents_10_ActivateEventHandler(InspectorWrapper_Activate);
            inspector = null;
        }
    }
}
