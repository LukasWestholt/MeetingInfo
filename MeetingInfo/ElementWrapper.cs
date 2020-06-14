namespace MeetingInfo
{
    public class ElementWrapper
    {
        private string _label = string.Empty;
        private string _screentip = string.Empty;
        private string _supertip = string.Empty;
        private bool _visible = true;
        private System.Drawing.Bitmap _image;
        private Microsoft.Office.Interop.Outlook.AppointmentItem _apptItem;
        private readonly Ribbon _ribbon;

        public ElementWrapper(Ribbon _ribbon)
        {
            this._ribbon = _ribbon;
        }

        public string Label
        {
            get
            {
                if (_label != string.Empty)
                    return _label;
                else
                    return "NULL";
            }
            set
            {
                if (_label != value)
                {
                    _label = value;
                    Update();
                }
            }
        }

        public string Screentip
        {
            get
            {
                if (_screentip != string.Empty)
                    return _screentip;
                else
                    return "NULL";
            }
            set
            {
                if (_screentip != value)
                {
                    _screentip = value;
                    Update();
                }
            }
        }

        public string Supertip
        {
            get
            {
                if (_supertip != string.Empty)
                    return _supertip;
                else
                    return "NULL";
            }
            set
            {
                if (_supertip != value)
                {
                    _supertip = value;
                    Update();
                }
            }
        }

        public bool Visible
        {
            get
            {
                return _visible;
            }
            set
            {
                if (_visible != value)
                {
                    _visible = value;
                    Update();
                }
            }
        }

        public System.Drawing.Bitmap Image
        {
            get
            {
                return _image;
            }
            set
            {
                if (_image != value)
                {
                    _image = value;
                    Update();
                }
            }
        }

        public Microsoft.Office.Interop.Outlook.AppointmentItem AppointmentItem
        {
            get
            {
                return _apptItem;
            }
            set
            {
                if (_apptItem != value)
                {
                    _apptItem = value;
                }
            }
        }

        public void Update()
        {
            _ribbon.Invalidate();
        }

    }
}
