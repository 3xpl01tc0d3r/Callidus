using System;
using System.Collections.Generic;
using System.Text;

namespace OutlookC2
{
    class ClassProperties
    {
    }
    public class getuserid
    {
        public string id { get; set; }
    }

    public class getmailfolderid
    {
        public string id { get; set; }
    }
    public class createmessages
    {
        public string subject { get; set; }
        public messagebody body { get; set; }
    }

    public class getmessages
    {
        public string id { get; set; }
        public string subject { get; set; }
        public messagebody body { get; set; }
    }

    public class messagebody
    {
        public string contentType { get; set; }
        public string content { get; set; }
    }
}
