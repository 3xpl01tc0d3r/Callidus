using System;
using System.Collections.Generic;
using System.Text;

namespace TeamsC2
{
    class ClassProperties
    {
    }
    public class groupdetails
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string mail { get; set; }
    }

    public class channeldetails
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string email { get; set; }
    }

    public class channelmessages
    {
        public string id { get; set; }
        public messagebody body { get; set; }
    }
    public class createmessages
    {
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
