using System;
using System.Collections.Generic;
using System.Text;

namespace OneNoteC2Client
{
    class ClassProperties
    {
    }
    public class getuserid
    {
        public string id { get; set; }
    }

    public class sectiondetails
    {
        public string id { get; set; }
        public string displayName { get; set; }
    }

    public class pagedetails
    {
        public string id { get; set; }
        public string title { get; set; }
        public string content { get; set; }
    }
    public class createcontent
    {
        public string target { get; set; }
        public string action { get; set; }
        public string content { get; set; }
        //public messagebody body { get; set; }
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

    public static class OidcConstants
    {
        public const string AdditionalClaims = "claims";
        public const string ScopeOfflineAccess = "offline_access";
        public const string ScopeProfile = "profile";
        public const string ScopeOpenId = "openid";
    }
}
