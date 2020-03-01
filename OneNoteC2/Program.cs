using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading;
using System.Web;
using System.Security;
using System.Configuration;

namespace OneNoteC2
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
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }

        private static async Task RunAsync()
        {
            AuthenticationResult result = null;
            result = await Auth();
            DateTime exp = result.ExpiresOn.DateTime;


            if (result != null)
            {
                var pageid = await GetPageID(result.AccessToken);
                while (true)
                {
                    Console.Write("#> ");
                    string input = null;
                    input = Console.ReadLine().Trim();
                    while (input != null && input != "")
                    {
                        int exptime = DateTime.Compare(exp, DateTime.Now.ToUniversalTime().AddMinutes(10));
                        if (exptime < 0)
                        {
                            result = await Auth();
                            exp = result.ExpiresOn.DateTime;
                        }
                        await CreateTask(result.AccessToken, pageid,input);
                        input = null;
                        Thread.Sleep(2000);
                    }

                }
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
                    Console.WriteLine(ex.Message);
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

        public static async Task CreateTask(string AccessToken, string pageid, string taskcommand)
        {
            var httpClient = new HttpClient();
            var apiCaller = new ProtectedApiCallHelper(httpClient);
            string data = null;

            var url = $"https://graph.microsoft.com/beta/me/onenote/pages/{pageid}/content";
            createcontent createcontentdetails = new createcontent();

            createcontentdetails.target = "body";
            createcontentdetails.action = "prepend";
            createcontentdetails.content = $"<p data-tag='to-do'>{taskcommand}</p>";

            data = "[" + JsonConvert.SerializeObject(createcontentdetails) + "]";


            JObject createtaskoutput = await apiCaller.CallPatchWebApiAndProcessResultASync(url, AccessToken, data);

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
