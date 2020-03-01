using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OutlookC2Client
{
    public class ProtectedApiCallHelper
    {

        public ProtectedApiCallHelper(HttpClient httpClient)
        {
            HttpClient = httpClient;
        }

        protected HttpClient HttpClient { get; private set; }

        public async Task<JObject> CallGetWebApiAndProcessResultASync(string webApiUrl, string accessToken)
        {
            JObject result = null;
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.GetAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    result = JsonConvert.DeserializeObject(json) as JObject;
                    //Console.ForegroundColor = ConsoleColor.Gray;

                }
                else
                {
                    //Console.ForegroundColor = ConsoleColor.Red;
                    //Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    JObject content = JObject.Parse(await response.Content.ReadAsStringAsync());

                    //Console.WriteLine($"Content: {content}");
                    result = content;
                    
                }
                //Console.ResetColor();
            }
            return result;
        }

        public async Task<JObject> CallPostWebApiAndProcessResultASync(string webApiUrl, string accessToken, string data)
        {
            JObject result = null;
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                var body = new StringContent(data, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await HttpClient.PostAsync(webApiUrl, body);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    result = JsonConvert.DeserializeObject(json) as JObject;
                    //Console.ForegroundColor = ConsoleColor.Gray;
                }
                else
                {
                    //Console.ForegroundColor = ConsoleColor.Red;
                    //Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    JObject content = JObject.Parse(await response.Content.ReadAsStringAsync());
                    //Console.WriteLine($"Content: {content}");
                    result = content;
                }
                //Console.ResetColor();
            }
            return result;
        }

        public async Task<JObject> CallDeleteWebApiAndProcessResultASync(string webApiUrl, string accessToken)
        {
            JObject result = null;
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.DeleteAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    result = JsonConvert.DeserializeObject(json) as JObject;
                    //Console.ForegroundColor = ConsoleColor.Gray;
                }
                else
                {
                    //Console.ForegroundColor = ConsoleColor.Red;
                    //Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    JObject content = JObject.Parse(await response.Content.ReadAsStringAsync());
                    //Console.WriteLine($"Content: {content}");
                    result = content;

                }
                //Console.ResetColor();
            }
            return result;
        }
    }
}