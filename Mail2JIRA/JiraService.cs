using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using System.Net;
using Mail2JIRA.JiraDTO;
using System.IO;

namespace Mail2JIRA
{
    /// <summary>
    /// JIRA 5.0 REST remote API wrapper.
    /// <seealso cref="https://developer.atlassian.com/display/JIRADEV/JIRA+REST+APIs"/>
    /// </summary>
    public class JiraService
    {
        private const string CREATE_ISSUE_PATH = "issue";

        private Uri uri;
        private string username;
        private string password;
        private string challenge;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="uri">Service uri</param>
        /// <param name="username">Username to login</param>
        /// <param name="password">Password to login</param>
        public JiraService(Uri uri, string username, string password)
        {
            this.uri = uri;
            this.username = username;
            this.password = password;
            this.challenge = Convert.ToBase64String(Encoding.ASCII.GetBytes(String.Concat(username, ":", password)));
        }

        /// <summary>
        /// Creating a new Issue.
        /// </summary>
        /// <param name="obj">Create issue request object</param>
        /// <returns>Create issue response object</returns>
        public CreateIssueResponseObject CreateIssue(CreateIssueRequestObject obj)
        {
            UriBuilder ub = new UriBuilder(uri);
            ub.Path += CREATE_ISSUE_PATH;

            JsonSerializerSettings settings = new JsonSerializerSettings();
            settings.ContractResolver = new LowercaseContractResolver();
            string json = JsonConvert.SerializeObject(obj, Formatting.None, settings);

            HttpWebRequest webRequest = (HttpWebRequest)HttpWebRequest.Create(ub.Uri);
            webRequest.Method = "POST";
            webRequest.ContentType = "application/json";
            webRequest.Headers.Add("Authorization", "Basic " + challenge);
            webRequest.AllowAutoRedirect = true;
            webRequest.Timeout = 30000;

            Stream buffer = webRequest.GetRequestStream();
            byte[] bytesToWrite = Encoding.UTF8.GetBytes(json);
            buffer.Write(bytesToWrite, 0, bytesToWrite.Length);
            buffer.Close();

            HttpWebResponse webResponse = null;
            StreamReader rcvdStream = null;
            try
            {
                webResponse = (HttpWebResponse)webRequest.GetResponse();
                rcvdStream = new StreamReader(webResponse.GetResponseStream());

                return JsonConvert.DeserializeObject<CreateIssueResponseObject>(rcvdStream.ReadToEnd());
            }
            catch
            {
                throw;
            }
            finally 
            {
                if(webResponse != null)
                    webResponse.Close();
                if(rcvdStream != null)
                    rcvdStream.Close();
            }
        }
    }
}
