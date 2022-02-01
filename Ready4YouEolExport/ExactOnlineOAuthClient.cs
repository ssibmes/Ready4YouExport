using DotNetOpenAuth.OAuth2;
using Newtonsoft.Json.Linq;
using SolEolImportExport.Domain;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.SessionState;
using System.Windows.Forms;

namespace SolEolImportExport.Service
{
    /// <summary>
    /// Our pages (SalesInvoiceNew.aspx, SalesInvoiceEdit.aspx, SalesInvoiceList.aspx) each contain an instance of this class.
    /// The authorization state is shared through the session.
    /// </summary>
    public class ExactOnlineOAuthClient : WebServerClient
    {
        #region Properties
        public IAuthorizationState Authorization { get; set; }
        public static string clientId { get; set; }
        public static string clientSecret { get; set; }

        #endregion

        #region Constructor
        public ExactOnlineOAuthClient()
            : base(CreateAuthorizationServerDescription(), clientId, clientSecret)
        {
            // initialization is already done through the base constructor
            ClientCredentialApplicator = ClientCredentialApplicator.PostParameter(clientSecret);
        }
        #endregion

        #region Public Methods
        System.Collections.Generic.Dictionary<int, token_record> token_records = new System.Collections.Generic.Dictionary<int, token_record>();
        class token_record
        {
            // Now it's based on key int CompanyCode, but I would changed it to: (string client_id, because it's uniek)

            public int CompanyCode { get; set; }
            public string refresh_token { get; set; }
            public string access_token { get; set; }
            public string token_type { get; set; }
            public string expires_in { get; set; }
            public DateTime access_token_expires { get; set; }

            public token_record(int CompanyCode)
            {
                this.CompanyCode = CompanyCode;
                refresh_token = "";
                access_token = "";
                token_type = "";
                expires_in = "300";
                access_token_expires = new DateTime();
            }
        }
        public string GetAccessToken(string refreshToken, bool regenerateToken, int CompanyCode)
        {
            token_record tr = (token_records.ContainsKey(CompanyCode)) ? token_records[CompanyCode] : new token_record(CompanyCode);
            {
                if (tr.access_token_expires > DateTime.Now) return tr.access_token;
            }

            HttpWebRequest webRequest = null;
            HttpWebResponse webResponse = null;
            string responseBody = "";
            string accessToken = "";
            string refreshTokenNew = "";
            string tempRefreshToken = "";
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                const int SecurityProtocolTypeTls11 = 768;
                const int SecurityProtocolTypeTls12 = 3072;
                webRequest = (HttpWebRequest)WebRequest.Create("https://start.exactonline.nl/api/oauth2/token");
                ServicePointManager.SecurityProtocol |= (SecurityProtocolType)(SecurityProtocolTypeTls12 | SecurityProtocolTypeTls11);
                webRequest.Method = "POST";
                webRequest.ContentType = "application/x-www-form-urlencoded";

                // Wierd
                //webRequest.UserAgent = "SOLProcess.API";
                webRequest.Headers.Add("Cache-Control: no-store,no-cache");
                webRequest.Headers.Add("Pragma: no-cache");
                webRequest.ServicePoint.Expect100Continue = false;
                //refreshTokenNew = FetchRefreshToken();


                using (StreamWriter writer = new StreamWriter(webRequest.GetRequestStream()))
                {
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    sb.Append("grant_type=refresh_token");
                    //sb.AppendFormat("&refresh_token={0}", HttpUtility.UrlEncode(refreshTokenNew));

                    //sb.AppendFormat("grant_type={0}", HttpUtility.UrlEncode("refresh_token"));
                    if (regenerateToken)
                    {
                        string tokens = FetchRefreshToken();
                        string oldrefreshToken = null;
                        if (!string.IsNullOrEmpty(tokens))
                        {
                            string[] invidualToken = tokens.Split(new string[] { ":#" }, StringSplitOptions.None);
                            oldrefreshToken = invidualToken[0];
                            refreshTokenNew = oldrefreshToken;
                        }
                        sb.AppendFormat("&refresh_token={0}", HttpUtility.UrlEncode(refreshTokenNew));
                    }
                    else
                        sb.AppendFormat("&refresh_token={0}", HttpUtility.UrlEncode(refreshToken));
                    sb.AppendFormat("&client_id={0}", HttpUtility.UrlEncode(clientId));
                    sb.AppendFormat("&client_secret={0}", HttpUtility.UrlEncode(clientSecret));
                    sb.AppendFormat("&format={0}", HttpUtility.UrlEncode("xml"));
                    writer.WriteLine(sb.ToString());
                }
                webResponse = (HttpWebResponse)webRequest.GetResponse();
                using (StreamReader reader = new StreamReader(webResponse.GetResponseStream()))
                {
                    responseBody = reader.ReadToEnd();
                    if (responseBody != "" && responseBody != null)
                    {
                        JObject auth = JObject.Parse(responseBody);
                        accessToken = auth["access_token"].ToString();
                        // Wierd
                        tr.access_token = accessToken;
                        tr.token_type = auth["token_type"].ToString(); // default "bearer", for now not used, but it could changed somtime
                        tr.expires_in = auth["expires_in"].ToString();
                        int int_expires_in = 0;
                        int.TryParse(tr.expires_in, out int_expires_in);
                        if (int_expires_in == 0) int_expires_in = 300;
                        tr.access_token_expires = DateTime.Now.AddSeconds(int_expires_in - 10); // 600 or 300  minus 10 to be save
                        // Wierd
                        if (regenerateToken)
                            tempRefreshToken = auth["refresh_token"].ToString();
                        else
                            tempRefreshToken = refreshTokenNew;
                        tr.refresh_token = tempRefreshToken;
                        token_records.Add(CompanyCode, tr);
                        string updateToken = tempRefreshToken + ":#" + tr.access_token + ":#" + tr.access_token_expires.ToString("dd-MM-yyyy HH:mm:ss");
                        if (!string.IsNullOrEmpty(updateToken) && regenerateToken)
                            UpdateRefreshToken(updateToken);
                    }
                }
            }
            catch (WebException webEx)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(webEx.Message);
                sb.AppendLine();
                sb.AppendLine("REQUEST: ");
                sb.AppendLine();

                sb.AppendLine(string.Format("Request URL: {0} {1}", webRequest.Method, webRequest.Address));
                sb.AppendLine("Headers:");
                foreach (string header in webRequest.Headers)
                {
                    sb.AppendLine(header + ": " + webRequest.Headers[header]);
                }

                sb.AppendLine();
                sb.AppendLine("RESPONSE: ");
                sb.AppendLine();

                sb.AppendLine(string.Format("Status: {0}", webEx.Status));

                if (null != webEx.Response)
                {
                    HttpWebResponse response = (HttpWebResponse)webEx.Response;

                    sb.AppendLine(string.Format("Status Code: {0} {1}", (int)response.StatusCode, response.StatusDescription));
                    if (0 != webEx.Response.ContentLength)
                    {
                        using (var stream = webEx.Response.GetResponseStream())
                        {
                            if (null != stream)
                            {
                                using (var reader = new StreamReader(stream))
                                {
                                    sb.AppendLine(string.Format("Response: {0}", reader.ReadToEnd()));
                                }
                            }
                        }
                    }
                }

            }
            return accessToken;
        }
        #endregion

        #region Private Methods
        private static AuthorizationServerDescription CreateAuthorizationServerDescription()
        {
            var baseUri = WebConfigurationManager.AppSettings["BaseUri"];
            var uri = new Uri(baseUri.EndsWith("/") ? baseUri : baseUri + "/");
            var serverDescription = new AuthorizationServerDescription
            {
                AuthorizationEndpoint = new Uri(uri, "api/oauth2/auth"),
                TokenEndpoint = new Uri(uri, "api/oauth2/token")
            };

            return serverDescription;
        }
        private string FetchRefreshToken()
        {
            try
            {
                System.Configuration.AppSettingsReader appConfig = new System.Configuration.AppSettingsReader();
                FileInfo tokenFile = new FileInfo(appConfig.GetValue("TokenFilePath", Type.GetType("System.String")).ToString());

                if (tokenFile.Exists)
                {
                    StringBuilder refreshToken = new StringBuilder();
                    using (StreamReader sr = new StreamReader(tokenFile.FullName))
                    {
                        var myStringRow = sr.ReadLine();
                        while (myStringRow != null)
                        {
                            refreshToken.Append(myStringRow);
                            myStringRow = sr.ReadLine();
                        }
                        return refreshToken.ToString();
                    }
                }
                else
                {
                    MessageBox.Show("Token File Not Exists in the given path \n" + tokenFile.FullName, "File Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //var writer = File.CreateText(fileInfo.FullName);
                    //writer.WriteLine("");
                    //writer.Flush();
                    //writer.Close();
                    return "";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                using (var writer = File.AppendText("ErrorLog.txt"))
                {
                    writer.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss:ffffff") + ": " + ex.Message);
                    writer.Flush();
                    writer.Close();
                }
                return "";
            }
        }

        public string FetchExpiredTime()
        {
            try
            {
                System.Configuration.AppSettingsReader appConfig = new System.Configuration.AppSettingsReader();
                FileInfo tokenFile = new FileInfo(appConfig.GetValue("TokenFilePath", Type.GetType("System.String")).ToString());

                if (tokenFile.Exists)
                {
                    StringBuilder refreshToken = new StringBuilder();
                    using (StreamReader sr = new StreamReader(tokenFile.FullName))
                    {
                        var myStringRow = sr.ReadLine();
                        while (myStringRow != null)
                        {
                            refreshToken.Append(myStringRow);
                            myStringRow = sr.ReadLine();
                        }
                        return refreshToken.ToString();
                    }
                }
                else
                {
                    MessageBox.Show("Token File Not Exists in the given path \n" + tokenFile.FullName, "File Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //var writer = File.CreateText(fileInfo.FullName);
                    //writer.WriteLine("");
                    //writer.Flush();
                    //writer.Close();
                    return "";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                using (var writer = File.AppendText("ErrorLog.txt"))
                {
                    writer.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss:ffffff") + ": " + ex.Message);
                    writer.Flush();
                    writer.Close();
                }
                return "";
            }
        }

        private void UpdateRefreshToken(string refreshToken)
        {
            try
            {
                System.Configuration.AppSettingsReader appConfig = new System.Configuration.AppSettingsReader();
                FileInfo tokenFile = new FileInfo(appConfig.GetValue("TokenFilePath", Type.GetType("System.String")).ToString());

                var writer = File.CreateText(tokenFile.FullName);
                writer.WriteLine(refreshToken);
                writer.Flush();
                writer.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                using (var writer = File.AppendText("ErrorLog.txt"))
                {
                    writer.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss:ffffff") + ": " + ex.Message);
                    writer.Flush();
                    writer.Close();
                }
            }

        }
        #endregion
    }
}
