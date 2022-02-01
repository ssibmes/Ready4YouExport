using SolEolImportExport.Domain;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.XPath;
namespace SolEolImportExport
{
    public class XmlClient
    {
        #region Fields
        private readonly List<Company> _companies = new List<Company>();
        private CookieContainer _cookieContainer = new CookieContainer();
        private Topics _xmlTopics = new Topics();
        #endregion

        #region Properties
        public Credentials Credentials { get; set; }
        public ProxyServer ProxyServer { get; set; }
        public XmlFile XmlFile { get; set; }
        #endregion Properties

        #region Constructor

        public XmlClient()
        {
            ProxyServer = new ProxyServer();
        }
        #endregion

        #region Public methods
        public bool SelectCompany(string url, long company, string accessToken)
        {
            if (string.IsNullOrEmpty(url)) return false;
            return SwitchCompany(AddSlashToUrl(url), company, accessToken);
        }
        public string ImportFile(string url, string applicationKey, string topic, string fileName, int Companycode)
        {
            if (string.IsNullOrEmpty(url)) return "";
            return ImportFile(url, applicationKey, topic, null, fileName, Companycode);
        }
        public string ImportFile(string url, string applicationKey, string topic, Dictionary<string, string> parameterValues, string fileName, int CompanyCode)
        {
            if (string.IsNullOrEmpty(url)) return "";
            return UploadFile(AddSlashToUrl(url), applicationKey, topic, parameterValues, fileName, CompanyCode);
        }


        //public void GetCompanies(string url)
        //{
        //    if (string.IsNullOrEmpty(url)) return;
        //    url = AddSlashToUrl(url);

        //    _companies.Clear();
        //    if (!Login(url)) throw new Exception("Access denied");
        //    SetCompanies(url);
        //}

        //public void GetTopics(string url)
        //{
        //    if (string.IsNullOrEmpty(url)) return;
        //    SetTopics(AddSlashToUrl(url));
        //}

        //public string ExportFile(string url, string accessToken, string topic, int companycode, bool readAttachments, string timeStamp, string journal)
        //{
        //    if (string.IsNullOrEmpty(url)) return "";
        //    return ExportFile(url, topic, null, readAttachments, companycode, accessToken, timeStamp, journal);
        //}

        //public string ExportFile(string url, string topic, Dictionary<string, string> parameterValues, bool readAttachments, int companyCode, string accessToken, string timeStamp, string journal)
        //{
        //    if (string.IsNullOrEmpty(url)) return null;
        //    return DownloadFile(AddSlashToUrl(url), accessToken, topic, parameterValues, readAttachments, companyCode, timeStamp, journal);
        //}

        #endregion

        #region Private Methods

        private bool Login(string baseUrl)
        {
            _cookieContainer = new CookieContainer();

            var url = string.Format("{0}/docs/XMLDivisions.aspx", baseUrl);
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";
            request.AllowWriteStreamBuffering = true;
            request.Proxy = ProxyServer.GetProxy();

            var credentials = string.Format("_UserName_={0}&_Password_={1}", Credentials.UserName, Credentials.Password);
            var data = Encoding.UTF8.GetBytes(credentials);
            using (var requestStream = request.GetRequestStream())
            {
                requestStream.Write(data, 0, data.Length);
            }

            Debug.WriteLine("Login");
            Debug.WriteLine("POST");
            Debug.WriteLine(url);

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                var responseUri = response.ResponseUri.AbsoluteUri;
                request.Abort();
                return (string.Compare(responseUri, url, StringComparison.Ordinal) == 0);
            }
        }

        private void SetCompanies(string baseUrl)
        {
            var url = string.Format("{0}/docs/XMLDivisions.aspx", baseUrl);
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.Method = "GET";

            Debug.WriteLine("Get divisions");
            Debug.WriteLine("GET");
            Debug.WriteLine(url);

            var response = (HttpWebResponse)request.GetResponse();

            using (var stream = response.GetResponseStream())
            {
                if (stream != null)
                {
                    var xml = new XmlDocument();
                    xml.Load(stream);
                    var list = xml.SelectNodes("/Administrations/Administration");
                    if (list != null)
                    {
                        foreach (XmlNode node in list)
                        {
                            AddCompany(node);
                        }
                    }
                }
            }
            request.Abort();
            response.Close();
        }

        private bool SwitchCompany(string baseUrl, long company, string accessToken)
        {
            //The Remember parameter with a value '3' is needed to be able to switch divisions without affecting the last used company,
            //and to do nothing when the company has been switched already to the correct one.
            var url = string.Format("{0}/docs/ClearSession.aspx?Division={1}&Remember=3", baseUrl, company);
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.Method = "GET";
            request.Headers.Add("Authorization", "Bearer" + " " + accessToken);
            Debug.WriteLine("Switch division");
            Debug.WriteLine("GET");
            Debug.WriteLine(url);
            var response = (HttpWebResponse)request.GetResponse();
            var succeeded = CheckCompanyIsSet(baseUrl, company, accessToken);
            request.Abort();
            response.Close();
            return succeeded;
        }

        private bool CheckCompanyIsSet(string baseUrl, long company, string accessToken)
        {
            var url = string.Format("{0}/docs/XMLDivisions.aspx", baseUrl);
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.Method = "GET";
            request.Headers.Add("Authorization", "Bearer" + " " + accessToken);
            Debug.WriteLine("Get divisions");
            Debug.WriteLine("GET");
            Debug.WriteLine(url);

            var response = (HttpWebResponse)request.GetResponse();
            var stream = response.GetResponseStream();
            if (stream != null)
            {
                var nav = new XPathDocument(stream).CreateNavigator();
                var current = nav.SelectSingleNode("/Administrations/Administration[@Current=\"True\"]");
                if (current != null)
                {
                    var code = current.GetAttribute("Code", string.Empty);
                    return String.CompareOrdinal(code, company.ToString(CultureInfo.InvariantCulture)) == 0;
                }
            }
            response.Close();
            request.Abort();
            return false;
        }

        private void AddCompany(XmlNode node)
        {
            if (node.Attributes == null) return;

            var code = Convert.ToInt64(node.Attributes["Code"].Value);
            var hid = Convert.ToInt64(node.Attributes["HID"].Value);
            var current = node.Attributes["Current"] != null && Convert.ToBoolean(node.Attributes["Current"].Value);
            var descNode = node.SelectSingleNode("Description");
            var description = (descNode != null) ? descNode.InnerText : string.Empty;

            _companies.Add(new Company(code, hid, description, current));
        }

        private void SetTopics(string baseUrl)
        {
            //This aspx will retrieve all available topics and their import/export parameters.
            var url = string.Format("{0}/docs/XMLTopicParameters.aspx", baseUrl);
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.Method = "GET";
            try
            {
                Debug.WriteLine("Get XML topics");
                Debug.WriteLine("GET");
                Debug.WriteLine(url);

                var response = (HttpWebResponse)request.GetResponse();
                using (var stream = response.GetResponseStream())
                {
                    if (stream != null)
                    {
                        var readStream = new StreamReader(stream, Encoding.UTF8);
                        var serializer = new XmlSerializer(typeof(Topics));
                        _xmlTopics = (Topics)serializer.Deserialize(readStream);
                    }
                }
                response.Close();
            }
            catch (WebException ex)
            {
                XmlFile.ErrorMessages = ex.Message;
            }
            request.Abort();
        }

        private string UploadFile(string baseUrl, string applicationKey, string topic, Dictionary<string, string> parameterValues, string fileName, int CompanyCode)
        {
            string returnMessage = "";
            var url = string.Format("{0}/docs/XMLUpload.aspx?Topic={1}&_Division_={2}", baseUrl, topic, CompanyCode.ToString());
            url = AddParameterValuesToUrl(url, parameterValues);

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";
            request.Headers.Add("Authorization", "Bearer" + " " + applicationKey);
            request.AllowWriteStreamBuffering = true;
            request.Proxy = ProxyServer.GetProxy();
			Console.WriteLine("1");

			using (var requestStream = request.GetRequestStream())
            {
				
                using (var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    var data = new byte[4096];
                    var bytesRead = fileStream.Read(data, 0, data.Length);
                    while (bytesRead > 0)
                    {
                        requestStream.Write(data, 0, bytesRead);
                        bytesRead = fileStream.Read(data, 0, data.Length);
                    }
                }
            }

            Debug.WriteLine("Upload");
            Debug.WriteLine("POST");
            Debug.WriteLine(url);
			Console.WriteLine("Upload");
			try
            {
                var response = (HttpWebResponse)request.GetResponse();
                if (response.ContentType != "text/xml")
                {
                    using (var stream = response.GetResponseStream())
                    {
                        XmlFile.GetErrorsFromHtml(stream);
                    }
                }
                else
                {
                    using (var stream = response.GetResponseStream())
                    {
                        XmlFile.GetMessages(stream);
                        returnMessage = XmlFile.ErrorMessages;
                    }
                }
                response.Close();
            }
            catch (Exception ex)
            {
                XmlFile.ErrorMessages = ex.Message + "\n " + ex.InnerException;
				Console.WriteLine(ex.Message + "\n " + ex.InnerException);
                //UtilityLayer.Common.ErrorLog(ex.Message + " " + ex.InnerException + " " + "XmlClient - UploadFile");

            }
			
			request.Abort();
            return returnMessage;
        }

        private string DownloadFile(string baseUrl, string applicationKey, string topic, Dictionary<string, string> parameterValues, bool readAttachments, int CompanyCode, string timeStamp, string journal)
        {
            var url = string.Format("{0}/docs/XMLDownload.aspx?Topic={1}&_Division_={2}&output=1&Params_Documents=0", baseUrl, topic, CompanyCode.ToString());
            if(journal !="")
             url = string.Format("{0}/docs/XMLDownload.aspx?Topic={1}&_Division_={2}&output=1&Params_Documents=0&Params_Journal={3}", baseUrl, topic, CompanyCode.ToString(),journal);

            if (timeStamp != "" && XmlFile.PagingTimestamp == "")
            {
                url = AddPagingTimestampToUrl(url, timeStamp);
            }
            if (XmlFile.PagingTimestamp != "")
            {
                url = AddPagingTimestampToUrl(url, XmlFile.PagingTimestamp);
            }
            url = AddParameterValuesToUrl(url, parameterValues);
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = _cookieContainer;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "GET";
            request.Headers.Add("Authorization", "Bearer" + " " + applicationKey);
            request.AllowWriteStreamBuffering = true;
            request.Proxy = ProxyServer.GetProxy();

            var buffer = string.Empty;
            try
            {
                Debug.WriteLine("Download");
                Debug.WriteLine("GET");
                Debug.WriteLine(url);

                var response = (HttpWebResponse)request.GetResponse();
                using (var stream = response.GetResponseStream())
                {
                    if (response.ContentType != "text/xml")
                    {
                        XmlFile.GetErrorsFromHtml(stream);
                    }
                    else
                    {
                        buffer = XmlFile.ReadStreamToBuffer(stream, readAttachments, topic);
                    }
                }

                response.Close();
            }
            catch (WebException ex)
            {
                XmlFile.ErrorMessages = ex.Message;
            }
            request.Abort();

            return buffer;
        }

        private static string AddParameterValuesToUrl(string url, Dictionary<string, string> parameterValues)
        {
            if (parameterValues != null)
            {
                url = parameterValues.Aggregate(url, (current, parameter) => current + string.Format("&{0}={1}", parameter.Key, HttpUtility.UrlEncode(parameter.Value)));
            }
            return url;
        }

        private static string AddPagingTimestampToUrl(string url, string pagingTimestamp)
        {
            if (pagingTimestamp.Length > 0)
            {
                url += string.Format("&TSPaging={0}", pagingTimestamp);
            }

            return url;
        }

        private static string AddSlashToUrl(string url)
        {
            return (url.EndsWith("/") ? url.Remove(url.Length - 1) : url);
        }

        #endregion
    }
}
