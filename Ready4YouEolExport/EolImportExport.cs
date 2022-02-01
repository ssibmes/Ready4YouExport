using SolEolImportExport;
using SolEolImportExport.Service;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Xml.Linq;

namespace Ready4YouEolExport
{
    class EolImportExport
    {
        public string ImportFiles(string file, string topic, string url, int companyCode, string clientId, string clientSecret)
        {
            try
            {
                XmlClient _xmlClient = new XmlClient();
                ExactOnlineOAuthClient.clientId = clientId;
                ExactOnlineOAuthClient.clientSecret = clientSecret;
                ExactOnlineOAuthClient objExactOnlineOAuthClient = new ExactOnlineOAuthClient();
                string refreshToken = string.Empty;
                string tokens = objExactOnlineOAuthClient.FetchExpiredTime();
                string oldAccessToken = null, oldrefreshToken = null, oldExpiredTime = null;
                string accessToken = string.Empty;
                if (!string.IsNullOrEmpty(tokens))
                {
                    string[] invidualToken = tokens.Split(new string[] { ":#" }, StringSplitOptions.None);
                    if (invidualToken.Length == 3)
                    {
                        oldrefreshToken = invidualToken[0];
                        oldAccessToken = invidualToken[1];
                        oldExpiredTime = invidualToken[2];
                    }
                }
                DateTime expiredDate;
                var tokenexpired = false;
                if (!string.IsNullOrEmpty(oldAccessToken) && !string.IsNullOrEmpty(oldExpiredTime))
                {
                    DateTime _date;
                    string[] formats1 = new string[11] { "yyyy-MM-dd HH:mm:ss", "dd-MM-yyyy HH:mm:ss", "MM-dd-yyyy HH:mm:ss", "yyyy/MM/dd HH:mm:ss", "MM/dd/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm:ss", "yyyy.MM.dd HH:mm:ss", "MM.dd.yyyy HH:mm:ss", "dd.MM.yyyy HH:mm:ss", "d-M-yyyy HH:mm:ss", "M-d-yyyy HH:mm:ss" };
                    DateTime.TryParseExact(oldExpiredTime, formats1, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out expiredDate);
                    DateTime.TryParseExact(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), formats1, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out _date);

                    if (_date < expiredDate.AddSeconds(-20))
                    {
                        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                        const int SecurityProtocolTypeTls11 = 768;
                        const int SecurityProtocolTypeTls12 = 3072;

                        System.Net.ServicePointManager.SecurityProtocol |= (System.Net.SecurityProtocolType)(SecurityProtocolTypeTls12 | SecurityProtocolTypeTls11);
                        accessToken = oldAccessToken;
                    }
                    else
                        tokenexpired = true;
                }
                else
                    tokenexpired = true;

                if (tokenexpired)
                    accessToken = objExactOnlineOAuthClient.GetAccessToken(refreshToken, true, companyCode);
                string strMessages = "";
                DateTime expireTime = DateTime.Now.AddMinutes(9);

                _xmlClient.XmlFile = new SolEolImportExport.Domain.XmlFile();

                if (_xmlClient.SelectCompany(url, companyCode, accessToken))
                {
                    if (expireTime == DateTime.Now)
                    {
                        accessToken = objExactOnlineOAuthClient.GetAccessToken(refreshToken, true, companyCode);
                        expireTime = DateTime.Now.AddMinutes(9);
                    }
                    strMessages = _xmlClient.ImportFile(url, accessToken, topic, file, companyCode);
                }
                _xmlClient.XmlFile = null;
                return strMessages;
            }
            catch (WebException e)
            {
                Console.WriteLine("This program is expected to throw WebException on successful run." +
                                    "\n\nException Message :" + e.Message);
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    Console.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    Console.WriteLine("Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
                throw;
            }
            catch
            {
                throw;
            }
        }
    }
}
