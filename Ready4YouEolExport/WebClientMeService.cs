using DotNetOpenAuth.OAuth2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace SolEolImportExport.Service
{
    public class WebClientMeService : WebClientServiceBase
    {
        #region Properties

        public static Uri ServiceUri
        {
            get { return new Uri(BaseUri, "api/v1/current/Me"); }
        }

        #endregion

        #region Constructor

        public WebClientMeService(IAuthorizationState authorizationState)
            : base(authorizationState)
        {
        }

        #endregion

        #region Public Methods

        public int GetCurrentCompany()
        {
            var uriBuilder = new UriBuilder(ServiceUri) { Query = "$select=CurrentDivision" };
            string jsonResponse = WebClient.DownloadString(uriBuilder.Uri);
            return ParseJson(jsonResponse);
        }

        #endregion

        #region Private Methods

        private static int ParseJson(string jsonString)
        {
            var serializer = new JavaScriptSerializer();
            var jsonObject = serializer.DeserializeObject(jsonString) as Dictionary<string, object>;
            var jsonDictionary = (Dictionary<string, object>)jsonObject["d"];
            var results = (object[])jsonDictionary["results"];
            var result = (Dictionary<string, object>)results.First();
            return Convert.ToInt32(result["CurrentDivision"]);
        }

        #endregion

    }
}