using System;
using System.Net;

namespace SolEolImportExport.Domain
{
    public class ProxyServer
    {
        public bool UseProxy { get; set; }

        public string Server { get; set; }

        public string Port { get; set; }

        public Credentials Credentials { get; set; }

        public IWebProxy GetProxy()
        {
            try
            {
                IWebProxy proxy = null;
                if (UseProxy)
                {
                    proxy = Port.Length > 0 ? new WebProxy(Server, Convert.ToInt32(Port)) : new WebProxy(Server);

                    if (string.IsNullOrEmpty(Credentials.Domain))
                    {
                        proxy.Credentials = new NetworkCredential(Credentials.UserName, Credentials.Password);
                    }
                    else
                    {
                        proxy.Credentials = new NetworkCredential(Credentials.UserName, Credentials.Password, Credentials.Domain);
                    }
                }
                return proxy;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
