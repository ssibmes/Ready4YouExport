using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.XPath;

namespace SolEolImportExport.Domain
{
    public class XmlFile
    {
        #region Enums

        private enum MessageType
        {
            Error = 0,
            Warning = 1,
            Message = 2
        }

        #endregion

        #region Properties
        public string ErrorMessages { get; set; }
        public string Messages { get; set; }
        public bool MorePages { get; private set; }
        public bool DataFound { get; private set; }
        public string PagingTimestamp { get; set; }
        public List<Attachment> Attachments { get; set; }
        #endregion

        #region Constructor
        public XmlFile()
        {
            ErrorMessages = string.Empty;
            PagingTimestamp = string.Empty;
            DataFound = false;
            MorePages = false;
            Attachments = new List<Attachment>();
        }
        #endregion

        #region Public Methods
        public string ReadStreamToBuffer(Stream stream, bool readAttachments, string topic)
        {
            try
            {
                string buffer = string.Empty;
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        buffer = reader.ReadToEnd();
                        byte[] data = Encoding.UTF8.GetBytes(buffer);

                        if (!GetErrorMessages(new MemoryStream(data)))
                        {

                            PagingTimestamp = GetTimestamp(new MemoryStream(data));
                            int resultCount = GetResultCount(new MemoryStream(data));
                            int pageSize = GetPageSize(new MemoryStream(data));

                            DataFound = resultCount > 0;
                            MorePages = PagingTimestamp.Length > 0 && resultCount == pageSize;

                            if (readAttachments) ReadAttachments(new MemoryStream(data), topic);
                        }
                    }
                }
                return buffer;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void GetErrorsFromHtml(Stream stream)
        {
            try
            {
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        string buffer = reader.ReadToEnd();
                        string errorCode = string.Empty;
                        int pos = buffer.IndexOf("id=\"mode\"", StringComparison.Ordinal);
                        if (pos > 0)
                        {
                            pos = buffer.IndexOf("value=\"", pos, StringComparison.Ordinal);
                            if (pos > 0)
                            {
                                int pos2 = buffer.IndexOf("\"", pos + 7, StringComparison.Ordinal);
                                errorCode = buffer.Substring(pos + 7, pos2 - pos - 7);
                            }
                        }

                        switch (errorCode)
                        {
                            case "0":
                                ErrorMessages += "You have insufficient rights to perform this operation.";
                                break;
                            case "8":
                                ErrorMessages += "Page requested too many times, please retry in a few moments.";
                                break;
                            default:
                                ErrorMessages +=
                                    "Response is not xml.\r\nYou might have insufficient rights to perform this operation.\r\n\r\n" + buffer;
                                break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void GetMessages(Stream stream)
        {
            var nav = new XPathDocument(stream).CreateNavigator();
            XPathNodeIterator messageNodes = nav.Select("/eExact/Messages/Message");
            if (messageNodes.Count > 0)
            {
                ErrorMessages = messageNodes.Current.InnerXml;
            }
            Messages = ErrorMessages;
        }
        #endregion

        #region Private Methods
        private static string GetTimestamp(Stream stream)
        {
            var doc = new XPathDocument(stream);
            var nav = doc.CreateNavigator();
            return nav.Evaluate("string(/eExact/Topics/Topic/@ts_d)") as string;
        }
        private static int GetResultCount(Stream stream)
        {
            var doc = new XPathDocument(stream);
            var nav = doc.CreateNavigator();
            var count = nav.Evaluate("string(/eExact/Topics/Topic/@count)") as string;
            return !string.IsNullOrEmpty(count) ? Convert.ToInt32(count) : 0;
        }
        private static int GetPageSize(Stream stream)
        {
            var doc = new XPathDocument(stream);
            var nav = doc.CreateNavigator();
            var pageSize = nav.Evaluate("string(/eExact/Topics/Topic/@pagesize)") as string;
            return !string.IsNullOrEmpty(pageSize) ? Convert.ToInt32(pageSize) : 0;
        }
        private bool GetErrorMessages(Stream stream)
        {
            bool errorsFound = false;
            var nav = new XPathDocument(stream).CreateNavigator();
            XPathNodeIterator messageNodes = nav.Select("/eExact/Messages/Message");
            if (messageNodes.Count > 0)
            {
                while (messageNodes.MoveNext())
                {
                    nav = messageNodes.Current;
                    int messageType = Convert.ToInt32(nav.GetAttribute("type", string.Empty));
                    if (messageType == (int)MessageType.Warning)
                    {
                        ErrorMessages += "Warning: ";
                    }
                    else
                    {
                        ErrorMessages += "Error: ";
                        errorsFound = true;
                    }

                    var topicName = nav.Evaluate("string(Topic/@node)") as string;
                    var key = nav.Evaluate("string(Topic/Data/@key)") as string;
                    if (string.IsNullOrEmpty(key)) key = nav.Evaluate("string(Topic/Data/@keyAlt)") as string;
                    if (!string.IsNullOrEmpty(topicName) || !string.IsNullOrEmpty(key))
                    {
                        ErrorMessages += string.Format("{0} {1}: {2}\r\n", topicName, key, nav.Evaluate("string(Description)"));
                    }
                }
            }
            return errorsFound;
        }
        private void ReadAttachments(Stream stream, string topic)
        {
            XPathNavigator nav = new XPathDocument(stream).CreateNavigator();
            switch (topic.ToLowerInvariant())
            {
                case "accounts":
                    ReadAttachmentNodes(nav.Select("/eExact/Accounts/Account/Image"));
                    ReadAttachmentNodes(nav.Select("/eExact/Accounts/Account/Contact/Image"));
                    break;
                case "invoices":
                    ReadAttachmentNodes(nav.Select("/eExact/Invoices/Invoice/Document/Attachments/Attachment"));
                    break;
                case "documents":
                case "layouts":
                    ReadAttachmentNodes(nav.Select("/eExact/Documents/Document/Attachments/Attachment"));
                    ReadAttachmentNodes(nav.Select("/eExact/Documents/Document/Images/Image"));
                    break;
                case "items":
                    ReadAttachmentNodes(nav.Select("/eExact/Items/Item/Image"));
                    break;
            }
        }
        private void ReadAttachmentNodes(XPathNodeIterator attachmentNodes)
        {
            try
            {
                if (attachmentNodes == null) return;

                while (attachmentNodes.MoveNext())
                {
                    XPathNavigator nav = attachmentNodes.Current;
                    var name = nav.Evaluate("string(Name)") as string;
                    if (string.IsNullOrEmpty(name)) continue;

                    XPathNavigator navBinData = nav.SelectSingleNode("BinaryData");
                    if (navBinData == null) continue;

                    string file = Path.GetFileName(name);
                    byte[] data = Convert.FromBase64String(navBinData.InnerXml);
                    if (!string.IsNullOrEmpty(file) && data.GetLength(0) > 0)
                    {
                        Attachments.Add(new Attachment { FileName = file, Data = data });
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion
    }
}
