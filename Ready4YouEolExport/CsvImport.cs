using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Drawing;

namespace Ready4YouEolExport
{
    public partial class CsvImport : Form
    {
        DirectoryInfo xmlFolder;
        DirectoryInfo fallOutFolder;
        DirectoryInfo statusLogFolder;
        DirectoryInfo csvBackupFolder;
        DirectoryInfo eolOutputFolder;
        string fallOutFileName = string.Empty;
        Dictionary<int, string> dicUrenStatusLog = new Dictionary<int, string>();
        Dictionary<int, string> dicKostenStatusLog = new Dictionary<int, string>();
        Dictionary<int, string> dicUrenFallout = new Dictionary<int, string>();
        Dictionary<int, string> dicKostenFallout = new Dictionary<int, string>();
        bool werksoortEditMode = false;

        public CsvImport()
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");
            InitializeComponent();
        }
        private void CsvImport_Load(object sender, EventArgs e)
        {
            this.Text += "-" + this.ProductVersion;

            tbWerksoort.Text = "";
            tbVervangingWerksoort.Text = "";
            tbWerksoort.Enabled = false;
            tbVervangingWerksoort.Enabled = false;
            btnAdd.Enabled = true;
            btnEdit.Enabled = false;
            btnDelete.Enabled = false;
            btnSave.Enabled = false;
            LoadWeksoortDetails();

            xmlFolder = new DirectoryInfo("XML");
            if (!xmlFolder.Exists) xmlFolder.Create();

            fallOutFolder = new DirectoryInfo("FallOut");
            if (!fallOutFolder.Exists) fallOutFolder.Create();

            csvBackupFolder = new DirectoryInfo("CsvBackup");
            if (!csvBackupFolder.Exists) csvBackupFolder.Create();

            statusLogFolder = new DirectoryInfo("StatusLog");
            if (!statusLogFolder.Exists) statusLogFolder.Create();

            eolOutputFolder = new DirectoryInfo("EolOutput");
            if (!eolOutputFolder.Exists) eolOutputFolder.Create();

            lblUrenFileName.Text = string.Empty;
            lblKostenFileName.Text = string.Empty;

            ReadDataValues();
        }

        private void BackupCsv(string filename)
        {
            FileInfo fileInfo = new FileInfo(filename);
            fileInfo.CopyTo(csvBackupFolder.FullName + "//" + DateTime.Now.ToString("yyyyMMddHHmmss_") + fileInfo.Name);
        }
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void BtnGenerateXML_Click(object sender, EventArgs e)
        {
            if (IsFactorValuesAndFileDetailsProvided())
            {
                dicUrenStatusLog = new Dictionary<int, string>();
                dicKostenStatusLog = new Dictionary<int, string>();
                dicUrenFallout = new Dictionary<int, string>();
                dicKostenFallout = new Dictionary<int, string>();

                SaveDataValues();
                btnGenerateXML.Enabled = false;

                DataTable dt = CsvToDataTable();

                if (dt.Rows.Count > 0)
                {
                    //System.Threading.Thread.CurrentThread.CurrentUICulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");
                    DataView view = new DataView(dt);
                    DataTable dtJaarWeeknummerKlant = view.ToTable(true, new string[] { "Jaar", "Weeknummer", "Klantnummer" });

                    EolImportExport eolObj = new EolImportExport();
                    int eolCompanyId = Convert.ToInt32(ReadConfig("EOLCompanyId"));
                    string clientId = ReadConfig("ClientId");
                    string clientSecret = ReadConfig("ClientSecret");
                    int lineNo = 0;
                    //for loop per week
                    foreach (DataRow drJaarWeeknummerKlant in dtJaarWeeknummerKlant.Rows)
                    {
                        string xmlFileName = xmlFolder.FullName + "//"
                            + drJaarWeeknummerKlant["Jaar"].ToString()
                            + "_" + drJaarWeeknummerKlant["Weeknummer"].ToString()
                            + "_" + drJaarWeeknummerKlant["Klantnummer"].ToString()
                            + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xml";

                        #region Group Data By Year, week, klant, tankstation, Uitzendkracht, kostenplaats(factors), werksoort, age and uurloon
                        var dtKlantnData = dt.AsEnumerable()
                                 .Where(r => r.Field<string>("Klantnummer") == drJaarWeeknummerKlant["Klantnummer"].ToString()
                                        && r.Field<int>("Weeknummer") == Convert.ToInt32(drJaarWeeknummerKlant["Weeknummer"].ToString())
                                        && r.Field<int>("Jaar") == Convert.ToInt32(drJaarWeeknummerKlant["Jaar"].ToString())
                                        )
                                 .GroupBy(row => new
                                 {
                                     FileType = row.Field<string>("FileType"),
                                     Jaar = row.Field<int>("Jaar"),
                                     Weeknummer = row.Field<int>("Weeknummer"),
                                     Klantnummer = row.Field<string>("Klantnummer"),
                                     Functie = row.Field<string>("Functie"),
                                     CostCenterCode = row.Field<string>("Debiteurnummer"),
                                     Uitzendkracht = row.Field<string>("Uitzendkracht"),
                                     Kostenplaats = row.Field<string>("Kostenplaats"),
                                     Werksoort = row.Field<string>("Werksoort"),
                                     Leeftijd = (row.Field<int?>("Leeftijd") == null) ? 0 : row.Field<int>("Leeftijd"),
                                     Bedrag = (row.Field<decimal?>("Bedrag") == null) ? 0.0m : row.Field<decimal>("Bedrag"),
                                     Uurloon = (row.Field<decimal?>("Uurloon") == null) ? 0.0m : row.Field<decimal>("Uurloon")
                                 })
                                 .Select(g => new
                                 {
                                     g.Key.FileType,
                                     g.Key.Jaar,
                                     g.Key.Weeknummer,
                                     g.Key.Klantnummer,
                                     g.Key.Functie,
                                     g.Key.CostCenterCode,
                                     g.Key.Uitzendkracht,
                                     g.Key.Kostenplaats,
                                     g.Key.Werksoort,
                                     g.Key.Leeftijd,
                                     g.Key.Bedrag,
                                     g.Key.Uurloon,
                                     sumAantal = g.Sum(f => f.Field<decimal>("Aantal")),
                                     lineNos = string.Join(",", g.Select(f => f.Field<int>("LineNo").ToString()))
                                 })
                                 .OrderBy(o => o.Klantnummer)
                                 .ThenBy(o => o.Jaar)
                                 .ThenBy(o => o.Weeknummer)
                                 .ThenBy(o => o.CostCenterCode)
                                 .ThenBy(o => o.Uitzendkracht)
                                 ;
                        #endregion Group Data By Year, week, klant, tankstation, Uitzendkracht, kostenplaats(factors), werksoort, age and uurloon

                        #region XML
                        #region Header
                        XmlDocument doc = new XmlDocument();
                        //doc.CreateXmlDeclaration("1.0", "UTF-8\" standalone=\"yes", "");
                        doc.LoadXml(
                            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                            "<eExact xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='eExact-XML.xsd'>" +
                            "</eExact>");
                        XmlElement root = doc.DocumentElement;
                        #endregion Header

                        #region Invoices
                        XmlElement ndInvoices = doc.CreateElement(string.Empty, "Invoices", string.Empty);
                        {
                            XmlElement ndInvoice = doc.CreateElement(string.Empty, "Invoice", string.Empty);
                            ndInvoice.SetAttribute("type", ReadConfig("InvoiceType"));
                            {
                                #region InvoiceHeader
                                XmlElement ndJournal = doc.CreateElement(string.Empty, "Journal", string.Empty);
                                ndJournal.SetAttribute("code", ReadConfig("InvoiceJournalCode"));
                                ndInvoice.AppendChild(ndJournal);

                                XmlElement ndInvoiceDescription = doc.CreateElement(string.Empty, "Description", string.Empty);
                                //string weekNo = "Week " + IsoWeekNumber(dtpOrderDate.Value).ToString() + " " + IsoWeekYear(dtpOrderDate.Value).ToString();
                                string weekNo = "Week " + drJaarWeeknummerKlant["Weeknummer"].ToString() + " " + drJaarWeeknummerKlant["Jaar"].ToString();
                                XmlText xtInvoiceDescription = doc.CreateTextNode(weekNo);
                                ndInvoiceDescription.AppendChild(xtInvoiceDescription);
                                ndInvoice.AppendChild(ndInvoiceDescription);

                                XmlElement ndOrderedBy = doc.CreateElement(string.Empty, "OrderedBy", string.Empty);
                                ndOrderedBy.SetAttribute("code", drJaarWeeknummerKlant["Klantnummer"].ToString());
                                ndInvoice.AppendChild(ndOrderedBy);

                                XmlElement ndInvoiceTo = doc.CreateElement(string.Empty, "InvoiceTo", string.Empty);
                                ndInvoiceTo.SetAttribute("code", drJaarWeeknummerKlant["Klantnummer"].ToString());
                                ndInvoice.AppendChild(ndInvoiceTo);

                                //XmlElement ndPaymentCondition = doc.CreateElement(string.Empty, "PaymentCondition", string.Empty);
                                //ndPaymentCondition.SetAttribute("code", ReadConfig("PaymentCondition"));
                                //ndInvoice.AppendChild(ndPaymentCondition);
                                #endregion InvoiceHeader

                                lineNo = 0;
                                string oldCostCenterCode = string.Empty;
                                string oldPersonName = string.Empty;

                                foreach (var drKlantnData in dtKlantnData)
                                {
                                    string itemCode = drKlantnData.Werksoort;
                                    decimal itemQuantity = Convert.ToDecimal(drKlantnData.sumAantal.ToString());
                                    string personName = drKlantnData.Uitzendkracht;
                                    string functie = drKlantnData.Functie;
                                    string costCenterCode = drKlantnData.CostCenterCode;
                                    decimal itemUnitPriceValue = 0;

                                    if (string.Compare(itemCode, "Telefoon", true) == 0)
                                    {
                                        itemUnitPriceValue = Convert.ToDecimal(drKlantnData.Bedrag.ToString("##.00"));
                                    }
                                    else if (string.Compare(itemCode, "CONSIGNATIE", true) == 0)
                                    {
                                        itemUnitPriceValue = Convert.ToDecimal(
                                                                (drKlantnData.Bedrag
                                                                    * (string.IsNullOrEmpty(tbConsignatie.Text.Trim()) ? 0 : Convert.ToDecimal(tbConsignatie.Text.Trim()))
                                                                ).ToString("##.00")
                                                            );
                                    }
                                    else
                                    {
                                        itemUnitPriceValue = CalculateUnitPrice(drKlantnData.Kostenplaats, drKlantnData.Uurloon, itemCode, drKlantnData.Leeftijd);
                                    }

                                    decimal itemNetPriceValue = itemUnitPriceValue;

                                    lineNo++;
                                    //generate dummy invoice line to seperate each tankstation and each people with in tank station
                                    if (costCenterCode != oldCostCenterCode || personName != oldPersonName)
                                    {
                                        //Insert Empty row before the new costcenter starts 
                                        if (costCenterCode != oldCostCenterCode && lineNo > 1)
                                        {
                                            #region InvoiceLine
                                            XmlElement ndInvoiceLine = doc.CreateElement(string.Empty, "InvoiceLine", string.Empty);
                                            ndInvoiceLine.SetAttribute("line", lineNo.ToString());
                                            lineNo++;
                                            #region LineItems
                                            {
                                                XmlElement ndItem = doc.CreateElement(string.Empty, "Item", string.Empty);
                                                ndItem.SetAttribute("code", ".");
                                                ndInvoiceLine.AppendChild(ndItem);

                                                XmlElement ndQuantity = doc.CreateElement(string.Empty, "Quantity", string.Empty);
                                                XmlText xtQuantity = doc.CreateTextNode("0");
                                                ndQuantity.AppendChild(xtQuantity);
                                                ndInvoiceLine.AppendChild(ndQuantity);
                                            }
                                            #endregion LineItems
                                            ndInvoice.AppendChild(ndInvoiceLine);
                                            #endregion InvoiceLine
                                        }
                                        //for all 
                                        {
                                            #region InvoiceLine
                                            XmlElement ndInvoiceLine = doc.CreateElement(string.Empty, "InvoiceLine", string.Empty);
                                            ndInvoiceLine.SetAttribute("line", lineNo.ToString());
                                            lineNo++;
                                            #region LineItems
                                            {
                                                XmlElement ndItem = doc.CreateElement(string.Empty, "Item", string.Empty);
                                                ndItem.SetAttribute("code", ".");
                                                ndInvoiceLine.AppendChild(ndItem);

                                                XmlElement ndQuantity = doc.CreateElement(string.Empty, "Quantity", string.Empty);
                                                XmlText xtQuantity = doc.CreateTextNode("0");
                                                ndQuantity.AppendChild(xtQuantity);
                                                ndInvoiceLine.AppendChild(ndQuantity);

                                                if (costCenterCode != oldCostCenterCode)
                                                {
                                                    XmlElement ndCostcenter = doc.CreateElement(string.Empty, "Costcenter", string.Empty);
                                                    ndCostcenter.SetAttribute("code", drKlantnData.CostCenterCode);
                                                    ndInvoiceLine.AppendChild(ndCostcenter);
                                                    oldCostCenterCode = costCenterCode;
                                                }

                                                XmlElement ndNote = doc.CreateElement(string.Empty, "Note", string.Empty);
                                                XmlText xtNote = doc.CreateTextNode(personName + " - Functie: " + functie);
                                                ndNote.AppendChild(xtNote);
                                                ndInvoiceLine.AppendChild(ndNote);
                                                oldPersonName = personName;
                                            }
                                            #endregion LineItems
                                            ndInvoice.AppendChild(ndInvoiceLine);
                                            #endregion InvoiceLine
                                        }
                                    }

                                    //for all entries with working hours or declaraties
                                    {
                                        #region InvoiceLine
                                        XmlElement ndInvoiceLine = doc.CreateElement(string.Empty, "InvoiceLine", string.Empty);
                                        ndInvoiceLine.SetAttribute("line", lineNo.ToString());
                                        #region LineItems
                                        {
                                            XmlElement ndItem = doc.CreateElement(string.Empty, "Item", string.Empty);
                                            ndItem.SetAttribute("code", CheckAndGetReplacementWerksoort(itemCode));
                                            ndInvoiceLine.AppendChild(ndItem);

                                            XmlElement ndQuantity = doc.CreateElement(string.Empty, "Quantity", string.Empty);

                                            //change the culture to en-US to get correct decimal format in XML
                                            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                                            XmlText xtQuantity = doc.CreateTextNode(itemQuantity.ToString());

                                            //change the culture to nl-NL to get NL decimal format in application
                                            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");

                                            ndQuantity.AppendChild(xtQuantity);
                                            ndInvoiceLine.AppendChild(ndQuantity);

                                            //if (drKlantnData.FileType == "Uren")
                                            if (itemCode != "KM" && itemCode != "KM-EG")
                                            {
                                                #region UnitPrice
                                                XmlElement ndUnitPrice = doc.CreateElement(string.Empty, "UnitPrice", string.Empty);

                                                XmlElement ndCurrency = doc.CreateElement(string.Empty, "Currency", string.Empty);
                                                ndCurrency.SetAttribute("code", ReadConfig("CurrencyCode"));
                                                ndUnitPrice.AppendChild(ndCurrency);

                                                XmlElement ndUnitPriceValue = doc.CreateElement(string.Empty, "Value", string.Empty);

                                                //change the culture to en-US to get correct decimal format in XML
                                                System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                                                XmlText xtUnitPriceValue = doc.CreateTextNode(itemUnitPriceValue.ToString());

                                                //change the culture to nl-NL to get NL decimal format in application
                                                System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");

                                                ndUnitPriceValue.AppendChild(xtUnitPriceValue);
                                                ndUnitPrice.AppendChild(ndUnitPriceValue);

                                                XmlElement ndVat = doc.CreateElement(string.Empty, "VAT", string.Empty);
                                                ndVat.SetAttribute("code", ReadConfig("VatCode"));
                                                ndUnitPrice.AppendChild(ndVat);

                                                ndInvoiceLine.AppendChild(ndUnitPrice);
                                                #endregion UnitPrice

                                                #region NetPrice
                                                XmlElement ndNetPrice = doc.CreateElement(string.Empty, "NetPrice", string.Empty);

                                                ndCurrency = doc.CreateElement(string.Empty, "Currency", string.Empty);
                                                ndCurrency.SetAttribute("code", ReadConfig("CurrencyCode"));
                                                ndNetPrice.AppendChild(ndCurrency);

                                                XmlElement ndNetPriceValue = doc.CreateElement(string.Empty, "Value", string.Empty);

                                                //change the culture to en-US to get correct decimal format in XML
                                                System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                                                XmlText xtNetPriceValue = doc.CreateTextNode(itemNetPriceValue.ToString());

                                                //change the culture to nl-NL to get NL decimal format in application
                                                System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");

                                                ndNetPriceValue.AppendChild(xtNetPriceValue);
                                                ndNetPrice.AppendChild(ndNetPriceValue);

                                                ndVat = doc.CreateElement(string.Empty, "VAT", string.Empty);
                                                ndVat.SetAttribute("code", ReadConfig("VatCode"));
                                                ndNetPrice.AppendChild(ndVat);

                                                ndInvoiceLine.AppendChild(ndNetPrice);
                                                #endregion NetPrice
                                            }
                                        }
                                        #endregion LineItems
                                        ndInvoice.AppendChild(ndInvoiceLine);
                                        #endregion InvoiceLine
                                    }
                                }//foreach (DataRow drKlantnData in dtKlantnData.Rows)
                            }
                            ndInvoices.AppendChild(ndInvoice);
                        }
                        root.AppendChild(ndInvoices);
                        #endregion Invoices

                        doc.Save(xmlFileName);
                        //MessageBox.Show("XML Generated\n" + xmlpath);
                        #endregion XML

#if DEBUG
                        var a = 10;
#else
                        if (lineNo > 0) //only when there are Invoice Lines created
                        {
                        #region SendToEOL
                            FileInfo xmlFileInfo = new FileInfo(xmlFileName);
                            string result = eolObj.ImportFiles(xmlFileInfo.FullName, "Invoices", "https://start.exactonline.nl/", eolCompanyId, clientId, clientSecret);

                            EolOutput("EolResult_" + xmlFileInfo.Name, result);

                        #region Read The result
                            string resultType = string.Empty;

                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(result);

                            string xpath = "eExact/Messages";
                            var nodes = xmlDoc.SelectNodes(xpath);

                            foreach (XmlNode childrenNode in nodes)
                            {
                                resultType = childrenNode.SelectSingleNode("//Message").Attributes["type"].Value;
                                if (!string.IsNullOrEmpty(resultType))
                                {
                                    break;
                                }
                            }
                        #endregion Read The result

                            if (resultType == "2") //success
                            {
                                //write success log
                                string invoiceCode = string.Empty;
                                xpath = "eExact/Messages/Message/Topic/Data";
                                nodes = xmlDoc.SelectNodes(xpath);

                                foreach (XmlNode childrenNode in nodes)
                                {
                                    if (childrenNode.Attributes.GetNamedItem("key") != null)
                                    {
                                        invoiceCode = childrenNode.Attributes["key"].Value;
                                        if (!string.IsNullOrEmpty(invoiceCode))
                                        {
                                            break;
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(lblUrenFileName.Text))
                                {
                                    var dicSuccessRows = dt.AsEnumerable()
                                     .Where(r => r.Field<string>("Klantnummer") == drJaarWeeknummerKlant["Klantnummer"].ToString()
                                            && r.Field<int>("Weeknummer") == Convert.ToInt32(drJaarWeeknummerKlant["Weeknummer"].ToString())
                                            && r.Field<int>("Jaar") == Convert.ToInt32(drJaarWeeknummerKlant["Jaar"].ToString())
                                            && r.Field<int?>("Leeftijd") != null
                                            )
                                     .ToDictionary(r => r.Field<int>("LineNo"), r => "Success : Invoice Code : " + invoiceCode);

                                    foreach (KeyValuePair<int, string> dicSuccessRow in dicSuccessRows)
                                    {
                                        dicUrenStatusLog.Add(dicSuccessRow.Key, dicSuccessRow.Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(lblKostenFileName.Text))
                                {
                                    var dicSuccessRows = dt.AsEnumerable()
                                     .Where(r => r.Field<string>("Klantnummer") == drJaarWeeknummerKlant["Klantnummer"].ToString()
                                            && r.Field<int>("Weeknummer") == Convert.ToInt32(drJaarWeeknummerKlant["Weeknummer"].ToString())
                                            && r.Field<int>("Jaar") == Convert.ToInt32(drJaarWeeknummerKlant["Jaar"].ToString())
                                            && r.Field<int?>("Leeftijd") == null
                                            )
                                     .ToDictionary(r => r.Field<int>("LineNo"), r => "Success : Invoice Code : " + invoiceCode);

                                    foreach (KeyValuePair<int, string> dicSuccessRow in dicSuccessRows)
                                    {
                                        dicKostenStatusLog.Add(dicSuccessRow.Key, dicSuccessRow.Value);
                                    }
                                }

                                DirectoryInfo directoryInfo = new DirectoryInfo(xmlFolder + "//Success");
                                if (!directoryInfo.Exists) directoryInfo.Create();
                                xmlFileInfo.MoveTo(xmlFolder.FullName + "//Success//" + xmlFileInfo.Name);
                                //xmlFileInfo.Delete();
                            }
                            else if (resultType == "0") //failure
                            {
                                string errorMsg = string.Empty;
                                xpath = "eExact/Messages/Message/Description";
                                nodes = xmlDoc.SelectNodes(xpath);

                                foreach (XmlNode childrenNode in nodes)
                                {
                                    errorMsg = childrenNode.SelectSingleNode("//Description").InnerText;
                                    if (!string.IsNullOrEmpty(errorMsg))
                                    {
                                        break;
                                    }
                                }
                                //create fallout csv
                                //write log
                                CsvToFalloutFile("All", "All", drJaarWeeknummerKlant["Klantnummer"].ToString(), drJaarWeeknummerKlant["Jaar"].ToString(), drJaarWeeknummerKlant["Weeknummer"].ToString(), errorMsg);
                                DirectoryInfo directoryInfo = new DirectoryInfo(xmlFolder + "//Failed");
                                if (!directoryInfo.Exists) directoryInfo.Create();
                                xmlFileInfo.MoveTo(xmlFolder.FullName + "//Failed//" + xmlFileInfo.Name);
                                //xmlFileInfo.Delete();
                            }

                        #endregion SendToEOL
                        }
#endif
                    }
                }
                else
                {
                    MessageBox.Show("No Data to Export to EOL");
                }

                DataStatusLog();
                FallOut();
                MessageBox.Show("Done! Please check the Status Log files for details");
                lblUrenFileName.Text = string.Empty;
                lblKostenFileName.Text = string.Empty;
                btnGenerateXML.Enabled = true;
            }
        }
        private void BtnKostenFileUpload_Click(object sender, EventArgs e)
        {
            lblKostenFileName.Text = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "csv files|*.csv";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //The path of specified file
                    lblKostenFileName.Text = openFileDialog.FileName;
                }
            }
        }
        private void BtnUrenFileUpload_Click(object sender, EventArgs e)
        {
            lblUrenFileName.Text = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "csv files|*.csv";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //The path of specified file
                    lblUrenFileName.Text = openFileDialog.FileName;
                }
            }
        }
        private decimal CalculateUnitPrice(string kostenplaats, decimal uurloon, string werksoort, int? age = 999)
        {
            decimal unitPrice = 0;
            decimal youth = 0;
            if (age < 21 && uurloon < 10.2M)
                youth = string.IsNullOrEmpty(tbJegud.Text.Trim()) ? 0 : Convert.ToDecimal(tbJegud.Text.Trim());

            switch (kostenplaats)
            {
                case "Werving":
                    unitPrice = (uurloon * (string.IsNullOrEmpty(tbWerving.Text.Trim()) ? 0 : Convert.ToDecimal(tbWerving.Text.Trim()))) + youth;
                    break;
                case "Opleiding":
                    unitPrice = (uurloon * (string.IsNullOrEmpty(tbOpleiding.Text.Trim()) ? 0 : Convert.ToDecimal(tbOpleiding.Text.Trim()))) + youth;
                    break;
                case "Selectie":
                    unitPrice = (uurloon * (string.IsNullOrEmpty(tbSelectie.Text.Trim()) ? 0 : Convert.ToDecimal(tbSelectie.Text.Trim()))) + youth;
                    break;
                case "Spoed":
                    unitPrice = string.IsNullOrEmpty(tbSpoed.Text.Trim()) ? 0 : Convert.ToDecimal(tbSpoed.Text.Trim());
                    if (werksoort.All(char.IsDigit))
                    {
                        werksoort = CheckAndGetReplacementWerksoort(werksoort);
                        unitPrice *= Convert.ToDecimal(werksoort) / 100M;
                    }
                    break;
                case "Uitzenden":
                    unitPrice = string.IsNullOrEmpty(tbUitzenden.Text.Trim()) ? 0 : Convert.ToDecimal(tbUitzenden.Text.Trim());
                    if (werksoort.All(char.IsDigit))
                    {
                        werksoort = CheckAndGetReplacementWerksoort(werksoort);
                        unitPrice *= Convert.ToDecimal(werksoort) / 100M;
                    }
                    break;
            }

            //round on 2 decimals
            return Convert.ToDecimal(unitPrice.ToString("##.00"));
        }
        private string CheckAndGetReplacementWerksoort(string werksoort)
        {
            foreach (DataGridViewRow dgvr in dgWerksoortInfo.Rows)
            {
                if (dgvr.Cells["Werksoort"].Value != null && dgvr.Cells["Werksoort"].Value.Equals(werksoort)) return dgvr.Cells["Vervanging Werksoort"].Value.ToString();
            }
            return werksoort;
        }
        private DataTable CsvToDataTable()
        {
            DataTable dt = new DataTable();
            //add the columns to the datatable
            dt.Columns.Add("Klantnummer", typeof(string));
            dt.Columns.Add("Weeknummer", typeof(int));
            dt.Columns.Add("Jaar", typeof(int));
            dt.Columns.Add("Debiteurnummer", typeof(string));
            dt.Columns.Add("Tankstation", typeof(string));
            dt.Columns.Add("Uitzendkracht", typeof(string));
            dt.Columns.Add("Leeftijd", typeof(int));
            dt.Columns.Add("Functie", typeof(string));
            dt.Columns.Add("Werksoort", typeof(string));
            dt.Columns.Add("Kostenplaats", typeof(string));
            dt.Columns.Add("Datum", typeof(DateTime));
            dt.Columns.Add("Aantal", typeof(decimal));
            dt.Columns.Add("Bedrag", typeof(decimal));
            dt.Columns.Add("Uurloon", typeof(decimal));
            dt.Columns.Add("LineNo", typeof(int));
            dt.Columns.Add("FileName", typeof(string));
            dt.Columns.Add("FileType", typeof(string)); //Uren or Kosten

            DataTable dtUren = CsvUrenToDataTable();
            DataTable dtKosten = CsvKostenToDataTable();

            foreach (DataRow dr in dtUren.Rows)
            {
                dt.ImportRow(dr);
            }

            foreach (DataRow dr in dtKosten.Rows)
            {
                dt.ImportRow(dr);
            }

            return dt;
        }
        private void CsvToFalloutFile(string lineNos, string fileType, string klantnummer, string jaar, string weeknummer, string errorMsg)
        {
            if (fileType == "All" || fileType == "Uren")
            {
                if (!string.IsNullOrEmpty(lblUrenFileName.Text))
                {
                    using (StreamReader sr = new StreamReader(lblUrenFileName.Text, System.Text.Encoding.Default))
                    {
                        //1st row... column headers //skip
                        sr.ReadLine();
                        int lineNo = 1;
                        var csvDataRow = sr.ReadLine();
                        while (csvDataRow != null)
                        {
                            //runs until string reader returns null and adds rows to fallout 
                            lineNo++;
                            if (lineNos == "All" || (lineNos.Split(',').ToList().Contains(lineNo.ToString())))
                            {
                                var rows = csvDataRow.Split(';').ToList();
                                if (rows.Count == 13)
                                {
                                    if (rows[0] == klantnummer && rows[1] == weeknummer && rows[2] == jaar)
                                    {
                                        if (!dicUrenFallout.Keys.Contains(lineNo))
                                            dicUrenFallout.Add(lineNo, csvDataRow);
                                        if (!dicUrenStatusLog.Keys.Contains(lineNo))
                                            dicUrenStatusLog.Add(lineNo, "Failed :" + errorMsg);
                                    }
                                }
                            }
                            csvDataRow = sr.ReadLine();
                        }
                    }
                }
            }

            if (fileType == "All" || fileType == "Kosten")
            {
                if (!string.IsNullOrEmpty(lblKostenFileName.Text))
                {
                    using (StreamReader sr = new StreamReader(lblKostenFileName.Text, System.Text.Encoding.Default))
                    {
                        //1st row... column headers //skip
                        sr.ReadLine();
                        int lineNo = 1;
                        var csvDataRow = sr.ReadLine();
                        while (csvDataRow != null)
                        {
                            //runs until string reader returns null and adds rows to fallout
                            lineNo++;
                            if (lineNos == "All" || (lineNos.Split(',').ToList().Contains(lineNo.ToString())))
                            {
                                var rows = csvDataRow.Split(';').ToList();
                                if (rows.Count == 12)
                                {
                                    if (rows[0] == klantnummer && rows[1] == weeknummer && rows[2] == jaar)
                                    {
                                        if (!dicKostenFallout.Keys.Contains(lineNo))
                                            dicKostenFallout.Add(lineNo, csvDataRow);
                                        if (!dicKostenStatusLog.Keys.Contains(lineNo))
                                            dicKostenStatusLog.Add(lineNo, "Failed :" + errorMsg);
                                    }
                                }
                            }
                            csvDataRow = sr.ReadLine();
                        }
                    }
                }
            }
        }
        private DataTable CsvKostenToDataTable()
        {
            DataTable dtKosten = new DataTable();
            //add the columns to the datatable
            dtKosten.Columns.Add("Klantnummer", typeof(string)); //0
            dtKosten.Columns.Add("Weeknummer", typeof(int)); //1
            dtKosten.Columns.Add("Jaar", typeof(int)); //2
            dtKosten.Columns.Add("Debiteurnummer", typeof(string)); //3
            dtKosten.Columns.Add("Tankstation", typeof(string));//4
            dtKosten.Columns.Add("Uitzendkracht", typeof(string));//5
            dtKosten.Columns.Add("Functie", typeof(string));//6
            dtKosten.Columns.Add("Werksoort", typeof(string));//7
            dtKosten.Columns.Add("Kostenplaats", typeof(string));//8
            dtKosten.Columns.Add("Datum", typeof(DateTime));//9
            dtKosten.Columns.Add("Aantal", typeof(decimal));//10
            dtKosten.Columns.Add("Bedrag", typeof(decimal));//11
            dtKosten.Columns.Add("Uurloon", typeof(decimal));//12
            dtKosten.Columns.Add("LineNo", typeof(int));//13
            dtKosten.Columns.Add("FileName", typeof(string));//14
            dtKosten.Columns.Add("FileType", typeof(string)); //15--Uren or Kosten

            if (!string.IsNullOrEmpty(lblKostenFileName.Text))
            {
                //backup the original csv
                BackupCsv(lblKostenFileName.Text);

                //catch the klantnummer, week and year for faulty records
                List<string> fallOutLines = new List<string>();

                //check for faulty records to exclude the complete set
                using (StreamReader sr = new StreamReader(lblKostenFileName.Text, System.Text.Encoding.Default))
                {
                    int lineNo = 1;
                    //1st row... column headers //skip
                    sr.ReadLine();

                    var csvDataRow = sr.ReadLine();

                    //runs until string reader returns null 
                    while (csvDataRow != null)
                    {
                        lineNo++;
                        var rows = csvDataRow.Split(';').ToList().Select(t => t.Trim()).ToList();
                        if (rows.Count != 13
                            //Klantnummer
                            || string.IsNullOrEmpty(rows[0])
                            //Weeknummer
                            || string.IsNullOrEmpty(rows[1])
                            //Jaar
                            || string.IsNullOrEmpty(rows[2])
                            //Debiteurnummer
                            || string.IsNullOrEmpty(rows[3])
                            //Tankstation
                            || string.IsNullOrEmpty(rows[4])
                            //Uitzendkracht
                            || string.IsNullOrEmpty(rows[5])
                            //Functie
                            //|| string.IsNullOrEmpty( rows[6])
                            //Werksoort
                            || string.IsNullOrEmpty(rows[7])
                            //Kostenplaats
                            //|| string.IsNullOrEmpty( rows[8])
                            //Datum
                            || string.IsNullOrEmpty(rows[9])
                            //Aantal
                            //|| string.IsNullOrEmpty(rows[10])
                            //Bedrag
                            //|| string.IsNullOrEmpty( rows[11])
                            //Uurloon
                            //|| string.IsNullOrEmpty( rows[12])
                            )
                        {
                            //invalid entry
                            if (!fallOutLines.Contains(rows[0] + "," + rows[1] + "," + rows[2]))
                                fallOutLines.Add(rows[0] + "," + rows[1] + "," + rows[2]);
                        }
                        else if (string.IsNullOrEmpty(rows[10]) || Convert.ToDecimal(rows[10]) == 0)
                        {
                            //Empty Aantal
                            dicKostenStatusLog.Add(lineNo, "Aantal : 0");
                        }
                        csvDataRow = sr.ReadLine();
                    }
                }

                using (StreamReader sr = new StreamReader(lblKostenFileName.Text, System.Text.Encoding.Default))
                {
                    int lineNo = 1;
                    var csvDataRow = sr.ReadLine();
                    //1st row... column headers
                    dicKostenFallout.Add(lineNo, csvDataRow);
                    csvDataRow = sr.ReadLine();

                    //runs until string reader returns null and adds rows to dt 
                    while (csvDataRow != null)
                    {
                        lineNo++;
                        var rows = csvDataRow.Split(';').ToList().Select(t => t.Trim()).ToList();

                        if (fallOutLines.Contains(rows[0] + "," + rows[1] + "," + rows[2]))
                        {
                            if (rows.Count != 13
                            //Klantnummer
                            || string.IsNullOrEmpty(rows[0])
                            //Weeknummer
                            || string.IsNullOrEmpty(rows[1])
                            //Jaar
                            || string.IsNullOrEmpty(rows[2])
                            //Debiteurnummer
                            || string.IsNullOrEmpty(rows[3])
                            //Tankstation
                            || string.IsNullOrEmpty(rows[4])
                            //Uitzendkracht
                            || string.IsNullOrEmpty(rows[5])
                            //Functie
                            //|| string.IsNullOrEmpty( rows[6])
                            //Werksoort
                            || string.IsNullOrEmpty(rows[7])
                            //Kostenplaats
                            //|| string.IsNullOrEmpty( rows[8])
                            //Datum
                            || string.IsNullOrEmpty(rows[9])
                            //Aantal
                            //|| string.IsNullOrEmpty(rows[10])
                            //Bedrag
                            //|| string.IsNullOrEmpty( rows[11])
                            //Uurloon
                            //|| string.IsNullOrEmpty( rows[12])
                            )
                            {
                                dicKostenStatusLog.Add(lineNo, "Error : Invalid Entry");
                            }
                            else
                            {
                                dicKostenStatusLog.Add(lineNo, "Other line in this set has problem(s)");
                            }
                            dicKostenFallout.Add(lineNo, csvDataRow);
                        }
                        else if (!string.IsNullOrEmpty(rows[10]) && Convert.ToDecimal(rows[10]) > 0) //not in fallout list & Aantal > 0
                        {
                            rows[7] = rows[7].Split(',')[0]; //remove anything after ',' with comma

                            if (string.Compare(rows[7], "Telefoon", true) == 0) //For Telefoon the aantal is always 1
                            {
                                rows[10] = "1";
                            }
                            else
                            {
                                rows[10] = Convert.ToInt32(Convert.ToDecimal(rows[10])).ToString(); //KM count should be rounded to INT directly after reading from the csv
                            }

                            rows[11] = Convert.ToDecimal(Regex.Match(rows[11], @"\d*\.*\d*\,*\d+").Value).ToString();
                            if (rows[12] == "") rows[12] = "0";

                            rows.Add(lineNo.ToString());
                            rows.Add(lblKostenFileName.Text);
                            rows.Add("Kosten");
                            dtKosten.Rows.Add(rows.ToArray());
                        }
                        csvDataRow = sr.ReadLine();
                    }
                }
            }
            return dtKosten;
        }
        private DataTable CsvUrenToDataTable()
        {
            DataTable dtUren = new DataTable();
            //add the columns to the datatable
            dtUren.Columns.Add("Klantnummer", typeof(string)); //0
            dtUren.Columns.Add("Weeknummer", typeof(int)); //1
            dtUren.Columns.Add("Jaar", typeof(int)); //2
            dtUren.Columns.Add("Debiteurnummer", typeof(string)); //3
            dtUren.Columns.Add("Tankstation", typeof(string));//4
            dtUren.Columns.Add("Uitzendkracht", typeof(string));//5
            dtUren.Columns.Add("Leeftijd", typeof(int));//6
            dtUren.Columns.Add("Functie", typeof(string));//7
            dtUren.Columns.Add("Werksoort", typeof(string));//8
            dtUren.Columns.Add("Kostenplaats", typeof(string));//9
            dtUren.Columns.Add("Datum", typeof(DateTime));//10
            dtUren.Columns.Add("Aantal", typeof(decimal));//11
            dtUren.Columns.Add("Uurloon", typeof(decimal));//12
            dtUren.Columns.Add("LineNo", typeof(int));//13
            dtUren.Columns.Add("FileName", typeof(string));//14
            dtUren.Columns.Add("FileType", typeof(string)); //15--Uren or Kosten

            if (!string.IsNullOrEmpty(lblUrenFileName.Text))
            {
                //backup the original csv
                BackupCsv(lblUrenFileName.Text);

                //catch the klantnummer, week and year for faulty records
                List<string> fallOutLines = new List<string>();

                //check for faulty records to exclude the complete set
                using (StreamReader sr = new StreamReader(lblUrenFileName.Text, System.Text.Encoding.Default))
                {
                    int lineNo = 1;
                    //1st row... column headers //skip
                    sr.ReadLine();

                    var csvDataRow = sr.ReadLine();

                    //runs until string reader returns null 
                    while (csvDataRow != null)
                    {
                        lineNo++;
                        var rows = csvDataRow.Split(';').ToList().Select(t => t.Trim()).ToList();

                        if (rows.Count != 13 || rows.Contains(string.Empty))
                        {
                            //invalid entry
                            if (!fallOutLines.Contains(rows[0] + "," + rows[1] + "," + rows[2]))
                                fallOutLines.Add(rows[0] + "," + rows[1] + "," + rows[2]);
                        }
                        else if (Convert.ToDecimal(rows[11]) == 0)
                        {
                            //Empty Aantal
                            dicUrenStatusLog.Add(lineNo, "Aantal : 0");
                        }
                        csvDataRow = sr.ReadLine();
                    }
                }

                using (StreamReader sr = new StreamReader(lblUrenFileName.Text, System.Text.Encoding.Default))
                {
                    int lineNo = 1;
                    var csvDataRow = sr.ReadLine();
                    //1st row... column headers
                    dicUrenFallout.Add(lineNo, csvDataRow);
                    csvDataRow = sr.ReadLine();
                    //runs until string reader returns null 
                    while (csvDataRow != null)
                    {
                        lineNo++;
                        var rows = csvDataRow.Split(';').ToList().Select(t => t.Trim()).ToList();
                        //catch all the related records of the fallout records
                        if (fallOutLines.Contains(rows[0] + "," + rows[1] + "," + rows[2]))
                        {
                            if (rows.Count != 13 || rows.Contains(string.Empty))
                            {
                                //invalid entry
                                dicUrenStatusLog.Add(lineNo, "Error : Invalid Entry");
                            }
                            else
                            {
                                dicUrenStatusLog.Add(lineNo, "Other line in this set has problem(s)");
                            }
                            dicUrenFallout.Add(lineNo, csvDataRow);
                        }
                        else if (Convert.ToDecimal(rows[11]) > 0) //not in fallout list & Aantal > 0
                        {
                            rows[8] = rows[8].Split(',')[0]; //remove anything after ',' with comma

                            //if the Werksoort value is numeric, convert the uurloon value to the percentage value mentioned in Werksoort.
                            //Hourly rates:- take value from the cvs and round on 2 decimals
                            //Hourly rates:- multiply by the percentage and round again on 2 decimals
                            rows[11] = Convert.ToDecimal(rows[11]).ToString("##.00");

                            if (!string.IsNullOrEmpty(rows[8]) && rows[8].All(char.IsDigit))
                            {
                                rows[12] = Convert.ToDecimal(
                                                    Convert.ToDecimal(Convert.ToDecimal(Regex.Match(rows[12], @"\d*\.*\d*\,*\d+").Value).ToString("##.00"))
                                                    * Convert.ToDecimal(CheckAndGetReplacementWerksoort(rows[8])) / 100M
                                                    ).ToString("##.00");
                            }
                            else
                            {
                                rows[12] = Convert.ToDecimal(Regex.Match(rows[12], @"\d*\.*\d*\,*\d+").Value).ToString("##.00");
                            }


                            rows.Add(lineNo.ToString());
                            rows.Add(lblUrenFileName.Text);
                            rows.Add("Uren");
                            dtUren.Rows.Add(rows.ToArray());
                        }
                        csvDataRow = sr.ReadLine();
                    }
                }
            }
            return dtUren;
        }
        public void DataStatusLog()
        {
            string statusLogFileName = "_statuslog" + DateTime.Now.ToString("_yyyyMMddHHmmss") + ".txt";
            if (dicUrenStatusLog.Count > 0)
            {
                FileInfo urenFileInfo = new FileInfo(lblUrenFileName.Text);
                FileInfo logfileInfo = new FileInfo(statusLogFolder + "//" + urenFileInfo.Name);
                using (var writer = File.AppendText(logfileInfo.FullName.Replace(".csv", statusLogFileName)))
                {
                    foreach (KeyValuePair<int, string> kvLog in dicUrenStatusLog.OrderBy(k => k.Key))
                    {
                        writer.WriteLine("Row: " + kvLog.Key.ToString() + " : " + kvLog.Value);
                    }
                    writer.Flush();
                    writer.Close();
                }
            }
            if (dicKostenStatusLog.Count > 0)
            {
                FileInfo kostenFileInfo = new FileInfo(lblKostenFileName.Text);
                FileInfo logfileInfo = new FileInfo(statusLogFolder + "//" + kostenFileInfo.Name);
                using (var writer = File.AppendText(logfileInfo.FullName.Replace(".csv", statusLogFileName)))
                {
                    foreach (KeyValuePair<int, string> kvLog in dicKostenStatusLog.OrderBy(k => k.Key))
                    {
                        writer.WriteLine("Row: " + kvLog.Key.ToString() + " : " + kvLog.Value);
                    }
                    writer.Flush();
                    writer.Close();
                }
            }
        }
        public void EolOutput(string fileName, string result)
        {
            FileInfo eolOutputfileInfo = new FileInfo(eolOutputFolder + "//" + fileName);
            using (var writer = File.CreateText(eolOutputfileInfo.FullName))
            {
                writer.Write(result);
                writer.Flush();
                writer.Close();
            }
        }
        public void FallOut()
        {
            string fallOutFileName = "_fallOut" + DateTime.Now.ToString("_yyyyMMddHHmmss") + ".csv";
            if (dicUrenFallout.Count > 1)
            {
                FileInfo urenFileInfo = new FileInfo(lblUrenFileName.Text);
                FileInfo fallOutFileInfo = new FileInfo(fallOutFolder + "//" + urenFileInfo.Name);
                using (var writer = File.AppendText(fallOutFileInfo.FullName.Replace(".csv", fallOutFileName)))
                {
                    foreach (KeyValuePair<int, string> kvLog in dicUrenFallout.OrderBy(k => k.Key))
                    {
                        writer.WriteLine(kvLog.Value);
                    }
                    writer.Flush();
                    writer.Close();
                }
            }
            if (dicKostenFallout.Count > 1)
            {
                FileInfo kostenFileInfo = new FileInfo(lblKostenFileName.Text);
                FileInfo fallOutFileInfo = new FileInfo(fallOutFolder + "//" + kostenFileInfo.Name);
                using (var writer = File.AppendText(fallOutFileInfo.FullName.Replace(".csv", fallOutFileName)))
                {
                    foreach (KeyValuePair<int, string> kvLog in dicKostenFallout.OrderBy(k => k.Key))
                    {
                        writer.WriteLine(kvLog.Value);
                    }
                    writer.Flush();
                    writer.Close();
                }
            }
        }
        private bool IsFactorValuesAndFileDetailsProvided()
        {
            if (string.IsNullOrEmpty(tbWerving.Text) || Convert.ToDecimal(tbWerving.Text) == 0)
            {
                MessageBox.Show("Please fill Werving factor value");
                tbWerving.Focus();
                return false;
            }
            else if (string.IsNullOrEmpty(tbOpleiding.Text) || Convert.ToDecimal(tbOpleiding.Text) == 0)
            {
                MessageBox.Show("Please fill Opleiding factor value");
                tbOpleiding.Focus();
                return false;
            }
            else if (string.IsNullOrEmpty(tbJegud.Text) || Convert.ToDecimal(tbJegud.Text) == 0)
            {
                MessageBox.Show("Please fill Jegud factor value");
                tbJegud.Focus();
                return false;
            }
            else if (string.IsNullOrEmpty(tbSpoed.Text) || Convert.ToDecimal(tbSpoed.Text) == 0)
            {
                MessageBox.Show("Please fill Spoed factor value");
                tbSpoed.Focus();
                return false;
            }
            else if (string.IsNullOrEmpty(tbUitzenden.Text) || Convert.ToDecimal(tbUitzenden.Text) == 0)
            {
                MessageBox.Show("Please fill Uitzenden factor value");
                tbUitzenden.Focus();
                return false;
            }
            else if (string.IsNullOrEmpty(tbSelectie.Text) || Convert.ToDecimal(tbSelectie.Text) == 0)
            {
                MessageBox.Show("Please fill Selectie factor value");
                tbSelectie.Focus();
                return false;
            }
            else if (string.IsNullOrEmpty(lblUrenFileName.Text) && string.IsNullOrEmpty(lblKostenFileName.Text))
            {
                MessageBox.Show("Please Select either Uren or Kosten file");
                btnUrenFileUpload.Focus();
                return false;
            }
            else if (!string.IsNullOrEmpty(lblUrenFileName.Text) && !lblUrenFileName.Text.EndsWith(".csv"))
            {
                MessageBox.Show("Please Select valid Uren file");
                btnUrenFileUpload.Focus();
                return false;
            }
            else if (!string.IsNullOrEmpty(lblKostenFileName.Text) && !lblKostenFileName.Text.EndsWith(".csv"))
            {
                MessageBox.Show("Please Select valid Kosten file");
                btnKostenFileUpload.Focus();
                return false;
            }
            else
                return true;
        }
        private string ReadConfig(string value)
        {
            string strReturn;
            AppSettingsReader appConfig = new AppSettingsReader();
            try
            {
                strReturn = appConfig.GetValue(value, value.GetType()).ToString();
            }
            catch (Exception ex)
            {
                using (var writer = File.AppendText("ErrorLog.txt"))
                {
                    writer.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss:ffffff") + ": " + ex.Message);
                    writer.Flush();
                    writer.Close();
                }
                strReturn = "";
            }
            return strReturn;
        }
        private void ReadDataValues()
        {
            var datFile = new FileInfo("Ready4YouEolExport.dat");
            if (datFile.Exists)
            {
                using (StreamReader sr = new StreamReader("Ready4YouEolExport.dat", System.Text.Encoding.Default))
                {
                    var myStringRow = sr.ReadLine();
                    while (myStringRow != null)
                    {
                        //runs until string reader returns null and adds rows to dt 
                        var rows = myStringRow.Split(':');
                        switch (rows[0])
                        {
                            case "Werving": tbWerving.Text = rows[1].ToString(); break;
                            case "Opleiding": tbOpleiding.Text = rows[1].ToString(); break;
                            case "Jeugd": tbJegud.Text = rows[1].ToString(); break;
                            case "Spoed": tbSpoed.Text = rows[1].ToString(); break;
                            case "Uitzenden": tbUitzenden.Text = rows[1].ToString(); break;
                            case "Selectie": tbSelectie.Text = rows[1].ToString(); break;
                            case "Consignatie": tbConsignatie.Text = rows[1].ToString(); break;
                        }
                        myStringRow = sr.ReadLine();
                    }
                }
            }
            else
            {
                //values in nl-NL format
                tbUitzenden.Text = "27,95";
                tbSpoed.Text = "29,95";
                tbJegud.Text = "1,0";
            }
        }
        private void SaveDataValues()
        {
            using (var writer = File.CreateText("Ready4YouEolExport.dat"))
            {
                string data = "Werving:" + tbWerving.Text +
                              "\nOpleiding:" + tbOpleiding.Text +
                              "\nJeugd:" + tbJegud.Text +
                              "\nSpoed:" + tbSpoed.Text +
                              "\nUitzenden:" + tbUitzenden.Text +
                              "\nSelectie:" + tbSelectie.Text +
                              "\nConsignatie:" + tbConsignatie.Text;
                writer.WriteLine(data);
                writer.Flush();
                writer.Close();
            }
        }

        #region Werksoort replacement
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            tbWerksoort.Text = "";
            tbVervangingWerksoort.Text = "";
            tbWerksoort.Enabled = true;
            tbVervangingWerksoort.Enabled = true;
            btnAdd.Enabled = false;
            btnEdit.Enabled = false;
            btnDelete.Enabled = false;
            btnSave.Enabled = true;
            werksoortEditMode = false;
            tbWerksoort.Focus();
            dgWerksoortInfo.Enabled = false;
        }
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            ResetWerksoortPanel();
        }
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            var msg = MessageBox.Show("Do you want to delete the Werksoort " + tbWerksoort.Text + " ==> " + tbVervangingWerksoort.Text + " ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (msg.ToString() == "Yes")
            {
                String[] werksoortData = File.ReadAllLines(@"ReplacementWeksoortInfo.dat");
                String[] werksoortDataNew = new string[werksoortData.Length - 1];
                int newIndex = 0;
                for (int i = 0; i < werksoortData.Length; i++)
                {
                    var rows = werksoortData[i].Split(';');
                    if (!rows[0].Equals(dgWerksoortInfo.CurrentRow.Cells[0].Value.ToString()))
                    {
                        werksoortDataNew[newIndex++] = werksoortData[i];
                    }
                }
                File.WriteAllLines(@"ReplacementWeksoortInfo.dat", werksoortDataNew);

                ResetWerksoortPanel();
            }
        }
        private void BtnEdit_Click(object sender, EventArgs e)
        {
            btnAdd.Enabled = false;
            btnEdit.Enabled = false;
            btnDelete.Enabled = false;
            tbWerksoort.Enabled = true;
            tbVervangingWerksoort.Enabled = true;
            btnSave.Enabled = true;
            werksoortEditMode = true;
            tbWerksoort.Focus();
            dgWerksoortInfo.Enabled = false;
        }
        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbWerksoort.Text.Trim()))
            {
                MessageBox.Show("Please fill Werksoort value");
                tbWerksoort.Focus();
            }
            else if (string.IsNullOrEmpty(tbVervangingWerksoort.Text.Trim()))
            {
                MessageBox.Show("Please fill Vervanging Werksoort value");
                tbVervangingWerksoort.Focus();
            }
            else if (CheckWerksoortExists()) //check 
            {
                MessageBox.Show("Replacement value for this Werksoort already exists");
                tbWerksoort.Focus();
            }
            else
            {
                if (werksoortEditMode)
                {
                    String[] werksoortData = File.ReadAllLines(@"ReplacementWeksoortInfo.dat");
                    int totalLines = werksoortData.Length;
                    for (int i = 0; i < totalLines; i++)
                    {
                        var rows = werksoortData[i].Split(';');
                        if (rows[0].Equals(dgWerksoortInfo.CurrentRow.Cells[0].Value.ToString()))
                        {
                            werksoortData[i] = tbWerksoort.Text.Trim() + ";" + tbVervangingWerksoort.Text.Trim();
                        }
                    }
                    File.WriteAllLines(@"ReplacementWeksoortInfo.dat", werksoortData);
                }
                else //new record mode
                {
                    using (var writer = File.AppendText("ReplacementWeksoortInfo.dat"))
                    {
                        writer.WriteLine(tbWerksoort.Text.Trim() + ";" + tbVervangingWerksoort.Text.Trim());
                        writer.Flush();
                        writer.Close();
                    }
                }

                ResetWerksoortPanel();
            }
        }
        private bool CheckWerksoortExists()
        {
            foreach (DataGridViewRow dgvr in dgWerksoortInfo.Rows)
            {
                if (dgvr.Cells["Werksoort"].Value != null
                    && ((werksoortEditMode && dgWerksoortInfo.SelectedRows.Count > 0 && dgvr != dgWerksoortInfo.CurrentRow) || (!werksoortEditMode))
                    && dgvr.Cells["Werksoort"].Value.Equals(tbWerksoort.Text.Trim())
                    ) return true;
            }
            return false;
        }
        private void DgWerksoortInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                tbWerksoort.Text = dgWerksoortInfo.Rows[e.RowIndex].Cells["Werksoort"].Value.ToString();
                tbVervangingWerksoort.Text = dgWerksoortInfo.Rows[e.RowIndex].Cells["Vervanging Werksoort"].Value.ToString();
            }
        }
        private void DgWerksoortInfo_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                tbWerksoort.Text = dgWerksoortInfo.Rows[e.RowIndex].Cells["Werksoort"].Value.ToString();
                tbVervangingWerksoort.Text = dgWerksoortInfo.Rows[e.RowIndex].Cells["Vervanging Werksoort"].Value.ToString();

                btnAdd.Enabled = true;
                btnEdit.Enabled = true;
                btnDelete.Enabled = true;
                btnSave.Enabled = false;
                btnEdit.Focus();
            }
        }
        private void LoadWeksoortDetails()
        {
            var datFile = new FileInfo("ReplacementWeksoortInfo.dat");
            DataTable dtWerksoort = new DataTable();
            dtWerksoort.Columns.Add("Werksoort", typeof(string));
            dtWerksoort.Columns.Add("Vervanging Werksoort", typeof(string));
            if (datFile.Exists)
            {
                using (StreamReader sr = new StreamReader("ReplacementWeksoortInfo.dat", System.Text.Encoding.Default))
                {
                    var myStringRow = sr.ReadLine();
                    //runs until string reader returns null and adds rows to dt 
                    while (myStringRow != null)
                    {
                        if (!string.IsNullOrEmpty(myStringRow))
                        {
                            var rows = myStringRow.Split(';');
                            DataRow drWerksoort = dtWerksoort.NewRow();
                            drWerksoort["Werksoort"] = rows[0];
                            drWerksoort["Vervanging Werksoort"] = rows[1];
                            dtWerksoort.Rows.Add(drWerksoort);
                        }
                        myStringRow = sr.ReadLine();
                    }
                }
                dgWerksoortInfo.DataSource = dtWerksoort;
                SetDataGridViewStyles();
            }
        }
        private void SetDataGridViewStyles()
        {
            dgWerksoortInfo.ColumnHeadersDefaultCellStyle.Font = new Font(dgWerksoortInfo.DefaultCellStyle.Font.Name, 12, FontStyle.Regular);
            dgWerksoortInfo.ColumnHeadersDefaultCellStyle.SelectionBackColor = dgWerksoortInfo.ColumnHeadersDefaultCellStyle.BackColor;
            dgWerksoortInfo.ColumnHeadersDefaultCellStyle.SelectionForeColor = dgWerksoortInfo.ColumnHeadersDefaultCellStyle.ForeColor;

            dgWerksoortInfo.DefaultCellStyle.Font = new Font(dgWerksoortInfo.DefaultCellStyle.Font.Name, 12, FontStyle.Regular);
            dgWerksoortInfo.AutoResizeColumnHeadersHeight();
            int lastCol = 0;
            for (int i = dgWerksoortInfo.Columns.Count - 1; i >= 0; i--)
            {
                if (dgWerksoortInfo.Columns[i].Visible)
                {
                    dgWerksoortInfo.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgWerksoortInfo.Columns[i].Width = dgWerksoortInfo.Columns[i].Width;
                    dgWerksoortInfo.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    lastCol = i;
                }
            }

            if (dgWerksoortInfo.Columns[lastCol].Visible)
            {
                dgWerksoortInfo.Columns[lastCol].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgWerksoortInfo.Columns[lastCol].Width = dgWerksoortInfo.Columns[dgWerksoortInfo.Columns.Count - 1].Width;
                dgWerksoortInfo.Columns[lastCol].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }

            dgWerksoortInfo.EnableHeadersVisualStyles = false;
            dgWerksoortInfo.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgWerksoortInfo.ReadOnly = true;
            dgWerksoortInfo.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        private void ResetWerksoortPanel()
        {
            tbWerksoort.Text = "";
            tbVervangingWerksoort.Text = "";
            tbWerksoort.Enabled = false;
            tbVervangingWerksoort.Enabled = false;
            btnAdd.Enabled = true;
            btnSave.Enabled = false;
            btnEdit.Enabled = false;
            btnDelete.Enabled = false;
            werksoortEditMode = false;
            LoadWeksoortDetails();
            dgWerksoortInfo.Enabled = true;
        }
        #endregion Werksoort replacement
    }
}
