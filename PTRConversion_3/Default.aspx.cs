using System;
using System.IO;
using System.Web.UI;
using System.Collections.Generic;
using System.IO.Compression;
using Syncfusion.XlsIO;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Xml.Linq;
using System.Web;

namespace PTRConversion_3
{
    internal class Transaction
    {
        #region Transactions
        private int m_ref_number;
        private string m_transaction_date;
        private string m_from_or_to;
        private double m_amount_local;
        private string m_foreign_currency_code;
        private double m_foreign_amount;
        private double m_foreign_exchange_rate;
        private string m_first_name;
        private string m_middle_name;
        private string m_last_name;
        private string m_dob;
        private string m_gender;
        private string m_email;
        private string m_id_type;
        private string m_id_number;
        private string m_address;
        private string m_city;
        private string m_client_country_code;
        private string m_bank_name;
        private string m_bank_swift;
        private string m_bank_account_number;
        #endregion

        #region Prperties
        [DisplayNameAttribute("ref_number transaction_date from_or_to amount_local foreign_currency_code foreign_amount foreign_exchange_rate first_name middle_name last_name dob gender email id_type id_number address city client_country_code bank_name bank_swift bank_account_number")]
        public int ref_number
        {
            get
            {
                return m_ref_number;
            }
            set
            {
                m_ref_number = value;
            }
        }

        public string transaction_date
        {
            get
            {
                return m_transaction_date;
            }
            set
            {
                m_transaction_date = value;
            }
        }

        public string from_or_to
        {
            get
            {
                return m_from_or_to;
            }
            set
            {
                m_from_or_to = value;
            }
        }

        public double amount_local
        {
            get
            {
                return m_amount_local;
            }
            set
            {
                m_amount_local = value;
            }
        }

        public string foreign_currency_code
        {
            get
            {
                return m_foreign_currency_code;
            }
            set
            {
                m_foreign_currency_code = value;
            }
        }

        public double foreign_amount
        {
            get
            {
                return m_foreign_amount;
            }
            set
            {
                m_foreign_amount = value;
            }
        }

        public double foreign_exchange_rate
        {
            get
            {
                return m_foreign_exchange_rate;
            }
            set
            {
                m_foreign_exchange_rate = value;
            }
        }

        public string first_name
        {
            get
            {
                return m_first_name;
            }
            set
            {
                m_first_name = value;
            }
        }

        public string middle_name
        {
            get
            {
                return m_middle_name;
            }
            set
            {
                m_middle_name = value;
            }
        }

        public string last_name
        {
            get
            {
                return m_last_name;
            }
            set
            {
                m_last_name = value;
            }
        }

        public string dob
        {
            get
            {
                return m_dob;
            }
            set
            {
                m_dob = value;
            }

        }

        public string gender
        {
            get
            {
                return m_gender;
            }
            set
            {
                m_gender = value;
            }
        }

        public string email
        {
            get
            {
                return m_email;
            }
            set
            {
                m_email = value;
            }
        }

        public string id_type
        {
            get
            {
                return m_id_type;
            }
            set
            {
                m_id_type = value;
            }
        }

        public string id_number
        {
            get
            {
                return m_id_number;
            }
            set
            {
                m_id_number = value;
            }
        }

        public string address
        {
            get
            {
                return m_address;
            }
            set
            {
                m_address = value;
            }
        }

        public string city
        {
            get
            {
                return m_city;
            }
            set
            {
                m_city = value;
            }
        }

        public string client_country_code
        {
            get
            {
                return m_client_country_code;
            }
            set
            {
                m_client_country_code = value;
            }

        }

        public string bank_name
        {
            get
            {
                return m_bank_name;
            }
            set
            {
                m_bank_name = value;
            }
        }

        public string bank_swift
        {
            get
            {
                return m_bank_swift;
            }
            set
            {
                m_bank_swift = value;
            }
        }

        public string bank_account_number
        {
            get
            {
                return m_bank_account_number;
            }
            set
            {
                m_bank_account_number = value;
            }
        }
        #endregion

        #region Intialization
        public Transaction() { }

        #endregion
    }

    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack) return;
            StatusLabel.Text = "Status: Please update your excel file.";
            Clear_Files();
        }

        protected void BtnUpload_Click(object sender, EventArgs e)
        {
            welcomeMessage.Visible = false;
            errorMessage.Visible = false;
            HttpPostedFile postedFile = System.Web.HttpContext.Current.Request.Files[0];
            if (postedFile != null && postedFile.ContentLength > 0)
            {
                try
                {
                    if (FileUploadControl.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        if (FileUploadControl.PostedFile.ContentLength < 102400)
                        {
                            DateTimeOffset dateNow = DateTimeOffset.UtcNow;
                            
                            string timeStamp = " @" + dateNow.ToString("yyyy-MM-dd hh_mm_ss");
                            string filenameConvert = "PTR.xlsx";
                            string filenameArchive = "PTR" + timeStamp + ".xlsx";
                            postedFile.SaveAs(Server.MapPath("~/Uploaded/") + filenameConvert);
                            postedFile.SaveAs(Server.MapPath("~/Archives/") + filenameArchive);
                            btnUpload.Attributes.Add("disabled", "true");
                            FileUploadControl.Attributes.Add("disabled", "true");
                            btnConvert.Attributes.Remove("disabled");
                            inputBox.Attributes["class"] = "file has-name is-boxed is-success";
                            StatusLabel.Text = "Upload status: File uploaded! Now you can convert it!";
                            statusMessage.Visible = true;
                        }
                        else
                        {
                            inputBox.Attributes["class"] = "file has-name is-boxed is-danger";
                            ErrorLabel.Text = "Upload status: The file has to be less than 100 kb!";
                            errorMessage.Visible = true;
                        }
                    }
                    else
                    {
                        inputBox.Attributes["class"] = "file has-name is-boxed is-danger";
                        ErrorLabel.Text = "Upload status: Only XLSX files are accepted!";
                        errorMessage.Visible = true;
                    }

                }
                catch (Exception ex)
                {
                    inputBox.Attributes["class"] = "file has-name is-boxed is-danger";
                    ErrorLabel.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message;
                    errorMessage.Visible = true;
                }
            }
            else
            {
                if (postedFile == null)
                {
                    inputBox.Attributes["class"] = "file has-name is-boxed is-danger";
                    ErrorLabel.Text = "Upload status: null";
                    errorMessage.Visible = true;
                }
                if (postedFile != null && postedFile.ContentLength <= 0)
                {
                    inputBox.Attributes["class"] = "file has-name is-boxed is-danger";
                    ErrorLabel.Text = "Upload status: <=0";
                    errorMessage.Visible = true;
                }
            }
        }

        protected void BtnDownload_Click(object sender, EventArgs e)
        {
            string startPath = Server.MapPath("~/Converted/");
            string zipPath = Server.MapPath("~/Download/PTR.zip");
            ZipFile.CreateFromDirectory(startPath, zipPath);
            Response.ContentType = "application/zip";
            Response.AppendHeader("Content-Disposition", "attachment; filename=PTR.zip");
            Response.TransmitFile(Server.MapPath("~/Download/PTR.zip"));
            Response.End();
            btnDownload.Attributes.Add("disabled", "true");
        }

        private void Jsonise_Excel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                //The workbook is opened.
                using (FileStream fileStream = new FileStream(Server.MapPath("~/Uploaded/PTR.xlsx"), FileMode.Open))
                {
                    IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    //Export worksheet data into CLR Objects
                    IList<Transaction> transactions = worksheet.ExportData<Transaction>(1, 1, worksheet.UsedRange.LastRow, workbook.Worksheets[0].UsedRange.LastColumn);

                    //open file stream
                    using (StreamWriter file = File.CreateText(Server.MapPath("~/Uploaded/PTR.json")))
                    {
                        JsonSerializer serializer = new JsonSerializer();

                        //serialize object directly into file stream
                        serializer.Serialize(file, transactions);

                    }
                }
            }
        }

        public string Reference_Number_Generator()
        {
            const string preRefNumber = "TW01";
            Random r = new Random();
            int randNum = r.Next(1000000);
            string sixDigitNumber = randNum.ToString("D6");
            int randNum2 = r.Next(1000000);
            string sixDigitNumber2 = randNum2.ToString("D6");
            string outputRefNumber = preRefNumber + randNum + "-" + randNum2;
            return outputRefNumber;
        }

        public void Json_To_Xml()
        {
            List<Transaction> transList;
            //int rentityId = 52038; //this is for the real environment
            int rentityId = 50379; //this is for the test environment
            int transactionCount = 1;

            using (StreamReader r = new StreamReader(Server.MapPath("~/Uploaded/PTR.json")))
            {
                string json = r.ReadToEnd();
                transList = JsonConvert.DeserializeObject<List<Transaction>>(json);
            }
            
            foreach (Transaction eachTransaction in transList)
            {
                string refNumber = Reference_Number_Generator();
                string xmlFilename = "PTR_" + transactionCount.ToString() + ".xml";

                switch (eachTransaction.from_or_to.ToString())
                {
                    case "F":
                    {
                        XDocument fromTransXml = new XDocument(
                            new XElement("report",
                                new XElement("rentity_id", rentityId.ToString()),
                                new XElement("submission_code", "E"),
                                new XElement("report_code", "IFT"),
                                new XElement("entity_reference", eachTransaction.ref_number.ToString()),
                                new XElement("submission_date", DateTimeOffset.Now.ToString("o")),
                                new XElement("currency_code_local", "NZD"),
                                new XElement("reporting_person",
                                    new XElement("gender", "M"),
                                    new XElement("title", "Mr"),
                                    new XElement("first_name", "Sai Shong"),
                                    new XElement("last_name", "Leung"),
                                    new XElement("phones",
                                        new XElement("phone",
                                            new XElement("tph_contact_type", "H"),
                                            new XElement("tph_communication_type", "B"),
                                            new XElement("tph_number", "+6492803716"))),
                                    new XElement("addresses",
                                        new XElement("address",
                                            new XElement("address_type", "H"),
                                            new XElement("address", "Level 2, 3 Margot Street"),
                                            new XElement("city", "Auckland"),
                                            new XElement("zip", "1051"),
                                            new XElement("country_code", "NZ"),
                                            new XElement("state", "Epsom"))),
                                    new XElement("email", "edward.leung@cjcmarkets.com")),
                                new XElement("reason", "n/a"),
                                new XElement("action", "n/a"),
                                new XElement("transaction",
                                    new XElement("transactionnumber", refNumber),
                                    new XElement("internal_ref_number", eachTransaction.ref_number.ToString()),
                                    new XElement("transaction_location", "HEAD OFFICE"),
                                    new XElement("date_transaction",
                                        eachTransaction.transaction_date.ToString() + "T00:00:00"),
                                    new XElement("transmode_code", "BA"),
                                    new XElement("amount_local", eachTransaction.amount_local.ToString()),
                                    new XElement("t_from",
                                        new XElement("from_funds_code", "N"),
                                        new XElement("from_foreign_currency",
                                            new XElement("foreign_currency_code",
                                                eachTransaction.foreign_currency_code.ToString()),
                                            new XElement("foreign_amount", eachTransaction.foreign_amount.ToString()),
                                            new XElement("foreign_exchange_rate",
                                                eachTransaction.foreign_exchange_rate.ToString())),
                                        new XElement("from_account",
                                            new XElement("institution_name", eachTransaction.bank_name.ToString()),
                                            new XElement("swift", eachTransaction.bank_swift.ToString()),
                                            new XElement("non_bank_institution", "false"),
                                            new XElement("account", eachTransaction.bank_account_number.ToString()),
                                            new XElement("signatory",
                                                new XElement("is_primary", "true"),
                                                new XElement("t_person",
                                                    new XElement("gender", eachTransaction.gender.ToString()),
                                                    new XElement("first_name", eachTransaction.first_name.ToString()),
                                                    new XElement("last_name", eachTransaction.last_name.ToString()),
                                                    new XElement("birthdate", eachTransaction.dob.ToString() + "T00:00:00"),
                                                    new XElement("phones",
                                                        new XElement("phone",
                                                            new XElement("tph_contact_type", "I"),
                                                            new XElement("tph_communication_type", "C"),
                                                            new XElement("tph_number", "n/a"))),
                                                    new XElement("addresses",
                                                        new XElement("address",
                                                            new XElement("address_type", "I"),
                                                            new XElement("address", eachTransaction.address.ToString()),
                                                            new XElement("city", eachTransaction.city.ToString()),
                                                            new XElement("zip", "n/a"),
                                                            new XElement("country_code",
                                                                eachTransaction.client_country_code.ToString()))),
                                                    new XElement("email", eachTransaction.email.ToString()),
                                                    new XElement("identification",
                                                        new XElement("type", eachTransaction.id_type.ToString()),
                                                        new XElement("number", eachTransaction.id_number.ToString()),
                                                        new XElement("issue_country",
                                                            eachTransaction.client_country_code.ToString()))),
                                                new XElement("role", "A"))),
                                        new XElement("from_country", eachTransaction.client_country_code.ToString())),
                                    new XElement("t_to_my_client",
                                        new XElement("to_funds_code", "N"),
                                        new XElement("to_account",
                                            new XElement("institution_name", "ASB Bank Limited"),
                                            new XElement("swift", "ASBBNZ2A"),
                                            new XElement("non_bank_institution", "false"),
                                            new XElement("branch", "Auckland Central"),
                                            new XElement("account", "26957204-USD-26"),
                                            new XElement("t_entity",
                                                new XElement("name", "Carrick Just Capital Markets Limited"),
                                                new XElement("incorporation_number", "6358657"),
                                                new XElement("phones",
                                                    new XElement("phone",
                                                        new XElement("tph_contact_type", "H"),
                                                        new XElement("tph_communication_type", "B"),
                                                        new XElement("tph_number", "092803716"))),
                                                new XElement("addresses",
                                                    new XElement("address",
                                                        new XElement("address_type", "H"),
                                                        new XElement("address", "Level 2, 3 Margot Street"),
                                                        new XElement("city", "Auckland"),
                                                        new XElement("zip", "1051"),
                                                        new XElement("country_code", "NZ"),
                                                        new XElement("state", "Epsom"))),
                                                new XElement("incorporation_country_code", "NZ"),
                                                new XElement("director_id",
                                                    new XElement("gender", "M"),
                                                    new XElement("first_name", "JIAN"),
                                                    new XElement("last_name", "SHEN"),
                                                    new XElement("birthdate", "1981-04-29T00:00:00"),
                                                    new XElement("identification",
                                                        new XElement("type", "F"),
                                                        new XElement("number", "LK960322"),
                                                        new XElement("issue_country", "NZ")),
                                                    new XElement("role", "A")),
                                                new XElement("incorporation_date", "2017-07-28T00:00:00")),
                                            new XElement("beneficiary_comment",
                                                eachTransaction.bank_name + ":" + eachTransaction.bank_account_number)),
                                        new XElement("to_country", "NZ"))),
                                new XElement("report_indicators",
                                    new XElement("indicator", "FMA"))));
                        fromTransXml.Save(Server.MapPath("~/Converted/" + xmlFilename));
                        break;
                    }

                    case "T":
                    {
                        XDocument toTransXml = new XDocument(
                            new XElement("report",
                                new XElement("rentity_id", rentityId.ToString()),
                                new XElement("submission_code", "E"),
                                new XElement("report_code", "IFT"),
                                new XElement("entity_reference", eachTransaction.ref_number.ToString()),
                                new XElement("submission_date", DateTimeOffset.Now.ToString("o")),
                                new XElement("currency_code_local", "NZD"),
                                new XElement("reporting_person",
                                    new XElement("gender", "M"),
                                    new XElement("title", "Mr"),
                                    new XElement("first_name", "Sai Shong"),
                                    new XElement("last_name", "Leung"),
                                    new XElement("phones",
                                        new XElement("phone",
                                            new XElement("tph_contact_type", "H"),
                                            new XElement("tph_communication_type", "B"),
                                            new XElement("tph_number", "+6492803716"))),
                                    new XElement("addresses",
                                        new XElement("address",
                                            new XElement("address_type", "H"),
                                            new XElement("address", "Level 2, 3 Margot Street"),
                                            new XElement("city", "Auckland"),
                                            new XElement("zip", "1051"),
                                            new XElement("country_code", "NZ"),
                                            new XElement("state", "Epsom"))),
                                    new XElement("email", "edward.leung@cjcmarkets.com")),
                                new XElement("reason", "n/a"),
                                new XElement("action", "n/a"),
                                new XElement("transaction",
                                    new XElement("transactionnumber", refNumber),
                                    new XElement("internal_ref_number", eachTransaction.ref_number.ToString()),
                                    new XElement("transaction_location", "HEAD OFFICE"),
                                    new XElement("date_transaction",
                                        eachTransaction.transaction_date.ToString() + "T00:00:00"),
                                    new XElement("transmode_code", "BA"),
                                    new XElement("amount_local", eachTransaction.amount_local.ToString()),
                                    new XElement("t_from",
                                        new XElement("from_funds_code", "N"),
                                        new XElement("from_foreign_currency",
                                            new XElement("foreign_currency_code",
                                                eachTransaction.foreign_currency_code.ToString()),
                                            new XElement("foreign_amount", eachTransaction.foreign_amount.ToString()),
                                            new XElement("foreign_exchange_rate",
                                                eachTransaction.foreign_exchange_rate.ToString())),
                                        new XElement("from_account",
                                            new XElement("institution_name", eachTransaction.bank_name.ToString()),
                                            new XElement("swift", eachTransaction.bank_swift.ToString()),
                                            new XElement("non_bank_institution", "false"),
                                            new XElement("account", eachTransaction.bank_account_number.ToString())),
                                        new XElement("from_country", eachTransaction.client_country_code.ToString())),
                                        new XElement("t_to_my_client",
                                            new XElement("to_funds_code", "N"),
                                            new XElement("to_account",
                                                new XElement("institution_name", "Carrick Just Capital Markets Limited"),
                                                new XElement("institution_code", rentityId.ToString()),
                                                new XElement("non_bank_institution", "true"),
                                                new XElement("branch", "Epsom, Auckland"),
                                                new XElement("account", "85201002"),
                                                new XElement("signatory",
                                                    new XElement("is_primary", "true"),
                                                    new XElement("t_person",
                                                        new XElement("gender", eachTransaction.gender.ToString()),
                                                        new XElement("first_name", eachTransaction.first_name.ToString()),
                                                        new XElement("last_name", eachTransaction.last_name.ToString()),
                                                        new XElement("birthdate", eachTransaction.dob.ToString() + "T00:00:00"),
                                                        new XElement("phones",
                                                            new XElement("phone",
                                                                new XElement("tph_contact_type", "I"),
                                                                new XElement("tph_communication_type", "C"),
                                                                new XElement("tph_number", "n/a"))),
                                                        new XElement("addresses",
                                                            new XElement("address",
                                                                new XElement("address_type", "I"),
                                                                new XElement("address", eachTransaction.address.ToString()),
                                                                new XElement("city", eachTransaction.city.ToString()),
                                                                new XElement("zip", "n/a"),
                                                                new XElement("country_code",
                                                                    eachTransaction.client_country_code.ToString()))),
                                                        new XElement("email", eachTransaction.email.ToString()),
                                                        new XElement("identification",
                                                            new XElement("type", eachTransaction.id_type.ToString()),
                                                            new XElement("number", eachTransaction.id_number.ToString()),
                                                            new XElement("issue_country",
                                                                eachTransaction.client_country_code.ToString()))),
                                                    new XElement("role", "A"))),
                                            new XElement("to_country", "NZ"))),
                                    new XElement("report_indicators",
                                        new XElement("indicator", "FMA"))));

                        toTransXml.Save(Server.MapPath("~/Converted/" + xmlFilename));
                        break;
                    }
                }

                transactionCount += 1;
            }
        }

        protected void BtnConvert_Click(Object sender, EventArgs e)
        {
            errorMessage.Visible = false;
            StatusLabel.Text = "Status: Reading your file now...";
            Jsonise_Excel();
            Json_To_Xml();
            StatusLabel.Text = "Status: Conversion completed. Ready to go...";
            btnConvert.Attributes.Add("disabled", "true");
            btnDownload.Attributes.Remove("disabled");
        }

        private void Clear_Files()
        {
            DirectoryInfo dirUploaded = new DirectoryInfo(Server.MapPath("~/Converted/"));
            DirectoryInfo dirConverted = new DirectoryInfo(Server.MapPath("~/Converted/"));
            DirectoryInfo dirDownload = new DirectoryInfo(Server.MapPath("~/Download/"));
            foreach (FileInfo file in dirUploaded.GetFiles())
            {
                file.Delete();
            }
            foreach (FileInfo file in dirConverted.GetFiles())
            {
                file.Delete();
            }
            foreach (FileInfo file in dirDownload.GetFiles())
            {
                file.Delete();
            }
        }

        protected void BtnReset_Click(object sender, EventArgs e)
        {
            Clear_Files();
            btnUpload.Attributes.Remove("disabled");
            inputBox.Attributes["class"] = "file has-name is-boxed";
            FileUploadControl.Disabled = false;
            statusMessage.Visible = false;
            errorMessage.Visible = false;
            welcomeMessage.Visible = true;
        }
    }
}