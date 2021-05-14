using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace AclUpdateEntity
{
    // Left
    public static class StringExtensions
    //https://stackoverflow.com/questions/7574606/left-function-in-c-sharp/7574645
    {
        public static string Left(this string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value)) return value;
        maxLength = Math.Abs(maxLength);

        return (value.Length <= maxLength
               ? value
               : value.Substring(0, maxLength)
               );
    }
}

public partial class AclUpdEntity : System.Web.UI.Page
    {
        public const string cCrLf = "\r\n";

        protected void Page_Load(object sender, EventArgs e)
        {
            // No processing
        }


        protected void UploadButton_Click(Object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                string filePath = string.Empty;

                string path = System.Web.HttpContext.Current.Server.MapPath("~/workfiles/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                // Get file path
                filePath = path + Path.GetFileName(FileUpload1.FileName);
                // Get file extenstion
                string extension = Path.GetExtension(FileUpload1.FileName);
                // Save file on "workfiles" folder of project
                FileUpload1.SaveAs(filePath);

                string conString = string.Empty;
                // Check file extension
                switch (extension)
                {
                    case ".xls": // Excel 97-03.
                        //conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        // See (https://stackoverflow.com/questions/1991643/microsoft-jet-oledb-4-0-provider-is-not-registered-on-the-local-machine No.6)
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        break;

                    case ".xlsx": // Excel 07 and above.
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                }

                // Create datatable object
                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                // Use OldDb to read excel
                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            // Get the name of First Sheet
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            // Read Data from First Sheet
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "A1:L7000]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }


                    // Create Destination/Update DT
                    DataTable dtUpd = new DataTable();
                    dtUpd.Columns.Add("Reference", Type.GetType("System.String"));
                    dtUpd.Columns.Add("RefNum", Type.GetType("System.String"));
                    dtUpd.Columns.Add("RefChar", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Name of Individual Or Entity", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Type", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Name Type", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Date of Birth", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Place of Birth", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Citizenship", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Address", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Additional Information", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Listing Information", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Committees", Type.GetType("System.String"));
                    dtUpd.Columns.Add("Control Date", Type.GetType("System.String"));

                    // Read the datatable (dt)
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sReference = FixNull(dr["Reference"].ToString().Trim());
                        string sRefNum = string.Empty;
                        string sRefChar = string.Empty;
                        Boolean bHasAlpha = isAlphaOnly(sReference);
                        if (sReference.Length > 1 & (bHasAlpha)) {
                            // string left = s.Left(number);  (See class StringExtensions - Left above)
                            sRefNum = sReference.Left(sReference.Length - 1);
                            // (Right) x = y.Substring(y.Length - z);
                            sRefChar = sReference.Substring(sReference.Length - 1);
                        }
                        else
                        {
                            sRefNum = sReference;
                        }
                        string sNameOfIndividualOrEntity = FixCrLfNl(FixNull(dr["Name of Individual Or Entity"].ToString()));
                        string sType = FixCrLfNl(FixNull(dr["Type"].ToString()));
                        string sNameType = FixCrLfNl(FixNull(dr["Name Type"].ToString()));
                        string sDateOfBirth = FixCrLfNl(FixNull(dr["Date of Birth"].ToString()));
                        string sPlaceOfBirth = FixCrLfNl(FixNull(dr["Place of Birth"].ToString()));
                        string sCitizenship = FixCrLfNl(FixNull(dr["Citizenship"].ToString()));
                        string sAddress = FixCrLfNl(FixNull(dr["Address"].ToString()));
                        string sAdditionalInformation = FixCrLfNl(FixNull(dr["Additional Information"].ToString()));
                        string sListingInformation = FixCrLfNl(FixNull(dr["Listing Information"].ToString()));
                        string sCommittees = FixCrLfNl(FixNull(dr["Committees"].ToString()));
                        string sControlDate = FixCrLfNl(FixNull(FixDate(dr["Control Date"].ToString())));

                        // Add the row data to dtUpd DT
                        DataRow newRow = dtUpd.NewRow();
                        dtUpd.Rows.Add(sReference, sRefNum, sRefChar, sNameOfIndividualOrEntity, sType, sNameType, sDateOfBirth, sPlaceOfBirth, sCitizenship, sAddress, sAdditionalInformation, sListingInformation, sCommittees, sControlDate);

                        Output.Text += "Reference: " + sReference + cCrLf;
                        Output.Text += "RefNum: " + sRefNum + cCrLf;
                        Output.Text += "RefChar: " + sRefChar + cCrLf;
                        Output.Text += sNameOfIndividualOrEntity + cCrLf;
                        Output.Text += sType + cCrLf;
                        Output.Text += sNameType + cCrLf;
                        Output.Text += sDateOfBirth + cCrLf;
                        Output.Text += sPlaceOfBirth + cCrLf;
                        Output.Text += sCitizenship + cCrLf;
                        Output.Text += sAddress + cCrLf;
                        Output.Text += sAdditionalInformation + cCrLf;
                        // Parse sAdditionalInformation  (InStr?, or Split?)


                        Output.Text += sListingInformation + cCrLf;
                        Output.Text += sCommittees + cCrLf;
                        Output.Text += sControlDate + cCrLf;
                        Output.Text += "-------------------------------------------------------------------" + cCrLf;
                    }


                }

                // bind datatable with GridView
                GridView1.DataSource = dt;
                GridView1.DataBind();

                // Clean up (delete) files in workfiles folder older than 61 days
                DirectoryInfo d = new DirectoryInfo(path);
                if (d.CreationTime < DateTime.Now.AddDays(-61))
                    d.Delete();

            }
            else
            {
                // You did not specify a file to upload
            }
        }






        //*********************************************************************
        // Misc Functions
        //*********************************************************************
        public static string FixNull(object dbvalue)
        {
            if (dbvalue == DBNull.Value)
                return "";
            else if (dbvalue == null)
                return string.Empty;
            else
                // NOTE: This will cast value to string if
                // it isn't a string.
                return dbvalue.ToString().Trim();
        }
        public static string FixCrLfNl(string CrLfNl)
        {
            CrLfNl = CrLfNl.Replace("'", "''");
            CrLfNl = CrLfNl.Replace(System.Environment.NewLine, string.Empty);
            CrLfNl = CrLfNl.TrimEnd('\r', '\n');

            return CrLfNl.ToString().Trim();
        }

        public static string FixDate(object sdate)
        {
            if (sdate.ToString().Length > 0)
            {
                string sVal = sdate.ToString().Replace(" 12:00:00 AM", "");
                return sVal.Trim();
            }
            else
                return Convert.ToString(sdate).Trim();
        }

        public static string FixApos(string apos_str)
        {
            if (apos_str.Contains('\''))
            {
                apos_str = apos_str.Replace("'","''");  
            }
            return apos_str.ToString().Trim();
        }

        public static Boolean isAlphaOnly(string strToCheck)
        {
            Regex rg = new Regex(@"[a-zA-Z]");
            return rg.IsMatch(strToCheck);
        }


    }
}