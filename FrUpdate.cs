using System;
using System.IO;
using System.Configuration;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;

namespace FRUpdate
{
    static class Program
    {
        public static string wl4email;

        static int Main(string[] args)
        {
        string wl;
        wl = WriteLog("\r\n" + "Start DateTime: " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + " ---------------------------------" + Environment.NewLine);
            var fr_debug = ConfigurationManager.AppSettings["FR_Debug"];

            // ArchiveFolder
            string AF = ArchiveFolder(ConfigurationManager.AppSettings["FR_Folder"] + "archive");
            if (AF != "0")
            {
                wl = WriteLog("ArchiveFolder Creation Error: " + AF + Environment.NewLine);
                return 0;
            }

            // GetAndValidate
            DataTable dtRet = GetAndValidate();
            string ErrMessages = string.Empty;

            if (dtRet.Rows.Count > 0)
            {
                // ReadFiles returns a DataTable 
                DataTable dtRetData = new DataTable();
                dtRetData = ReadFiles(dtRet);

                // Check for size of dtRetData and start processing it if records present
                if (dtRetData != null)
                {
                    // Call WriteFiles(dtRetData)
                    string retWriteFiles = WriteFiles(dtRetData);
                    if (retWriteFiles != "0")
                    {
                        ErrorHandler("FRUpdate Error", retWriteFiles);
                    }
                }
            }
            else
            {
                ErrorHandler("FRUpdate Success", "GetAndValidate() - No FR Files Found.");
            }

            //Console.WriteLine("No Errors");
            //Console.ReadKey();

            wl = WriteLog("End DateTime: " + DateTime.Now.ToString("yyyyMMdd_HHmmss")) + " ---------------------------------" + Environment.NewLine;
            string retSendEmail3 = SendEmail3("FRUpdate Success", wl4email);

            return 0;
        }   // Main



        public static DataTable GetAndValidate()
        {
            string ErrMessage = string.Empty;

            DataTable dt = new DataTable();
            dt.Columns.Add("File", typeof(string));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("ProcessDate", typeof(DateTime));
            dt.Columns.Add("Date", typeof(String));
            dt.Columns.Add("Volume", typeof(int));
            dt.Columns.Add("Page", typeof(int));
            dt.Columns.Add("ErrMessage", typeof(string));

            string fFile = string.Empty;
            string fType = string.Empty;
            string fDate = string.Empty;
            string fVolume = string.Empty;
            string fPage = string.Empty;

            var fr_folder = ConfigurationManager.AppSettings["FR_Folder"];
            var fr_emailsender = ConfigurationManager.AppSettings["FR_EmailSender"];
            var fr_emailrecipients = ConfigurationManager.AppSettings["FR_EmailRecipients"];
            var fr_debug = ConfigurationManager.AppSettings["FR_Debug"];
           
            WriteLog("FR_Folder: " + fr_folder + Environment.NewLine);
            WriteLog("FR_EmailSender: " + fr_emailsender + Environment.NewLine);
            WriteLog("FR_EmailRecipients: " + fr_emailrecipients + Environment.NewLine);

            bool bDateError = false;
            bool bVolumeError = false;
            bool bPageError = false;

            if (fr_debug == "1") Console.WriteLine("FR_Folder: " + fr_folder);
            if (fr_debug == "1") Console.WriteLine("FR_EmailSender: " + fr_emailsender);
            if (fr_debug == "1") Console.WriteLine("FR_EmailRecipients: " + fr_emailrecipients);

            try
            {
            string chkNum = string.Empty;
            foreach (var file in Directory.GetFiles(@fr_folder).OrderBy(f => f))
            {
                ErrMessage = string.Empty;
                if (fr_debug == "1") Console.WriteLine(file);

                string file_r = StrRev(file);
                if (fr_debug == "1") Console.WriteLine(file_r);

                int index = file_r.IndexOf(@"\");
                if (fr_debug == "1") Console.WriteLine($"{index}");

                string file_Left = StrLeft(file_r, index);
                if (fr_debug == "1") Console.WriteLine(file_Left);

                string file_name = StrRev(file_Left);
                if (fr_debug == "1") Console.WriteLine(file_name);
                WriteLog("file_name: " + file_name + Environment.NewLine);

                // Example FileName - FR_20190730_100_50.txt
                // Parse FileName
                string[] FN_elements = file_name.Split("_".ToCharArray());
                int FNlen = FN_elements.Length;
                if (FNlen == 4) 
                {
                    //if (fr_debug == "1") Console.WriteLine(element);
                    fType = FN_elements[0];
                    WriteLog("FileType: " + fType + Environment.NewLine);

                    WriteLog("FN_elements[1] (Date): " + FN_elements[1] + Environment.NewLine);
                    fDate = ValidateDateStr(FN_elements[1]);
                    // fDate
                    if (fDate == "1")
                    {
                        if (!bDateError)
                        {
                            ErrMessage += "FileName: " + file_name + ". ";
                            ErrMessage += Environment.NewLine + FN_elements[1] + " Filename uses Invalid date, Use YYYYMMDD. ";
                            bDateError = true;
                        }
                    }

                    // fVolume
                    fVolume = FN_elements[2];
                    WriteLog("fVolume: " + fVolume + Environment.NewLine);
                    chkNum = CheckIfNumeric(fVolume);
                    if (chkNum != "0")
                    {
                        ErrMessage += "FileName: " + file_name + ". ";
                        ErrMessage += Environment.NewLine + fVolume + " Filename Volume must be numeric. ";
                        bVolumeError = true;
                        chkNum = "0";
                    }

                    // fPage
                    fPage = FN_elements[3].Replace(".txt", "");
                    WriteLog("fPage: " + fPage + Environment.NewLine);
                    chkNum = CheckIfNumeric(fPage);
                    if (chkNum != "0")
                    {
                        ErrMessage += "FileName: " + file_name + ". ";
                        ErrMessage += Environment.NewLine + fPage + " Filename Page must be numeric. ";
                        bPageError = true;
                        chkNum = "0";
                    }
                    chkNum = "0";
                }   // if len == 4

                if (fr_debug == "1")
                {
                    Console.WriteLine("Filename: " + file_name);
                    Console.WriteLine("Type: " + fType);
                    Console.WriteLine("Date: " + fDate);
                    Console.WriteLine("Volume: " + fVolume);
                    Console.WriteLine("Page: " + fPage);
                }

                // Check if (!bDateError & !bVolumeError & !bPageError)
                string retSendEmail3 = string.Empty;
                if (!bDateError & !bVolumeError & !bPageError)
                    {
                    // Add to dt
                    var dr = dt.NewRow();
                        dr["File"] = file_name;
                        dr["Type"] = fType;
                        dr["ProcessDate"] = DateTime.Now;
                        dr["Date"] = fDate;
                        dr["Volume"] = fVolume;
                        dr["Page"] = fPage;
                        dr["ErrMessage"] = ErrMessage;
                        dt.Rows.Add(dr);
                    }
                    else
                    {
                        ErrorHandler("FRUpdate Error", ErrMessage + Environment.NewLine);
                    }
                bDateError = false;
                bVolumeError = false;
                bPageError = false;
            }   //foreach (var file in Directory...
        }   // End try
        catch (Exception ex)
        {
            ErrorHandler("FRUpdate Error", ex.Message + Environment.NewLine + ErrMessage + Environment.NewLine);
        }

        finally
        {
            // finally
            // Keep the console window open in debug mode.
            //if (fr_debug == "1") Console.WriteLine("Press any key to exit.");
            //if (fr_debug == "1") Console.ReadKey();
        }


        if ((!bDateError) && (!bVolumeError) && (!bPageError))
        {
            return dt;
        }
        else
        {
            return null;
        }
    }   // GetAndValidate()


    public static DataTable ReadFiles(DataTable dt)
    {
        string ErrMessage = string.Empty;
        var fr_debug = ConfigurationManager.AppSettings["FR_Debug"];

        DataTable dtData = new DataTable();
        dtData.Columns.Add("IdNum", typeof(string));
        dtData.Columns.Add("File", typeof(string));
        dtData.Columns.Add("Type", typeof(string));
        dtData.Columns.Add("ProcessDate", typeof(DateTime));
        dtData.Columns.Add("Date", typeof(string));
        dtData.Columns.Add("Volume", typeof(int));
        dtData.Columns.Add("Page", typeof(int));

        try
        {
            foreach (DataRow row in dt.Rows)
            {
                string PathAndFile = ConfigurationManager.AppSettings["FR_Folder"] + row["File"];
                PathAndFile = PathAndFile.Replace("/", "//");

                int counter = 0;
                string line;

                // Read file line by line.  
                StreamReader file = new StreamReader(@PathAndFile);
                while ((line = file.ReadLine()) != null)
                {
                    string strLine = line.ToLower();
                    strLine = strLine.Replace("\t", "");    // Get rid of any tabs (\t)
                    string LineIsNum = CheckIfNumeric(strLine);     // "0" if Numeric | "1" if NOT Numeric

                    // Ignore lines with "idnum" and lines that are Not Numeric
                    if (strLine != "idnum" && LineIsNum == "0")
                    {
                        // Add line to dtData - Only Add line value if NOT "idnum" and IS numeric
                        var dr = dtData.NewRow();
                        dr["IdNum"] = strLine;
                        dr["File"] = row["File"];
                        dr["Type"] = row["Type"];
                        dr["ProcessDate"] = row["ProcessDate"];
                        dr["Date"] = row["Date"];
                        dr["Volume"] = row["Volume"];
                        dr["Page"] = row["Page"];
                        dtData.Rows.Add(dr);

                        counter++;
                    }
                }
                file.Close();
            }
        }   // End try

        catch (Exception ex)
        {
            ErrorHandler("FRUpdate Error", ex.Message + Environment.NewLine + ErrMessage + Environment.NewLine);
        }
        finally
        {
            // finally
        }

        return dtData;
    }   // ReadFiles()


        public static string WriteFiles(DataTable dt)
        {
            try
            {
                // Check for size of dt and start processing it if records present
                if (dt != null)
                {
                    string strFile = string.Empty;
                    string strIdNums = string.Empty;
                    int intVol = 0;
                    int intPage = 0;
                    string strDate = string.Empty;
                    DateTime oDate = default(DateTime);

                    var foundRows = (from DataRow dRow in dt.Rows
                                        select dRow["File"]).Distinct();

                    foreach (var File in foundRows)
                    {
                        string s_sqls = "";
                        WriteLog("Processing file: " + File + Environment.NewLine);

                        // Use the Select method to find all rows matching the filter.  
                        DataRow[] foundRows1 = dt.Select("File = '" + File + "'");

                        // Print column 0 of each returned row.  
                        for (int i = 0; i < foundRows1.Length; i++)
                        {
                            strIdNums += foundRows1[i][0] + ",";
                            strFile = (string)foundRows1[i][1];
                            strDate = (string)foundRows1[i][4];
                            oDate = DateTime.Parse(strDate);
                            intVol = (int)foundRows1[i][5];
                            intPage = (int)foundRows1[i][6];
                        }
                        strIdNums = strIdNums.Substring(0, strIdNums.Length - 1);     // Left.  Get string without trailing comma

                        // sql1 --get the overflow ids and store them in memory
                        string sql1 = "DECLARE @fedRegFull TABLE(ID int) " +
                                        "INSERT INTO @fedRegFull(ID) " +
                                        "SELECT IDNUM " +
                                        "FROM Denial " +
                                        "WHERE FedDate8 IS NOT NULL AND IDNUM IN (" + strIdNums + ")";
                        string s_sql1 = ExecuteSQL(sql1);
                        s_sqls += "sql1 returned: " + s_sql1 + Environment.NewLine;

                        // sql2 -- convert oldest fed reg to a url
                        string sql2 = "DECLARE @fedRegFull TABLE(ID int) " +
                                        "INSERT INTO url (IDNUM, [URL]) " +
                                        "SELECT IDNUM, " + "'http://www.uptodateregs.com/_data/TDO_FRnotice.asp?date= + CONVERT(VARCHAR(10), feddate1, 101) + &pg=" + intPage + "' " +
                                        "FROM denial d JOIN @fedRegFull frf ON d.IDNUM = frf.ID " +
                                        "WHERE FedDate8 IS NOT NULL";
                        string s_sql2 = ExecuteSQL(sql2);
                        s_sqls += "sql2 returned: " + s_sql2 + Environment.NewLine;

                        // sql3 -- Shift fed info down one level
                        string sql3 = "DECLARE @fedRegFull TABLE(ID int) " +
                                        "UPDATE denial " +
                                        "SET FedDate1 = FedDate2, FedDate2 = FedDate3, FedDate3 = FedDate4, FedDate4 = FedDate5, FedDate5 = FedDate6, FedDate6 = FedDate7, FedDate7 = FedDate8, FedDate8 = NULL, " +
                                        "Vol1 = Vol2, Vol2 = Vol3, Vol3 = Vol4, Vol4 = Vol5, Vol5 = Vol6, Vol6 = Vol7, Vol7 = Vol8, Vol8 = '', " +
                                        "Page1 = Page2, Page2 = Page3, Page3 = Page4, Page4 = Page5, Page5 = Page6, Page6 = Page7, Page7 = Page8, Page8 = '' " +
                                        "FROM Denial d JOIN @fedRegFull frf ON d.IDNUM = frf.ID";
                        string s_sql3 = ExecuteSQL(sql3);
                        s_sqls += "sql3 returned: " + s_sql3 + Environment.NewLine;

                        // sp4 & Call AddFRNotice stored procedure -- Do regular inserts
                        string s_sp = ExecuteSP("dbo.AddFRNotice", intVol, intPage, oDate, strIdNums);
                        s_sqls += "For File (" + File + ") which contains IdNums(" + strIdNums + ") SP returned: " + s_sp + Environment.NewLine;
                        WriteLog(s_sqls);

                        FileArchive(strFile);
                        strIdNums = string.Empty;
                    }   // End for each file
                }
            }
            catch (Exception ex)
            {
                ErrorHandler("FRUpdate Error", ex.Message + Environment.NewLine);
            }

            return "0";
        }   // WriteFiles()


        public static string ErrorHandler(string subject, string errorString)
        {
            try
            {
                string retSendEmail3 = string.Empty;
                // Send the error email
                WriteLog("FRUpdate Error: " + errorString + Environment.NewLine);
                WriteLog("FRUpdate is Closing" + Environment.NewLine);

                if (errorString == "GetAndValidate() - No FR Files Found.")
                {
                    retSendEmail3 = SendEmail3("FRUpdate Success", errorString + " FRUpdate is Closing" + Environment.NewLine);
                }
                else
                {
                    retSendEmail3 = SendEmail3("FRUpdate Error", errorString + " FRUpdate is Closing" + Environment.NewLine);
                }

                // Check if SendEmail failed
                if (retSendEmail3 != "0")
                {
                    WriteLog("=======> FRUpdate Email Error. " + errorString + Environment.NewLine + retSendEmail3 + Environment.NewLine + "********");
                    Environment.Exit(0);
                }
            }

            catch (Exception ex)
            {
                return ex.Message;
            }

            string wl = WriteLog("End EndTime: " + DateTime.Now.ToString("yyyyMMdd_HHmmss")) + Environment.NewLine;
            Environment.Exit(0);
            return "1";
        }

        public static string WriteLog(string strMessage)
        {
            try
            {
                wl4email += strMessage;
                // Log Folder
                string LF = ConfigurationManager.AppSettings["FR_Folder"] + "log\\";
                FileStream objFilestream = new FileStream(LF + "ProcessingLog.txt", FileMode.Append, FileAccess.Write);
                StreamWriter objStreamWriter = new StreamWriter(objFilestream);
                objStreamWriter.WriteLine(strMessage);
                objStreamWriter.Close();
                objFilestream.Close();
                return "0";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        public static string ExecuteSP(string sp, int intVol, int intPage, DateTime oDate, string strIdNums)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sp, con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Input Params
                    cmd.Parameters.AddWithValue("@Volume", intVol);
                    cmd.Parameters.AddWithValue("@Page", intPage);
                    cmd.Parameters.AddWithValue("@Date", oDate);
                    cmd.Parameters.AddWithValue("@ID", strIdNums);

                    // Output Param
                    SqlParameter outParam = new SqlParameter("@outParam", SqlDbType.VarChar, 2000);
                    outParam.ParameterName = "@OutParam";
                    outParam.Value = string.Empty;
                    outParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(outParam);

                    con.Open();
                    cmd.ExecuteNonQuery();

                    string outval = (string)cmd.Parameters["@OutParam"].Value;
                    //Console.WriteLine("Out value: {0}", outval);
                    con.Close();
                    return outval;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        public static string ExecuteSQL(string sqlQuery)
        {
            string retSQL = "0";
            try
            {
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString))
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    cmd.CommandText = sqlQuery;
                    cmd.CommandType = CommandType.Text;
                    connection.Open();
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    cmd.Dispose();
                    retSQL = count + " rows returned";
                    connection.Close();
                }
            }

            catch (Exception ex)
            {
                return ex.Message;
            }

            return retSQL;
        }

        public static string ArchiveFolder(string folder)
        {
            try
            {
                string ArchiveFolder = CreateArchiveSuccess(folder);
                if (ArchiveFolder != "0")
                {
                    return "Archive Folder create failed.";
                }
            }

            catch (Exception ex)
            {
                return ex.Message;
            }

            return "0";
        }

        public static string FileArchive(string file)
        {
            string frFilesFolder = ConfigurationManager.AppSettings["FR_Folder"];
            string frArchiveFolder = ConfigurationManager.AppSettings["FR_Folder"] + "archive";

            try
            {
                string fileName = file;
                string sourcePath = @frFilesFolder;
                string targetPath = @frArchiveFolder;

                string sourceFile = Path.Combine(sourcePath, fileName);
                string destFile = Path.Combine(targetPath, fileName);

                if (!Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                }

                File.Move(sourceFile, destFile);
            }

            catch (Exception ex)
            {
                return ex.Message;
            }

            WriteLog(file + " archived in: " + frArchiveFolder + Environment.NewLine);
            return "0";
        }


        public static string StrLeft(string s, int len)
        {
            // Left function
            string start = s.Substring(0, len); 
            return start;
        }

        public static string StrRev(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        public static string ValidateDateStr(string date)
        {
            var dt_result = "";

            try
            {
                dt_result = DateTime.ParseExact(date, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "1";
            }

            finally
            {
                // finally
            }

            return dt_result;
        }

        public static string CheckIfNumeric(string input)
        {
            if (IsNumeric(input) == true)
            {
                //Console.WriteLine(input + " is numeric.");
                return "0";
            }
            else
            {
                //Console.WriteLine("CheckIfNumeric(" + input + ")" + " is NOT numeric.");
                return "1";
            }
        }

        public static bool IsNumeric(string input)
        {
            return Regex.IsMatch(input, @"^\d+$");
        }

        public static string CreateArchiveSuccess(string ArchivePath)
        {
            try
            {
                Directory.CreateDirectory(ArchivePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return ex.Message;
            }
            return "0";
        }


        public static string SendEmail3(string subject, string body)
        {
            try
            {
                MailMessage message = new MailMessage();
                message.To.Add(ConfigurationManager.AppSettings["FR_EmailRecipients"]);
                message.From = new MailAddress(ConfigurationManager.AppSettings["FR_EmailSender"]);
                message.Subject = subject;
                message.Body = body;
                SmtpClient smtp = new SmtpClient(ConfigurationManager.AppSettings["SMTPServer"]);
                smtp.Send(message);
            }

            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                return ex.Message;
            }
            return "0";
        }

        public static string SendEmail3x(string subject, string body)
        {
            return "1";
        }



    }   // static class Program

}   // namespace FRUpdate