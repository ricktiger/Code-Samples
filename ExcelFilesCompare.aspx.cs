using System;
using System.Linq;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.UI.WebControls;

namespace ExcelFilesCompare
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


    public partial class ExcelFilesCompare : System.Web.UI.Page
    {
        public const string cCrLf = "\r\n";
        public bool bHasError;
        public string KeyField = String.Empty;
        public string[] arrKeyField;
        public int KeyFieldCount;
        string ColumnsNumsOfKeyFields = String.Empty;

        public string File1OrigName = String.Empty;
        public DataTable dt1 = new DataTable();
        public string dtColumns1 = String.Empty;
        public int dtColumns1Count = 0;
        int iCount1;
        public string arrFile1KeyFieldCount = String.Empty;
        public string sFile1KeyFieldValues = String.Empty;
        public DataTable dtRowsIn1NotIn2 = new DataTable();

        public string File2OrigName = String.Empty;
        public DataTable dt2 = new DataTable();
        public string dtColumns2 = String.Empty;
        public int dtColumns2Count = 0;
        int iCount2;
        public string arrFile2KeyFieldCount = String.Empty;
        public string sFile2KeyFieldValues = String.Empty;
        public string sFile1File2MatchedRows = String.Empty;
        public DataTable dtRowsIn2NotIn1 = new DataTable();

        public DataTable dt1dt2MatchedRows = new DataTable();
        public DataTable dt2dt1MatchedRows = new DataTable();

        public DataTable dtRowsColsToChgColor = new DataTable();
        public DataTable dtFileOneTwoMrg = new DataTable();

        public DataRow[] foundRows1;
        public DataRow[] foundRows2;
        public DataRow[] foundRows3;
        public DataRow[] foundRows4;

        Dictionary<int, string> kvpKf = new Dictionary<int, string>();  // Not currently used
        List<int> liKf = new List<int>();

        private Excel.Worksheet worksheet;

        protected void Page_Load(object sender, EventArgs e)
        {
            OutputDiv.Visible = false;
            //WriteToLog("Page_Load");
        }

        protected void UploadButton_Click(Object sender, EventArgs e)
        {
            // *************************************************************************
            // Process FileUpload1
            // *************************************************************************
            if (FileUpload1.HasFile)
            {
                //WriteToLog("UploadButton_Click-FileUpload1.HasFile");
                Int64 t1 = GetTime();
                string filePath1 = string.Empty;

                string path1 = System.Web.HttpContext.Current.Server.MapPath("~/workfiles/");
                if (!Directory.Exists(path1))
                {
                    Directory.CreateDirectory(path1);
                }
                // Get file path
                filePath1 = path1 + Path.GetFileName(FileUpload1.FileName);
                // Get file extenstion
                string extension = Path.GetExtension(FileUpload1.FileName);
                // Get file Name
                string File1OrigName = FileUpload1.FileName;
                // Get KeyField or Fields
                KeyField = Request.Form["KeyField"].Replace(", ", ",");  // Get one or more field names (comma separated)
                arrKeyField = KeyField.Split(',');     // Split the files names into a string array
                KeyFieldCount = arrKeyField.Length;     // Get the number of key files entered
                // Save file on "workfiles" folder of project
                FileUpload1.SaveAs(filePath1);

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
                conString = string.Format(conString, filePath1);

                // Use OldDb to read excel
                using (OleDbConnection connExcel = new OleDbConnection(conString))
                using (OleDbCommand cmdExcel = new OleDbCommand())
                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                {
                    cmdExcel.Connection = connExcel;

                    // Get the name of First Sheet
                    connExcel.Open();
                    DataTable dtExcelSchema;
                    dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

                    // Get the Column Names
                    DataTable dtCols = connExcel.GetSchema("Columns");
                    DataView dvCols = new DataView(dtCols);
                    dvCols.RowFilter = "TABLE_NAME = '" + sheetName + "'";
                    dvCols.Sort = "ORDINAL_POSITION ASC";

                    // Read Column Names and stack them up in a string
                    int ColCount = 0;
                    foreach (DataRowView rowView in dvCols)
                    {
                        DataRow row = rowView.Row;
                        string rvcn = rowView["COLUMN_NAME"].ToString();    // Get the Column Name
                        // Make sure the Column Name is NOT F(n), ie; F11, F12
                        if (rvcn.Left(2) != "F1")
                        {
                            dtColumns1 += rowView["COLUMN_NAME"] + ",";     // Column Name NOT F(n), so save it

                            // Get a Column Name
                            string ColName = rowView["COLUMN_NAME"].ToString();

                            // Loop through the Key Field(s) Column Name(s) and Compare Column Name to One or more KeyFields
                            for (int i = 0; i < arrKeyField.Length; i++)
                            {
                                // If Column Name matches the Key Field(s)
                                if (ColName == arrKeyField[i])
                                {
                                    ColumnsNumsOfKeyFields += ColCount.ToString() + ",";
                                    //liKf.Add(i);
                                    liKf.Add(ColCount);
                                }
                            }
                            ColCount++;
                        }
                    }
                    dtColumns1Count = ColCount;
                    dtColumns1 = dtColumns1.Left(dtColumns1.Length - 1);
                    connExcel.Close();

                    // Insure that Key Field(s) are listed in Column Names
                    if (!CheckKFsInColNames(arrKeyField, dtColumns1))
                    {
                        bHasError = true;
                        Output.Text = Output.Text + "One or more Key Field is Not Found in Available File 1 Column Names." + cCrLf;
                    }

                    var ColumnNames = dtColumns1.Split(Convert.ToChar(","));
                    // Iterate through each Column Name and Define it
                    foreach (var ColumnName in ColumnNames)
                    {
                        dt1.Columns.Add(ColumnName, Type.GetType("System.String"));
                    }

                    // Read Data from First Sheet into dt
                    connExcel.Open();
                    cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "A1:" + NumToLetter(ColCount) + "30000]";
                    odaExcel.SelectCommand = cmdExcel;
                    odaExcel.Fill(dt1);
                    iCount1 = dt1.Rows.Count;
                    connExcel.Close();

                    // Sort dt1 by KeyField(s)
                    if (CheckKFsInColNames(arrKeyField, dtColumns1))
                    {
                        DataView dv1 = dt1.DefaultView;
                        dv1.Sort = KeyField;
                        dt1 = dv1.ToTable();
                    }

                    // Set KeyFields with empty values to some unused character
                    for (int i = 0; i < arrKeyField.Length; i++)
                    {
                        DataRow[] krows = dt1.Select("[" + arrKeyField[i] + "] IS NULL");

                        for (int k = 0; k < krows.Length; k++)
                        {
                            krows[k][arrKeyField[i]] = "^";
                        }
                    }

                    // Read & Save the values for KeyField from dt1
                    foreach (DataRow row in dt1.Rows)
                    {
                        string keyval = string.Empty;
                        string KeyValueHolder = string.Empty;
                        for (int i = 0; i < arrKeyField.Length; i++)    //Read through the KeyField(s) array
                        {
                            string val = FixNull(row[row.Table.Columns[arrKeyField[i]].Ordinal]);   // Get the data Value of a KeyField
                            if (val.Length > 0)
                            {
                                keyval = val;
                                keyval = keyval.Replace("'", "''");
                                KeyValueHolder += keyval;
                            }
                        }
                        sFile1KeyFieldValues += "'" + KeyValueHolder + "',";
                    }
                    // Trim off trailing comma
                    sFile1KeyFieldValues = sFile1KeyFieldValues.Left(sFile1KeyFieldValues.Length - 1);
                }

                // Clean up (delete) files in workfiles folder older than 61 days
                DateTime CutOffDate = DateTime.Now.AddDays(-61);
                DirectoryInfo di = new DirectoryInfo(path1);
                FileInfo[] fi = di.GetFiles();

                for (int i = 0; i < fi.Length; i++)
                {
                    if (fi[i].LastWriteTime < CutOffDate)
                    {
                        File.Delete(fi[i].FullName);
                    }
                }

            }
            else
            {
                // You did not specify a File1 to upload
                bHasError = true;
                Output.Text = Output.Text + "No File 1 specified for Upload." + cCrLf;
            }
            //WriteToLog("End - FileUpload1.HasFile");


            // *************************************************************************
            // Process FileUpload2
            // *************************************************************************
            if (FileUpload2.HasFile)
            {
                string filePath2 = string.Empty;

                string path2 = System.Web.HttpContext.Current.Server.MapPath("~/workfiles/");
                if (!Directory.Exists(path2))
                {
                    Directory.CreateDirectory(path2);
                }
                // Get file path
                filePath2 = path2 + Path.GetFileName(FileUpload2.FileName);
                // Get file extenstion
                string extension = Path.GetExtension(FileUpload2.FileName);
                // Get file Name
                string File2OrigName = FileUpload2.FileName;
                // Get KeyField or Fields
                KeyField = Request.Form["KeyField"].Replace(", ", ",");  // Get one or more field names (comma separated)
                arrKeyField = KeyField.Split(',');     // Split the files names into a string array
                KeyFieldCount = arrKeyField.Length;     // Get the number of key files entered
                // Save file on "workfiles" folder of project
                FileUpload2.SaveAs(filePath2);

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
                conString = string.Format(conString, filePath2);

                // Use OldDb to read excel
                using (OleDbConnection connExcel = new OleDbConnection(conString))
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

                        // Get the Column Names
                        DataTable dtCols = connExcel.GetSchema("Columns");
                        DataView dvCols = new DataView(dtCols);
                        dvCols.RowFilter = "TABLE_NAME = '" + sheetName + "'";
                        dvCols.Sort = "ORDINAL_POSITION ASC";

                        // Read Column Names and stack them up in a string
                        int ColCount = 0;
                        string ColumnsNumsOfKeyFields = String.Empty;
                        foreach (DataRowView rowView in dvCols)
                        {
                            DataRow row = rowView.Row;
                            string rvcn = rowView["COLUMN_NAME"].ToString();    // Get the Column Name
                                                                                // Make sure the Column Name is NOT F(n), ie; F11, F12
                                                                                //string rvcn_right = rvcn.Substring(rvcn.Length - 1);
                            if (rvcn.Left(2) != "F1")
                            {
                                dtColumns2 += rowView["COLUMN_NAME"] + ",";     // Column Name NOT F(n), so save it

                                // Get a Column Name
                                string ColName = rowView["COLUMN_NAME"].ToString();

                                // Loop through the Key Field(s) Column Name(s) and Compare Column Name to One or more KeyFields
                                for (int i = 0; i < arrKeyField.Length; i++)
                                {
                                    // If Column Name matches the Key Field(s)
                                    if (ColName == arrKeyField[i])
                                    {
                                        ColumnsNumsOfKeyFields += ColCount.ToString() + ",";
                                    }
                                }
                                ColCount++;
                            }
                        }
                        dtColumns2Count = ColCount;
                        dtColumns2 = dtColumns2.Left(dtColumns2.Length - 1);
                        // Trim off trailing comma
                        ColumnsNumsOfKeyFields = ColumnsNumsOfKeyFields.Left(ColumnsNumsOfKeyFields.Length - 1);

                        connExcel.Close();

                        // Insure that Key Field(s) are listed in Column Names
                        if (!CheckKFsInColNames(arrKeyField, dtColumns2))
                        {
                            bHasError = true;
                            Output.Text = Output.Text + "One or more Key Field is Not Found in Available File 2 Column Names." + cCrLf;
                        }

                        var ColumnNames = dtColumns2.Split(Convert.ToChar(","));
                        // Iterate through each Column Name and Define it
                        foreach (var ColumnName in ColumnNames)
                        {
                            dt2.Columns.Add(ColumnName, Type.GetType("System.String"));
                        }

                        // Read Data from First Sheet into dt2
                        connExcel.Open();
                        cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "A1:" + NumToLetter(ColCount) + "30000]";
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dt2);
                        iCount2 = dt2.Rows.Count;
                        connExcel.Close();

                        // Sort dt2 by KeyField(s)
                        if (CheckKFsInColNames(arrKeyField, dtColumns2))
                        {
                            DataView dv2 = dt2.DefaultView;
                            dv2.Sort = KeyField;
                            dt2 = dv2.ToTable();
                        }

                        // Set KeyFields with empty values to some unused character
                        for (int i = 0; i < arrKeyField.Length; i++)
                        {
                            DataRow[] krows = dt2.Select("[" + arrKeyField[i] + "] IS NULL");

                            for (int k = 0; k < krows.Length; k++)
                            {
                                krows[k][arrKeyField[i]] = "^";
                            }
                        }

                        // Read & Save the values for KeyField from dt2
                        foreach (DataRow row in dt2.Rows)
                        {
                            string keyval = string.Empty;
                            string KeyValueHolder = string.Empty;
                            for (int i = 0; i < arrKeyField.Length; i++)    //Read through the KeyField(s) array
                            {
                                string val = FixNull(row[row.Table.Columns[arrKeyField[i]].Ordinal]);   // Get the data Value of a KeyField
                                if (val.Length > 0)
                                {
                                    keyval = val;
                                    keyval = keyval.Replace("'", "''");
                                    KeyValueHolder += keyval;
                                }
                            }
                            sFile2KeyFieldValues += "'" + KeyValueHolder + "',";

                        }
                        // Trim off trailing comma
                        sFile2KeyFieldValues = sFile2KeyFieldValues.Left(sFile2KeyFieldValues.Length - 1);
                    }

                }

                OutputDiv.Visible = true;
            }
            else
            {
                // You did not specify a File2 to upload
                bHasError = true;
                Output.Text = Output.Text + "No File 2 specified for Upload." + cCrLf;
            }

            //WriteToLog("End - FileUpload2.HasFile");

            Output.Text = Output.Text + "File 1 Column Names: " + dtColumns1 + cCrLf;
            Output.Text = Output.Text + "File 1 Record Count: " + iCount1 + cCrLf + cCrLf;

            Output.Text = Output.Text + "File 2 Column Names: " + dtColumns2 + cCrLf;
            Output.Text = Output.Text + "File 2 Record Count: " + iCount2 + cCrLf + cCrLf;

            if (dtColumns1 != dtColumns2)
            {
                bHasError = true;
                Output.Text = Output.Text + "File 1 and File 2 Column Names DO NOT MATCH." + cCrLf;
            }


            // Timer: With ~18K records in each file: 1 minute to get to here

            if (!bHasError)
            {
                //WriteToLog("Start Compares");
                // Start Compares
                // ********************************
                // Define dtRowsIn1NotIn2 datatable
                // ********************************
                var ColumnNames = dtColumns1.Split(Convert.ToChar(","));
                // Iterate through each Column Name and Define it
                foreach (var ColumnName in ColumnNames)
                {
                    dtRowsIn1NotIn2.Columns.Add(ColumnName, Type.GetType("System.String"));
                }

                // Find Rows in dt1 that are not in dt2 using KeyField(s)
                //([Last Name]+[First Name]) NOT IN
                string query = GetKFsForSelect(arrKeyField) + " NOT IN(" + sFile2KeyFieldValues.Left(sFile2KeyFieldValues.Length - 1) + "')";
                foundRows1 = dt1.Select(query);
                if (foundRows1.Length != 0) dtRowsIn1NotIn2 = foundRows1.CopyToDataTable();

                // Set KeyFields with some unused character back to empty values
                for (int i = 0; i < arrKeyField.Length; i++)
                {
                    DataRow[] krows = dtRowsIn1NotIn2.Select("[" + arrKeyField[i] + "] = '^'");

                    for (int k = 0; k < krows.Length; k++)
                    {
                        krows[k][arrKeyField[i]] = string.Empty;
                    }
                }

                // bind datatable with GridView
                GridView1.DataSource = dtRowsIn1NotIn2;
                GridView1.DataBind();

                // Write dtRowsIn1NotIn2 out as CSV
                //dtRowsIn1NotIn2.WriteToCsvFile(System.Web.HttpContext.Current.Server.MapPath("~/workfiles/") + Path.GetFileName(FileUpload1.FileName) + "_RowsInFile1NotInFile2.csv");


                // ********************************
                // Define dtRowsIn2NotIn1 datatable
                // ********************************
                ColumnNames = dtColumns2.Split(Convert.ToChar(","));
                // Iterate through each Column Name and Define it
                foreach (var ColumnName in ColumnNames)
                {
                    dtRowsIn2NotIn1.Columns.Add(ColumnName, Type.GetType("System.String"));
                }

                // Find Rows in dt2 that are not in dt1 using KeyField
                query = GetKFsForSelect(arrKeyField) + " NOT IN(" + sFile1KeyFieldValues.Left(sFile1KeyFieldValues.Length - 1) + "')";
                foundRows2 = dt2.Select(query);
                if (foundRows2.Length != 0) dtRowsIn2NotIn1 = foundRows2.CopyToDataTable();

                // Set KeyFields with some unused character back to empty values
                for (int i = 0; i < arrKeyField.Length; i++)
                {
                    DataRow[] krows = dtRowsIn2NotIn1.Select("[" + arrKeyField[i] + "] = '^'");

                    for (int k = 0; k < krows.Length; k++)
                    {
                        krows[k][arrKeyField[i]] = string.Empty;
                    }
                }

                // bind datatable with GridView
                GridView2.DataSource = dtRowsIn2NotIn1;
                GridView2.DataBind();


                // ***********************************
                // Define dtRowsIn2DiffFrom1 datatable
                // ***********************************
                //WriteToLog("Start Define dtRowsIn2DiffFrom1");
                ColumnNames = dtColumns2.Split(Convert.ToChar(","));
                // Iterate through each Column Name
                foreach (var ColumnName in ColumnNames)
                {
                    dt1dt2MatchedRows.Columns.Add(ColumnName, Type.GetType("System.String"));
                }

                // Timer: With ~18K records in each file: 1 minute 30 seconds to get to here

                // Read & Save the values for KeyField from dt1dt2MatchedRows
                dtRowsColsToChgColor.Columns.Add("KeyFieldValue", Type.GetType("System.String"));
                dtRowsColsToChgColor.Columns.Add("Column", Type.GetType("System.Int32"));

                int ii = 0;
                int fr = 0;
                string frLast = "xx";

                foreach (DataRow rowdt1 in dt1.Rows)    // Read through dt1 for KeyField(s) matches found in dt2
                {
                    int index2 = 0;
                    // Build query w/ func to get KeyField name(s) & func to get actual KeyField(s) data values
                    query = GetKFsForSelect(arrKeyField) + "=" + GetKFsValuesForSelect(rowdt1, arrKeyField);
                    // Query will look something like this: ([Last Name]+[First Name])='AbadirMaher'
                    foundRows3 = dt2.Select(query);
                    // Get the data Value of a KeyField
                    string keyvalueholder = string.Empty;
                    for (int l = 0; l < liKf.Count; l++)
                    {
                        keyvalueholder += FixNull(rowdt1[liKf[l]]).Replace("'", "''");
                    }
                    // See if foundRows3
                    if (foundRows3.Length > 0)  // If we have query hits
                    {
                        if (foundRows3.Length > 1) index2 = 1;
                        //Output.Text = Output.Text + "KeyFieldValue: " + keyvalueholder + cCrLf;
                        //Output.Text = Output.Text + "foundRows3.Length: " + foundRows3.Length + cCrLf +"-----------------------------" + cCrLf;
                        for (fr = index2; fr <= foundRows3.GetUpperBound(0); fr++)
                        {
                            //if (frLast == keyvalueholder && foundRows3.GetUpperBound(0) < 3 && foundRows3.GetUpperBound(0) > fr) fr++;
                            if (frLast == keyvalueholder && foundRows3.GetUpperBound(0) < foundRows3.Length && foundRows3.GetUpperBound(0) > fr) fr++;
                            var array1 = rowdt1.ItemArray;          // Get record array from dt1
                            var array2 = foundRows3[fr].ItemArray;  // Get record array from foundRows3

                            // Read all Columns
                            bool b1 = false;
                            for (ii = 0; ii <= array1.Length - 1; ii++)     // Process array1 record (from dt1)
                            {
                                // No compares on KeyFields - Yes on Diffs(!=) between array1 (dt1) & array2 (foundRows3)
                                //if (!liKf.Contains(ii) && array1[ii].ToString() != array2[ii].ToString())
                                if (!liKf.Contains(ii) && !string.Equals(array1[ii].ToString(), array2[ii].ToString(), StringComparison.OrdinalIgnoreCase))
                                {
                                    dtRowsColsToChgColor.Rows.Add(keyvalueholder, ii);  // Add keyfields data & the index of column that is !=
                                    b1 = true;
                                }
                            }
                            frLast = keyvalueholder;    // Save off last key value
                            if (!b1) break;
                        }
                    }
                } // foreach (DataRow rowdt1 in dt1.Rows)    // Read through dt1 for KeyFiled(s) matches found in dt2

                // Timer: With ~18K records in each file: 20 minutes to get to here

                // ***********************************
                // Define dtFileOneTwoMrg datatable
                // ***********************************
                //WriteToLog("Start Define dtFileOneTwoMrg");
                ColumnNames = dtColumns2.Split(Convert.ToChar(","));
                // Iterate through each Column Name
                foreach (var ColumnName in ColumnNames)
                {
                    dtFileOneTwoMrg.Columns.Add(ColumnName, Type.GetType("System.String"));
                }

                // Read from dtRowsColsToChgColor for KeyFieldValue 
                string savedKeyFieldValue = "x";
                foreach (DataRow rowCC in dtRowsColsToChgColor.Rows)
                {
                    if (savedKeyFieldValue != (string)rowCC["KeyFieldValue"])
                    {
                        // Start Find & Add record from dt1 and add row to dtFileOneTwoMrg
                        query = GetKFsForSelect(arrKeyField) + "='" + rowCC["KeyFieldValue"] + "'";
                        // Query will look something like this: ([Last Name]+[First Name])='AaronBenjamin'
                        foundRows1 = dt1.Select(query);
                        if (foundRows1.Length > 0)  // If we have query hits
                        // Read from foundRows1 (dt1) for record hit
                        {
                            var array1 = foundRows1[0].ItemArray;        // Get record array from foundRows1 (dt1)
                            DataRow row = dtFileOneTwoMrg.NewRow();

                            for (ii = 0; ii <= array1.Length - 1; ii++)
                            {
                                row[ii] = array1[ii].ToString();
                            }
                            dtFileOneTwoMrg.Rows.Add(row);  // Add dt1.KeyFieldValue row to dtFileOneTwoMrg
                        }
                        // End Find & Add record from dt1 and add row to dtFileOneTwoMrg


                        // Start Find & Add record from dt2 and add row to dtFileOneTwoMrg
                        // Use the same query as above
                        // Query will look something like this: ([Last Name]+[First Name])='AaronBenjamin'
                        foundRows2 = dt2.Select(query);
                        // Read from foundRows2 (dt2) for record hit
                        if (foundRows2.Length > 0)  // If we have query hits
                        {
                            var array2 = foundRows2[0].ItemArray;
                            if (foundRows2.Length > 1) array2 = foundRows2[1].ItemArray;    // Get record array from foundRows2 (dt2)
                            DataRow row = dtFileOneTwoMrg.NewRow();

                            for (ii = 0; ii <= array2.Length - 1; ii++)
                            {
                                row[ii] = array2[ii].ToString();
                            }
                            dtFileOneTwoMrg.Rows.Add(row);  // Add dt2.KeyFieldValue row to dtFileOneTwoMrg
                        }
                        // End Find & Add record from dt2 and add row to dtFileOneTwoMrg
                        savedKeyFieldValue = (string)rowCC["KeyFieldValue"];
                    }
                }

                // Set KeyFields with some unused character back to empty values
                for (int i = 0; i < arrKeyField.Length; i++)
                {
                    DataRow[] krows = dtFileOneTwoMrg.Select("[" + arrKeyField[i] + "] = '^'");

                    for (int k = 0; k < krows.Length; k++)
                    {
                        krows[k][arrKeyField[i]] = string.Empty;
                    }
                }

                // Bind datatable with GridView
                //WriteToLog("Start GridView3.DataBind");

                GridView3.DataSource = dtFileOneTwoMrg;
                GridView3.DataBind();

                // Set the ForeColors for before and after data diffs
                int GV3rowCount = 0;
                savedKeyFieldValue = "x";   // Default savedKeyFieldValue to "x"

                foreach (DataRow rowCC in dtRowsColsToChgColor.Rows)
                {
                    int cc = (int)rowCC["Column"];  // Save off cc (Cell Count)

                    if (savedKeyFieldValue != (string)rowCC["KeyFieldValue"])   // Process unique KeyFieldValue values
                    {
                        GridView3.Rows[GV3rowCount].Cells[cc].ForeColor = Color.Green;
                        GridView3.Rows[GV3rowCount + 1].Cells[cc].ForeColor = Color.Red;
                        GV3rowCount += 2;
                    }
                    else
                    {
                        GridView3.Rows[GV3rowCount - 2].Cells[cc].ForeColor = Color.Green;
                        GridView3.Rows[GV3rowCount - 1].Cells[cc].ForeColor = Color.Red;
                    }
                    // Save off KeyFieldValue for next iteration
                    savedKeyFieldValue = (string)rowCC["KeyFieldValue"];
                }

            }

            // Set up for the GridViews exports to Excel
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);




            // Call the function (ExportToExcel) to export each Gridview to Execel
            ExportToExcel(app, workbook, GridView3, "RowsInFile2DiffThanFile1", 2);
            ExportToExcel(app, workbook, GridView2, "RowsInFile2NotInFile1", 1);
            ExportToExcel(app, workbook, GridView1, "RowsInFile1NotInFile2", 0);

            string url = SaveExportToExcel(workbook);

            url_label.Text = "<a href='" + url + "'" + "target='_blank' />Download Excel Results</a>";
        }




        //*********************************************************************
        // Misc Functions
        //*********************************************************************

        // https://stackoverflow.com/questions/33115067/how-do-i-format-my-export-to-excel-workbook-in-microsoft-office-interop-excel
        public void ExportToExcel(Microsoft.Office.Interop.Excel._Application app, Microsoft.Office.Interop.Excel._Workbook workbook, GridView gridview, string SheetName, int sheetid)
        {
            // see the excel sheet behind the program
            app.Visible = false;

            // get the reference of first sheet. By default its name is Sheet1
            worksheet = (Excel.Worksheet)workbook.Worksheets.Add();

            // changing the name of active sheet
            worksheet.Name = SheetName;

            int gridViewCellCount = gridview.Rows[0].Cells.Count;
            // string array to hold grid view column names.
            string[] columnNames = new string[gridViewCellCount];

            // gridview.Rows.Count
            int gridViewRowCount = gridview.Rows.Count;

            for (int i = 0; i < gridViewCellCount; i++)
            {
                columnNames[i] = ((System.Web.UI.WebControls.DataControlFieldCell)(gridview.Rows[0].Cells[i])).ContainingField.HeaderText;
            }

            int iCol = 1;
            foreach (var name in columnNames)
            {
                worksheet.Cells[1, iCol] = name;
                iCol++;
            }

            // storing Each row and column value to excel sheet
            for (int i = 0; i < gridViewRowCount; i++)
            {
                for (int j = 0; j < gridViewCellCount; j++)
                {
                    string cv = gridview.Rows[i].Cells[j].Text;
                    if (gridview.Rows[i].Cells[j].Text != "&nbsp;")
                    {
                        worksheet.Cells[i + 2, j + 1] = gridview.Rows[i].Cells[j].Text;
                    }
                    else
                    {
                        worksheet.Cells[i + 2, j + 1] = "";
                    }
                }
            }
        }

        public string SaveExportToExcel(Microsoft.Office.Interop.Excel._Workbook workbook)
        {
            // Save the application
            DateTime time = DateTime.Now;          // Use current time
            string format = "yyyyMMdd_HHmmss";     // Use this format

            // Get file path
            string filepathname = string.Empty;
            string pathx = System.Web.HttpContext.Current.Server.MapPath("~/workfiles/");
            filepathname = pathx + "Export_" + time.ToString(format) + ".xls";
            string url = "workfiles/Export_" + time.ToString(format) + ".xls";
            workbook.SaveAs(filepathname, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

            return url;
        }

        public static Boolean WriteToLog(string sLog)
        {
            using (StreamWriter sw = new StreamWriter(@System.Web.HttpContext.Current.Server.MapPath("~/workfiles/") + "0Log.txt", true))
            {
                sw.WriteLine(sLog);
            }
            return true;
        }

        public static string GetKFsForSelect(string[] arrKFs)
        {
            string ret_GetKFsForSelect = string.Empty;

            if (arrKFs.Length > 0)
            {
                for (int i = 0; i < arrKFs.Length; i++)     // Make this -> ([Last Name]+[First Name])
                {
                    ret_GetKFsForSelect += "[" + arrKFs[i].ToString() + "]+";
                }
                ret_GetKFsForSelect = "(" + ret_GetKFsForSelect.Left(ret_GetKFsForSelect.Length - 1) + ")";
            }
            return ret_GetKFsForSelect;
        }

        public static string GetKFsValuesForSelect(DataRow dr, string[] arrKFs)
        {
            string ret_GetKFsValuesForSelect = string.Empty;

            if (arrKFs.Length > 0)
            {
                for (int i = 0; i < arrKFs.Length; i++)
                {
                    //string rowValue = dr[arrKFs[i]].ToString();
                    ret_GetKFsValuesForSelect += dr[arrKFs[i]].ToString().Trim().Replace("'", "''");
                    //if (dr[arrKFs[i]].ToString().Trim() == null) ret_GetKFsValuesForSelect += " ";
                }
                ret_GetKFsValuesForSelect = "'" + ret_GetKFsValuesForSelect + "'";
            }
            return ret_GetKFsValuesForSelect;
        }

        public static string GetKFsForSelectRow(string[] arrKFs, string rowname)
        {
            string GetKFsForSelectRow = string.Empty;

            if (arrKFs.Length > 0)
            {
                for (int i = 0; i < arrKFs.Length; i++)     // Make this -> rowdt1["Last Name"] + rowdt1["First Name"]
                {
                    GetKFsForSelectRow += rowname + "[" + arrKFs[i] + "]+";
                }
                GetKFsForSelectRow = "(" + GetKFsForSelectRow.Left(GetKFsForSelectRow.Length - 1) + ")";
            }
            return GetKFsForSelectRow;
        }

        public static Boolean IsIndexNull(DataTable dt, string sIxField)
        {
            DataRow[] foundRows;
            Boolean bIsIndexNull = false;

            string q = "(([" + sIxField + "] IS NULL) OR (LEN([" + sIxField + "])=0))";
            foundRows = dt.Select(q);

            if (foundRows.Length != 0)
            {
                bIsIndexNull = true;
            }

            return bIsIndexNull;
        }

        public static Boolean CheckKFsInColNames(string[] arrKFs, string sColNames)
        {
            Boolean bCheckKFsInColNames = false;
            int iIndexOfarrKFs = 0;

            for (int i = 0; i < arrKFs.Length; i++)
            {
                iIndexOfarrKFs = sColNames.IndexOf(arrKFs[i]);
                if (iIndexOfarrKFs > -1)
                {
                    bCheckKFsInColNames = true;
                }
            }
            return bCheckKFsInColNames;
        }

        public static string FixNull(object dbvalue)
        {
            if (dbvalue == DBNull.Value)
                return "";
            else if (dbvalue == null)
                return string.Empty;
            else
                // NOTE: This will cast value to string if
                // it isn't a string.
                //return dbvalue.ToString().Trim();
                return dbvalue.ToString();
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
                apos_str = apos_str.Replace("'", "''");
            }
            return apos_str.ToString().Trim();
        }

        public static Boolean isAlphaOnly(string strToCheck)
        {
            Regex rg = new Regex(@"[a-zA-Z]");
            return rg.IsMatch(strToCheck);
        }

        public static Boolean isNumericOnly(string strToCheck)
        {
            Regex rg = new Regex(@"[0-1]");
            bool rg_ret = rg.IsMatch(strToCheck);
            return rg_ret;
        }

        private Int64 GetTime()
        {
            Int64 retval = 0;
            var st = new DateTime(1970, 1, 1);
            TimeSpan t = (DateTime.Now.ToUniversalTime() - st);
            retval = (Int64)(t.TotalMilliseconds + 0.5);
            return retval;
        }

        public static String NumToLetter(int num)
        {
            switch (num)
            {
                case 1:
                    return "A";
                case 2:
                    return "B";
                case 3:
                    return "C";
                case 4:
                    return "D";
                case 5:
                    return "E";
                case 6:
                    return "F";
                case 7:
                    return "G";
                case 8:
                    return "H";
                case 9:
                    return "I";
                case 10:
                    return "J";
                case 11:
                    return "K";
                case 12:
                    return "L";
                case 13:
                    return "M";
                case 14:
                    return "N";
                case 15:
                    return "O";
                case 16:
                    return "P";
                case 17:
                    return "Q";
                case 18:
                    return "R";
                case 19:
                    return "S";
                case 20:
                    return "T";
                case 21:
                    return "U";
                case 22:
                    return "V";
                case 23:
                    return "W";
                case 24:
                    return "X";
                case 25:
                    return "Y";
                case 26:
                    return "Z";
                default:
                    return "";
            }
        }

}
}