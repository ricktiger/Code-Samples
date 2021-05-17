Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO

Public Class AclUpdate
    Inherits System.Web.UI.Page

    Dim objCommand As New OleDbCommand()
    Dim reader As SqlDataReader

    Dim OrigFileNameOne As String
    Dim FileNameOne As String
    Dim dtFileOne As New DataTable
    Dim foundRows() As Data.DataRow
    Dim FileOneRowCount As New Integer

    Dim IsError As Boolean
    Dim ErrorMsg As New System.Text.StringBuilder

    Dim CodeLog As New System.Text.StringBuilder
    Dim fs As FileStream
    Dim bCodeLog As Boolean = False

    Dim dtTextSanitizer As New DataTable
    Dim SqlOut As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then  'First time
            OutputDiv.Visible = False
            ErrorOutput.Visible = False
        Else    'Is PostBack
            If (bCodeLog) Then
                'Create CodeLog file
                Dim path As String = Server.MapPath("workfiles") & "\CodeLog.txt"
                'Create or overwrite the file.
                fs = File.Create(path)
                'Add text CodeLog file.
                Dim info As Byte() = New UTF8Encoding(True).GetBytes("Code Log..." & vbCrLf)
                fs.Write(info, 0, info.Length)
                'fs.Close()
            End If

            'Start the Files Processing
            Output.Text = String.Empty
            import_for_check()
        End If
    End Sub

    Sub import_for_check()
        Dim exErr As String
        Dim currentDateTime As DateTime = Now()
        Dim YYYYMMDDHHMMSS As String
        YYYYMMDDHHMMSS = currentDateTime.ToString("yyyyMMddhhss")

        'File One
        If FileOne.HasFile = True Then
            ViewState("newexcel") = False
            'Save off Original FileOne Name
            OrigFileNameOne = FileOne.FileName.ToString
            If Right(FileOne.FileName.ToString, 4) = "xlsx" Then
                ViewState("newexcel") = True
                'FileNameOne = Server.MapPath("workfiles") & "\" & getrandomfile() & ".xlsx"
                FileNameOne = Server.MapPath("workfiles") & "\" & OrigFileNameOne.Replace(".xlsx", "_Update_" & YYYYMMDDHHMMSS & ".xlsx")
            Else
                ViewState("newexcel") = True
                'FileNameOne = Server.MapPath("workfiles") & "\" & getrandomfile() & ".xls"
                FileNameOne = Server.MapPath("workfiles") & "\" & OrigFileNameOne.Replace(".xls", "_Update_" & YYYYMMDDHHMMSS & ".xls")
            End If

            FileOne.SaveAs(FileNameOne)
            ViewState("FileNameOne") = FileNameOne

            Try
                'objCommand = ExcelConnection(OrigFileNameOne, FileNameOne, ViewState("newexcel"), FileType)
                objCommand = ExcelConnection2(OrigFileNameOne, FileNameOne, ViewState("newexcel"), "ACL")
            Catch ex As Exception
                IsError = True
                exErr = ex.Message
                Output.Text = Output.Text & ex.Message & vbCrLf
                ErrorMsg.Append("The file either Is Not an XLS/XLSX file Or the first sheet is not named Sheet1 Or " & ex.Message)
            Finally
                'Finally
            End Try

            Try
                If IsError <> True Then
                    'Get FileOne
                    dtFileOne.Dispose()

                    'dtFileOne Fields
                    dtFileOne.Columns.Add("Reference")
                    dtFileOne.Columns.Add("Name of Individual Or Entity")
                    dtFileOne.Columns.Add("Type")
                    dtFileOne.Columns.Add("Name Type")
                    dtFileOne.Columns.Add("Date of Birth")
                    dtFileOne.Columns.Add("Place of Birth")
                    dtFileOne.Columns.Add("Citizenship")
                    dtFileOne.Columns.Add("Address")
                    dtFileOne.Columns.Add("Additional Information")
                    dtFileOne.Columns.Add("Listing Information")
                    dtFileOne.Columns.Add("Committees")
                    dtFileOne.Columns.Add("Control Date")

                    'dtTextSanitizer Fields
                    dtTextSanitizer.Columns.Add("From")
                    dtTextSanitizer.Columns.Add("To")

                    'Fill dtTextSanitizer dt
                    Dim strSql As String = "SELECT [From],[To] FROM dbo.SanitizeTextPair WITH (NOLOCK)"
                    'Using cnn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("myProdDbConnection").ConnectionString)
                    Using cnn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("myLocalDbConnection").ConnectionString)
                        cnn.Open()
                        Using dad As New SqlDataAdapter(strSql, cnn)
                            dad.Fill(dtTextSanitizer)
                        End Using
                        cnn.Close()
                    End Using

                    Dim reader As OleDbDataReader
                    reader = objCommand.ExecuteReader()     'Errors here if Sheet name is not good, or check FileType
                    While reader.Read()
                        Dim dr As DataRow = dtFileOne.NewRow
                        dr.Item("Reference") = Trim(FixNull(reader(0)))
                        dr.Item("Name of Individual or Entity") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(1))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Type") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(2))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Name Type") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(3))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Date of Birth") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(4))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Place of Birth") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(5))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Citizenship") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(6))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Address") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(7))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Additional Information") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(8))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Listing Information") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(9))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Committees") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(reader(10))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dr.Item("Control Date") = TextSanitize(FixApos(Replace(Replace(Replace(Trim(FixNull(FixDate(reader(11).ToString()))), vbCr, ""), vbLf, ""), vbCrLf, "")))
                        dtFileOne.Rows.Add(dr)
                    End While
                End If
            Catch ex As Exception
                IsError = True
                ErrorMsg.Append("Error processing File One<br/>" & ex.Message)
            Finally
                'Finally
            End Try
            FileOneRowCount = dtFileOne.Rows.Count()
        End If

        Output.Text = Output.Text & OrigFileNameOne & " Row Count: " & FileOneRowCount & vbCrLf

        foundRows = Nothing

        'Read from dtFileOne
        Dim sqlStr As String = String.Empty
        Dim sqlCount As Integer = 0
        'INSERT INTO Customers (CustomerName, City, Country) VALUES('MytName', 'MyCity', 'MyCountry')
        For Each row As DataRow In dtFileOne.Rows
            sqlStr += "-- Reference=" & row("Reference") & vbCrLf &
                    "INSERT INTO SomeTable (" &
                        "[Name of Individual Or Entity]," &
                        "[Type]," &
                        "[Name Type]," &
                        "[Date of Birth]," &
                        "[Place of Birth]," &
                        "[Citizenship]," &
                        "[Address]," &
                        "[Additional Information]," &
                        "[Listing Information]," &
                        "[Committees]," &
                        "[Control Date]) " &
                    "VALUES(" &
                        "'" & row("Name of Individual or Entity") & "', " &
                        "'" & row("Type") & "', " &
                        "'" & row("Name Type") & "', " &
                        "'" & row("Date of Birth") & "', " &
                        "'" & row("Place of Birth") & "', " &
                        "'" & row("Citizenship") & "', " &
                        "'" & row("Address") & "', " &
                        "'" & row("Additional Information") & "', " &
                        "'" & row("Listing Information") & "', " &
                        "'" & row("Committees") & "', " &
                        "'" & row("Control Date") & "');" & vbCrLf
            sqlCount += 1
            If sqlCount > 29 Then
                sqlStr += "GO;" & vbCrLf
                sqlCount = 0
            End If
        Next row

        'Check for error(s)
        If (IsError) Then
            ErrorOutput.Visible = True
            Errors.Text = ErrorMsg.ToString
        Else
            OutputDiv.Visible = True
        End If

        Dim intdays As Integer = 90
        For Each file As IO.FileInfo In New IO.DirectoryInfo(Server.MapPath("workfiles") & "\").GetFiles("*.xls")
            If (Now - file.CreationTime).Days > intdays Then file.Delete()
        Next

        If (bCodeLog) Then
            fs.Close()
        End If

    End Sub

    Protected Function ExcelConnection(ByVal OrigFileName As String, ByVal myfile As Object, ByVal newexcel As Boolean, ByVal FileType As String) As System.Data.OleDb.OleDbCommand
        OrigFileName = Left(OrigFileName, OrigFileName.Length - 4)
        Dim xConnStr As String = String.Empty
        Select Case newexcel
            Case False
                xConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                                "Data Source=" & myfile & ";" &
                                "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'"
            Case True
                xConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                                 "Data Source=" & myfile & ";" &
                                 "Extended Properties='Excel 12.0;HDR=Yes;imex=1;'"
        End Select
        ' Connect to the Excel Spreadsheet
        ' create excel connection object using the connection string
        Dim objXConn As New System.Data.OleDb.OleDbConnection(xConnStr)
        objXConn.Open()

        ' use a SQL Select command to retrieve the data from the Excel Spreadsheet
        ' the "table name" is the name of the worksheet within the spreadsheet
        ' in this case, the worksheet name is "Sheet1" and is expressed as: [Sheet1$]

        Dim cell_selection As String = String.Empty
        'Dim sheet_name As String = String.Empty
        Select Case FileType
            Case "ACL"
                'sheet_name = "SELECT * FROM [" & myfile & "$A2:L999]"     'A2
                'cell_selection = "SELECT * FROM [ACL Feb_5_2018$A2:L999]"     'A2
                cell_selection = "SELECT * FROM [Sheet1$A1:L7000]"     'A1

        End Select

        Dim objCommand As New System.Data.OleDb.OleDbCommand(cell_selection, objXConn)
        Return objCommand

    End Function

    Protected Function ExcelConnection2(ByVal OrigFileName As String, ByVal myfile As String, ByVal newexcel As Boolean, ByVal FileType As String) As System.Data.OleDb.OleDbCommand
        Dim cn As New OleDbConnection
        Dim cm As New OleDbCommand

        Select Case newexcel
            Case False
                cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                                "Data Source=" & myfile & ";" &
                                "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'"
            Case True
                cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                                 "Data Source=" & myfile & ";" &
                                 "Extended Properties='Excel 12.0;HDR=Yes;imex=1;'"
        End Select

        cn.Open()

        With cm
            .Connection = cn
            .CommandText = "SELECT * FROM [Sheet1$A1:L8000]"
            .ExecuteNonQuery()

            Return cm
        End With

        cn.Close()
    End Function

    Public Shared Function FixNull(ByVal dbvalue) As String
        If dbvalue Is DBNull.Value Then
            Return ""
        ElseIf dbvalue Is Nothing Then
            Return String.Empty
        Else
            'NOTE: This will cast value to string if
            'it isn't a string.
            Return dbvalue.ToString
        End If
    End Function
    Public Shared Function FixDate(ByVal sdate) As String
        If sdate.ToString.Length > 0 Then
            Dim sVal As String = sdate.ToString.Replace(" 12:00:00 AM", "")
            Return sVal
        Else
            Return sdate
        End If
    End Function
    Public Shared Function FixApos(ByVal sApos) As String
        If sApos <> Nothing And Not IsDBNull(sApos) Then
            Dim sVal As String = sApos
            sVal = sVal.ToString.Replace("'", "''")
            sVal = sVal.ToString.Replace("‘", "''")
            sVal = sVal.ToString.Replace("’", "''")

            Return sVal
        Else
            If IsDBNull(sApos) Then sApos = String.Empty
            Return sApos
        End If
    End Function

    Public Function ConvertToUnicode(ByVal unicodeString As String) As String
        'Create two different encodings.
        Dim ascii As Encoding = Encoding.ASCII
        Dim unicode As Encoding = Encoding.Unicode

        'Convert the string into a byte array.
        Dim unicodeBytes As Byte() = unicode.GetBytes(unicodeString)

        'Perform the conversion from one encoding to the other.
        Dim asciiBytes As Byte() = Encoding.Convert(unicode, ascii, unicodeBytes)

        'Convert the new byte array into a char array and then into a string.
        Dim asciiChars(ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
        ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)

        Return asciiChars
    End Function

    Public Function TextSanitize(ByVal strInChars As String) As String
        'strInChars = "a-b,c+d*e"   'For unit testing
        Dim strOutChars As String = String.Empty
        If strInChars <> Nothing And Not IsDBNull(strInChars) Then
            Try
                'Step through the strInChars string
                For i As Integer = 0 To strInChars.Length - 1
                    'Find matching char in dtTextSanitizer.From
                    Dim sel As String = "From = '" & strInChars(i) & "'"
                    If strInChars(i) = "'" Then sel += "'"
                    foundRows = dtTextSanitizer.Select(sel)
                    If (foundRows.Length > 0 And strInChars(i) <> "'") Then
                        For count = 0 To foundRows.Length - 1
                            If (strInChars(i) = foundRows(count).ItemArray(0)) Then
                                'Set the char to the To value
                                strOutChars += foundRows(count).ItemArray(1)
                            End If
                        Next
                    Else
                        'No Hit, Keep the char from input
                        strOutChars += strInChars(i)
                    End If
                Next
            Catch ex As Exception
                IsError = True
                ErrorMsg.Append("Error processing File One<br/>" & ex.Message)
            End Try
        End If

        Return strOutChars

    End Function
End Class