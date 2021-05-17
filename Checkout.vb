
<!--#include file="_tools\Func_CleanDate.asp" -->
<!--#include file="_tools\Func_SendMailCDOSYS.asp" -->

<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.addHeader "cache-control", "no-cache"
Response.addHeader "cache-control", "no-store"
Response.CacheControl = "no-cache"

Function GetCRCertVoucherTotal(uCustomerIDf)
    GetCRCertVoucherTotal = 0
  	Set rsf = Server.CreateObject("ADODB.Recordset")
	sqlf = "SELECT SUM(PriceExtended) AS CRCertVoucherTotal FROM tbl_InvoiceItems WHERE InvNum = 0 AND CustomerID = '" & uCustomerIDf & "' AND ProductID IN(SELECT ProductID FROM tbl_Product WHERE SKU='Voucher')"
	'response.write("sqlf=" & sqlf & "<br>")
	rsf.Open sqlf, strConnect
	if (NOT(rsf.EOF)) then
        GetCRCertVoucherTotal = rsf("CRCertVoucherTotal") * -1      '(* -1) Make it a positive number
    end if
    rsf.Close
    Set rsf = Nothing
End Function


'Get folder name for possible devlopment
if InStr(LCase(request.servervariables("PATH_INFO")), "staffdev") > 0 then
    strLinkFolder = "StaffDev"
else
    strLinkFolder = "Staff"
end if

'Read for NACVA SSL URLs
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SSLURL,SSLSales FROM tbl_Company WHERE CompanyNum=1"
rs.Open sql, strConnect
if NOT(rs.EOF) then
    Application("eBiz_NACVA_SSLURL") = rs("SSLURL")
    Application("eBiz_NACVA_SSLSales") = rs("SSLSales")
end if
rs.Close
set rs = nothing

Dim RF
Set RF = Request.Form

If RF.Count > 0 Then
    For Each CollectionItem In RF
	    HTML = HTML & CollectionItem & " : <FONT COLOR=""#FF3333"">" & RF(CollectionItem) & "</FONT><br>"
    Next
Else
    HTML = HTML & "<FONT COLOR=""#3333FF"">The Form collection is empty</FONT><br>"
End If
'response.Write(HTML)

'Check for good in-coming CompanyID
if LEN(request.form("cy")) = 0 then
    response.Redirect("Login/")
end if

strToday = MONTH(now) & "/" & DAY(now) & "/" & YEAR(now)

if LEN(request.Form()) > 0 then

    uCustomerID = request.form("cu")

    session("NACVA_ShipMethodID") = request.form("ShipMethodID")

    'Read for CompanyID
	Set rs1 = Server.CreateObject("ADODB.Recordset")
    sql1= "SELECT CompanyID,ActiveMerchant FROM tbl_Company WHERE CompanyID = '" & request.form("cy") & "'"
    rs1.Open sql1, strConnect
	if (NOT(rs1.EOF)) then
        uCompanyID = rs1("CompanyID")
        intMerchantID = rs1("ActiveMerchant")
    else
        uCompanyID = NULL
        intMerchantID = 0
    end if
    rs1.Close
    Set rs1 = Nothing
    
    'Read for MerchantID
	'Set rs1 = Server.CreateObject("ADODB.Recordset")
    'sql1= "SELECT MerchantID FROM tbl_Merchantz WHERE CompanyID = '" & request.form("cy") & "'"
    'rs1.Open sql1, strConnect
	'if (NOT(rs1.EOF)) then
        'intMerchantID = rs1("MerchantID")
    'else
        'intMerchantID = 0
    'end if
    'rs1.Close
    'Set rs1 = Nothing
    
    'Process Affiliate
    intAffiliateID = CInt(6)
    if (NOT(IsNull(request.form("af"))) AND LEN(request.form("af")) > 0) then
        intAffiliateID = CInt(request.form("af"))
    end if
    
    'See if PromoAmount is there
    if LEN(request.Form("PromoAmount")) > 0 then
        dPromoAmount = CSng(request.Form("PromoAmount"))
    else
        dPromoAmount = 0
    end if
    
    'See if PromotionCodeID Amount is there
    intPromotionCodeID = NULL
    if request.Form("PromotionCodeID") > 0 then
        intPromotionCodeID = CInt(request.Form("PromotionCodeID"))
    end if
    
    'See if Addl Discount is there
    if LEN(request.form("AddlDiscAmt")) > 0 then
        dAddlDiscAmt = CSng(request.form("AddlDiscAmt"))
    else
        dAddlDiscAmt = 0
    end if
    
    'Process Tot Shipping
    if LEN(request.form("TotShip")) > 0 then
        dTotShipAmt = CSng(request.form("TotShip"))
    else
        dTotShipAmt = 0
    end if
    
    'Process Possible Credit
    dCreditAmt = 0
    if LEN(request.form("CreditAmt")) > 0 then
        if request.form("ApplyCR") = "1" then
            dCreditAmt = CSng(request.form("CreditAmt"))
        end if            
    end if
    
    sBody = ""
    sBody = sBody & "CustomerID: " & request.form("cu") & vbCRLF
    sBody = sBody & "request.form(""CreditAmt""): " & request.form("CreditAmt") & vbCRLF
    sBody = sBody & "request.form(""ApplyCR""): " & request.form("ApplyCR") & vbCRLF
    sBody = sBody & "dCreditAmt: " & dCreditAmt & vbCRLF
    sBody = sBody & "Staff Name: " & session("EmployeeFname") & " " & session("EmployeeLname") & vbCRLF

    'Read all form fields to see if match on ProductID
    dSubTotAmt = 0
    dDiscAmt = 0

    'Read Products from vw_ShopCartProduct
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM vw_ShopCartProduct WHERE USID = '" & request.form("USID") & "' AND CustomerID = '" & request.form("cu") & "'"
    'response.write sql
    rs.Open sql, strConnect
    do while NOT(rs.EOF)
        body = body & "rs(""SKU"")=" & rs("SKU") & vbCRLF
    
        'Write out InvoiceItems record
    	Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open "tbl_InvoiceItems", strConnect,1,2
        rs1.AddNew
        rs1("CustomerID") 	= request.form("cu")
        rs1("AffiliateID") 	= intAffiliateID
        rs1("CompanyID") 	= uCompanyID
        rs1("ProductID") 	= rs("ProductID")
        rs1("Price") 	    = rs("Price")
        rs1("DiscAmt") 	    = rs("Discount")
        rs1("PriceExtended")= rs("Qty") * rs("Price")
        rs1("Qty") 		    = rs("Qty")
        rs1("InvDate") 		= now
        rs1("InvNum") 		= 0
        rs1("ItemMessage")	= "Thanks for your purchase!"
        rs1("Weight") 	    = rs("Weight")
        rs1.Update
        rs1.Close
        Set rs1 = Nothing
    
        dSubTotAmt = ( dSubTotAmt + (rs("Qty") * rs("Price")) )
        dDiscAmt = FormatNumber(dDiscAmt + rs("Discount")*rs("Qty"),2)

        'Read Next Product from vw_ShopCartProduct
        rs.MoveNext
    loop
    rs.Close
    set rs = nothing
   
    if LEN(request.form("EmployeeID")) = 0 then
        intEmployeeID = 0
    else
        intEmployeeID = CInt(request.form("EmployeeID"))
    end if

    'Add Invoice Record
	Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open "tbl_Invoice", strConnect,1,2
    rs1.AddNew
    rs1("CompanyID") 	= uCompanyID
    rs1("CustomerID") 	= request.form("cu")
    rs1("InvDate") 	    = now
    rs1("SubTotalAmt") 	= dSubTotAmt
    rs1("TaxRate") 	    = 0
    rs1("TaxAmt") 	    = 0
    rs1("Handling") 	= 0
    rs1("ShippingAmt") 	= dTotShipAmt
    'rs1("InvAmt") 	    = (dSubTotAmt - (dDiscAmt + dAddlDiscAmt + dPromoAmount) )
    rs1("InvAmt") 	    = (dSubTotAmt - (dDiscAmt + dAddlDiscAmt + dPromoAmount + dCreditAmt) + dTotShipAmt)
    rs1("ShipMethod") 	= request.form("ShipMethodID")
    rs1("InvDate") 		= now
    rs1("DueDate") 		= CleanDate(request.form("DueDate"))
    rs1("AffiliateID") 	= intAffiliateID
    rs1("EmployeeID") 	= intEmployeeID
    rs1("PayAmt")   	= 0
    'Check for InvAmt GT Applied Credit Amount
    if (dSubTotAmt - (dDiscAmt + dAddlDiscAmt + dPromoAmount ) + dTotShipAmt) > dCreditAmt then
        'If the InvAmt GT Credit being used, then mark as un-PAID
        rs1("Paid") = FALSE
        rs1("InvoiceStatus") = "O"      'Open
    else
        'If the InvAmt LTE Credit being used, then mark as PAID
        rs1("Paid") = TRUE
        rs1("InvoiceStatus") = "P"      'Paid
    end if

    'See if ZERO priced invoice
    if FormatNumber(CSng(dSubTotAmt - (dDiscAmt + dAddlDiscAmt + dPromoAmount + dCreditAmt) + dTotShipAmt),2) = FormatNumber(CSng(0),2) then
        rs1("Paid") = TRUE
        rs1("InvoiceStatus") = "P"      'Paid
    end if

    rs1("PromotionCodeID") = intPromotionCodeID
	
	'Added by Rick
    intPayTypeID = 4  'Default to PO
    rs1("DiscountAmt")  = dDiscAmt + dAddlDiscAmt + dPromoAmount + dCreditAmt
    rs1("PONum")        = request.form("PO")
    rs1("InvoiceNote")  = request.form("Notes")
	rs1("IP")           = Request.ServerVariables("REMOTE_HOST")
    'See if Credit Pay Type
    if request.form("PayType") = "CC" then
	    rs1("Terms") = "CC"
	    intPayTypeID = 1  'CC
    end if
    'See if Check/Pay Later
    if request.form("PayType") = "CK" then
	    rs1("Terms") = "NET30"
	    intPayTypeID = 3 'Check
    end if
    'See if eCheck/ACH
    if request.form("PayType") = "EC" then
	    rs1("Terms") = "eCheck/ACH"
	    intPayTypeID = 2 'eCheck
    end if

    if dCreditAmt > 0 AND request.form("ApplyCR") = "1" then
	    rs1("Terms") = "Credit"
	    intPayTypeID = 6 'Credit Order
    end if
    rs1("PayTypeID") = intPayTypeID
    if LEN(request.form("MarketingID")) > 0 then
        rs1("MarketingID") = request.form("MarketingID")
    end if
    rs1.Update
    intInvoiceNum = rs1("InvoiceNum")
    uInvoiceID = rs1("InvoiceID")
    rs1.Close
    Set rs1 = Nothing
    
    'Process possible Credit Applied
    if dCreditAmt > 0 AND request.form("ApplyCR") = "1" then
        sBody = ""
        sBody = sBody & "Invoice: " & intInvoiceNum & vbCRLF
        sBody = sBody & "Amount: " & dCreditAmt & vbCRLF
        sBody = sBody & "Employee: " & session("EmployeeFname") & " " & session("EmployeeLname") & vbCRLF
        sBody = sBody & "Invoice Link: http://www.nacva.com/z/" & strLinkFolder & "/InvoiceShow.asp?InvoiceNum=" & intInvoiceNum & vbCRLF
        'SendMailCDOSYS(uCustomerID_f, smtp, sUsername, sPassword, sTo, sCC, sBCC, sFrom, sPriority, sSubject, sBody, sAttachment, bHTML, bLog)
        bSend = SendMailCDOSYS(request.form("cu"), NULL, NULL, NULL, "josephinev1@nacva.com", NULL, NULL, "invoice@nacva.com", "Normal", "Paid By Credit (Staff Invoice Create)-1AddInvoiceShopCart_do.asp", sBody, NULL, FALSE, FALSE)
        
        'Read for CreditVoucherID for this Customer with largest CreditAmount
	    Set rs1 = Server.CreateObject("ADODB.Recordset")
        sql1 = "SELECT CreditVoucherID,CreditNote FROM tbl_CreditVoucher WHERE " & _
                "CustomerID = '" & request.form("cu") & "' ORDER BY CreditAmount DESC"
        rs1.Open sql1, strConnect,1,2
        'Only good to write Applied record and update tbl_CreditVoucher.CreditNote if Voucher Exists
        if (NOT(rs1.EOF)) then
            'Add CreditVoucher_Applied record if Credit was used
	        Set rs2 = Server.CreateObject("ADODB.Recordset")
            rs2.Open "tbl_CreditVoucher_Applied", strConnect,1,2
            rs2.AddNew
            rs2("CreditVoucherID") = rs1("CreditVoucherID")
            rs2("AppliedToInvoice") = intInvoiceNum
            rs2("AppliedAmt") = FormatNumber(dCreditAmt,2)
            rs2("CreditAppliedDte") = strToday
            rs2.Update
            rs2.Close
            set rs2 = Nothing

            'Write or append to tbl_CreditVoucher.CreditNote
            if LEN(rs1("CreditNote")) > 0 then
                rs1("CreditNote") = rs1("CreditNote") & vbCRLF & FormatCurrency(dCreditAmt,2) & " applied to Invoice# " & intInvoiceNum
            else
                rs1("CreditNote") = FormatCurrency(dCreditAmt,2) & " applied to Invoice# " & intInvoiceNum
            end if
            rs1.Update
            
        else
            sBody = ""
            sBody = "Not Good - No Credit tbl_CreditVoucher for this CustomerID.  See sql1 below" & vbCRLF
            sBody = sBody & "sql1: " & sql1 & vbCRLF
            sBody = sBody & "rs1.EOF: " & rs1.EOF & vbCRLF
            'SendMailCDOSYS(uCustomerID_f, smtp, sUsername, sPassword, sTo, sCC, sBCC, sFrom, sPriority, sSubject, sBody, sAttachment, bHTML, bLog)
            bSend = SendMailCDOSYS(request.form("cu"), NULL, NULL, NULL, "rickp@nacva.com", NULL, NULL, "nacva1@nacva.com", "Normal", "1AddInvoiceShopCart_do.asp (Line316)", sBody, NULL, FALSE, FALSE)
        end if
        
        rs1.Close
        set rs1 = Nothing
    end if  'if LEN(request.form("CreditAmt")) > 0 AND request.form("ApplyCR") = "1" then
    
    'Update the InvoiceItems with this InvoiceNum, CustomerID, and set Paid according to credits according to credits being used
	Set rs1 = Server.CreateObject("ADODB.Recordset")
    strSQL = "UPDATE tbl_InvoiceItems SET InvNum = " & intInvoiceNum & " WHERE InvNum = 0 AND CustomerID = '" & request.form("cu") & "' " & _
            "AND CONVERT(datetime, CONVERT(char(10), InvDate, 101)) = '" & MONTH(now) & "/" & DAY(now) & "/" & YEAR(now) & "'"
    SET objComm = Server.CreateObject("ADODB.Command")
    objComm.ActiveConnection = strConnect
    objComm.CommandText = strSQL
    objComm.Execute
    Set objComm = nothing

    'Add Payment Record
	Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open "tbl_Payment", strConnect,1,2
    rs1.AddNew
	rs1("CustomerID") 	= request.form("cu")
	rs1("InvNum") 		= intInvoiceNum
	rs1("CompanyID") 	= uCompanyID
    'See if Credit being applied
    if dCreditAmt > 0 AND request.form("ApplyCR") = "1" then
	    'rs1("TransactRef") = "Credit Applied by " & LEFT(session("EmployeeFname"),1) & LEFT(session("EmployeeLname"),1)
	    rs1("PayNote") = "Credit Applied by " & LEFT(session("EmployeeFname"),1) & LEFT(session("EmployeeLname"),1)
    end if
    'See if ZERO priced invoice
    if FormatNumber(CSng(dSubTotAmt - (dDiscAmt + dAddlDiscAmt + dPromoAmount + dCreditAmt) + dTotShipAmt),2) = FormatNumber(CSng(0),2) then
    	rs1("PayDate") = MONTH(now) & "/" & DAY(now) & "/" & YEAR(now)
    else
    	rs1("PayDate") = NULL
    end if
	'rs1("PayAmount") 	= 0 
	rs1("PayAmount") 	= dCreditAmt 'this is zero if no credits are being applied to this purchase, otherwise - the credits are PayAmount
    rs1("PayTypeID")    = intPayTypeID
	rs1("IP")           = Request.ServerVariables("REMOTE_HOST")
	rs1("MerchantID")   = intMerchantID
    rs1.Update
    intPaymentID        = rs1("PaymentID")
    uPayID              = rs1("PayID")
    rs1.Close
    Set rs1 = Nothing

    'response.write("uPayID: " & uPayID & "<br>")
    'response.end
    
    'Delete records from tbl_Shopcart
    strSQL = "DELETE FROM tbl_Shopcart WHERE USID = '" & request.form("USID") & "'"
    SET objComm = Server.CreateObject("ADODB.Command")
    objComm.ActiveConnection = strConnect
    objComm.CommandText = strSQL
    objComm.Execute
    Set objComm = nothing
    
    'Reset Session Vars
    session("NACVA_MarketingID") = ""
    session("NACVA_CompanyID") = ""
    session("NACVA_CustomerID") = ""
    session("NACVA_AffiliateID") = 6
    session("NACVA_ShipMethodID") = 0
    session("PromoCode") = NULL
    
    'See if Member/Customer has Past Credit Cards
    bPastCCs = FALSE
    Set rsC = Server.CreateObject("ADODB.Recordset")
    sqlC = "SELECT Tx_Type,Tx_Num,Tx_Mo,Tx_Yr FROM tbl_Payment WHERE (Tx_Num IS NOT NULL) AND (CustomerID = '" & request.form("cu") & "') ORDER BY PayDate DESC"
    rsC.Open sqlC, strConnect
    'response.Write("rsC.EOF=" & rsC.EOF & "<br>")
    if (NOT(rsC.EOF)) then
        bPastCCs = TRUE
    end if
    rsC.Close
    Set rsC = Nothing
    
    strURL = NULL

    'See if IN Pay Type (INvoice only - No Payment)
    if request("PayType") = "IN" then
        strURL = "../shopcart/success.asp?i=" & intInvoiceNum & "&cu=" & request.form("cu")
    end if

    'See if Credit Pay Type
    if request("PayType") = "CC" then
        strURL = Application("eBiz_NACVA_SSLURL") & "z/" & strLinkFolder & "/MakePaymentByCC.asp?uInvoiceID=" & uInvoiceID & "&CustomerID=" & request.form("cu")
    end if

    'See if Check/Pay Later
    if request("PayType") = "CK" then
        strURL = Application("eBiz_NACVA_SSLURL") & "sysadmin/1MakePmtx_chk.asp?p=" & uPayID
    end if

    'See if eCheck/ACH
    if request("PayType") = "EC" then
        strURL = Application("eBiz_NACVA_SSLURL") & "z/" & strLinkFolder & "/MakePaymentByECheck.asp?uInvoiceID=" & uInvoiceID & "&CustomerID=" & request.form("cu")
    end if

    if (IsNull(strURL)) then strURL = "../shopcart/success.asp?i=" & intInvoiceNum & "&cu=" & request.form("cu")
end if  'End if LEN(request.Form()) > 0

if (NOT(IsNULL(strURL))) then
    %>
    <form name="myForm" id="myForm" action="<%=strURL%>" method="post">
    <!--<input type="hidden" name="PaymentID" value="<%=intPaymentID%>" />-->
    <input type="hidden" name="p" value="<%=uPayID%>" />
    <input type="hidden" name="MerchantID" value="<%=intMerchantID%>" />
    <input type="hidden" name="cy" value="<%=uCompanyID%>" />
    <input type="hidden" name="cu" value="<%=uCustomerID%>" />
    <!--<input type="hidden" name="i" value="<%=intInvoiceNum%>" />-->
    <input type="hidden" name="i" value="<%=uInvoiceID%>" />
    <input type="hidden" name="af" value="<%=request.form("af")%>" />
    </form>
    <%
end if
'response.end
%>    

<script type="text/javascript" language="javascript">
self.opener.location='CustomerInvoiceList.asp?CustomerID=<%=request.form("cu")%>&refresh=1';
<%if (NOT(IsNULL(strURL))) then%>
    document.getElementById('myForm').submit();
    <%if (bPastCCs) AND request("PayType") = "CC" then%>
        //window.open('<%=Application("eBiz_NACVA_SSLSales")%><%=strLinkFolder%>/CustomerPastCreditCards.asp?cu=<%=request.form("cu")%>','mywindow','location=0,status=1,scrollbars=1,width=300,height=300');
    <%end if%>                
<%else%>
    window.opener='Self'; window.open('','_parent',''); window.close();
<%end if%>
</script>