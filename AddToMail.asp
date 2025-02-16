<!-- This sequence will start every page. -->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->

<HTML>
<HEAD>
<TITLE>Add to Mailing List</TITLE>
</HEAD>
<BODY>
<H1>DreamHome Estate Agents</H1>
<h3>Mailing List</h3>
<%
	' Need to insert a record to Clients table of the database. 
	' We need to add a clientNo, so best approach is to open a writable recordset...
	Dim RS, conn
    Set conn = Application("Conn")
	Const adOpenDynamic = 2
	Const adLockOptimistic = 3
	Const adCmdTable = 2
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "Client", Conn, adOpenDynamic, adLockOptimistic, adCmdTable
	' The client table is now open.  Add a new record...
	RS.AddNew
		RS("fName") = Request.Form("txtFName")
		RS("lName") = Request.Form("txtLName")
		RS("TelNo") = Request.Form("txtTel")
		RS("Street") = Request.Form("txtStreet")
		RS("City") = Request.Form("txtCity")
		RS("PostCode") = Request.Form("txtPostCode")
		RS("email") = Request.Form("txtEmail")
		RS("JoinedOn") = Date
		RS("Region") = Request.Form("cboRegion")		
		RS("preType") = Request.Form("optPreType")
		RS("maxRent") = Request.Form("txtMaxRent")
	RS.Update
	' Now set the clientNo up in the user's browser...
	clientNo = RS("ID")
	Response.Cookies("ClientID") = clientNo
	Response.Cookies("ClientID").Expires = DateAdd("m", 3, Date)
    Session("ClientID") = clientNo
	If Err.Number > 0 Then
		Response.Write "<p>An error has occurred.  Please try to join our list later.</p>"
		'Response.Write "<p>ERROR: " & Err.Number & ": " & Err.Description & "</p>"
	Else
		Response.Write "<p>Details stored in your mailing list entry are: </p>"
		Response.Write "<p>Customer ID :           </b>" & RS("ID")
		Response.Write "<b>Name :                  </b>" & RS("fName") & " " & RS("lName") & "<br>"
		If RS("telNo") <> "" Then
			Response.Write "<b>Telephone no.:          </b>" & RS("telNo") & "<br>"
		End If
		Response.Write "<b>Address :               </b>" & RS("Street") & "<br>"
		Response.Write "<b></b>                        " & RS("City") & "<br>"
		Response.Write "<b></b>                        " & RS("PostCode") & "<br>"
		Response.Write "<b></b>                        " & RS("City") & "<br>"
		If RS("email") <> "" Then
			Response.Write "<b>Email address :        </b> " & RS("email") & "<br>"
		End If
		If RS("Region") <> "(N/A)" Then
			Response.Write "<b>Preferred area :       </b> " & RS("Region") & "<br>"
		End If
		If RS("preType") <> "" Then
			Response.Write "<b>Preferred rental type : </b>" & RS("preType") & "<br>"
		End If
		If RS("maxRent") > 0 Then
			Response.Write "<b>Maximum monthly rent :  </b>" & RS("maxRent") & "<br><br>"
		End If
		Response.Write "This information will be stored securely."
	End If
	RS.Close
	Set RS = Nothing
%>
<br><br><A HREF="index.asp">Home</a>
</BODY>
</HTML>