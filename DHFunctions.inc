<!--Create and open a connection to the database-->
<%
'**************************************************
'*** GetConnectionString function
'**************************************************
'*** Written 9/5/2004
'*** Author: A.McMonnies
'*** Revision 2.1
'*** Purpose - Function returns a string specifying
'*** a full connection string for opening an Access
'*** (JET 4.0) database in the same folder as the
'*** page the function is run from.
'==================================================
'*** Input parameter - db As String - name of MDB 
'*** database file.
'*** Result - Connection-String As String
'***************************************************
Function GetConnectionString(db)
	GetConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & Server.MapPath(db) & _
	           ";Persist Security Info=False"
End Function


'**************************************************
'*** GetClientName function
'**************************************************
'*** Written 19/5/2004
'*** Author: A.McMonnies
'*** Revision 1.0
'*** Purpose - Function returns a string containing
'*** client's name 
'==================================================
'*** Assumes a connection is already made
'==================================================
'*** Input parameter - id - a string of clientNo
'*** Returns Client's name.
'==================================================
Function GetClientName(id)
	Dim RS 
	Set RS = Application("Conn").Execute("select fName + ' ' + lName as Name from Client where ID=" & id)
	If Not RS.EOF Then
		GetClientName = RS("Name")
	Else
		GetClientName = "Client no " & id & " does not exist"
	End If
	RS.Close
	Set RS = Nothing
End Function

'**************************************************
'*** GetStaffName function
'**************************************************
'*** Written 19/5/2004
'*** Author: A.McMonnies
'*** Revision 1.0
'*** Purpose - Function returns a string containing
'*** staff member's name 
'==================================================
'*** Assumes a connection is already made
'==================================================
'*** Input parameter - staffNo - a string of staffNo
'*** Returns Staff member's name.
'==================================================
Function GetStaffName(staffNo)
	Dim RS 
	Set RS = Application("Conn").Execute("select fName + ' ' + lName as Name from Staff where staffNo='" & staffNo & "';")
	If Not RS.EOF Then
		GetStaffName = RS("Name")
	Else
		GetStaffName = "Staff no " & id & " does not exist"
	End If
	RS.Close
	Set RS = Nothing
End Function


'**************************************************
'*** RSToHTMLTable function
'**************************************************
'*** Written 9/5/2004
'*** Author: A.McMonnies
'*** Revision 1.0
'*** Purpose - Function returns a complete HTML
'*** table as a text string.  Table content is all
'*** of the data from a ADO Recordset.  First row
'*** is column (field) names, followed by one row
'*** per record for the full table.
'*** If table has an 'ID' column, a 'link' param
'*** -eter can be supplied to specify the URL of
'*** a linked ASP which will be sent the ID as a
'*** GET parameter. 
'*** NOTE - use SQL AS syntax to set column names.
'***        e.g.  SELECT CustName As Name FROM....
'==================================================
'*** Input parameters
'*** 1. RS As Recordset - an ADO recordset which is
'***    currently open.
'*** 2. borderwidth As Integer (normally a small
'***    integer value - 1 or 2.
'*** 3. URL of an ASP for accessing individual
'***    records from the table (i.e. a detail view
'***    page.  Only to be used if the table in RS
'***    has an ID field.
'*** Result - HTML Text As String
'***************************************************
Function RSToHTMLTable(RS, borderwidth, link)
Dim TBL, i
    ' Response.Write "<h2>" & link & "</h2>"
	' Generate a HTML Table from the recordset data...
	TBL = "<TABLE border=" & borderwidth & ">"
	' Start with a header row...
	TBL = TBL & "<TR>"
	F = ""
	For Each fld In RS.Fields
		F = F & "<TH>" & fld.Name & "</TH>"
	Next
	If link<>"" Then
		F = F & "<TH>Select</TH>"
	End If
	TBL = TBL & F & "</TR>"
	' Now the data...
	If Not RS.EOF Then
        RS.MoveFirst
	    Do
		    TBL = TBL & "<TR>"
    		F = ""
	    	For Each fld In RS.Fields
    	        F = F & "<TD>" & fld.Value & "</TD>"
    	    Next
	    	If link<>"" Then
		        F = F & "<TD><a href='" & link & "?ID=" & RS.Fields("ID").Value & "'>View Detail</a></TD>"
    	    End If
	        TBL = TBL & F & "</TR>"
	        RS.MoveNext
	    Loop Until RS.EOF
    End If
	TBL = TBL & "</TABLE>"
	RSToHTMLTable = TBL
End Function

'**************************************************
'*** GenerateDateOptions function
'**************************************************
'*** Written 19/5/2004
'*** Author: A.McMonnies
'*** Revision 1.2
'*** Purpose - Function returns a complete HTML
'*** tag set for an option button with a range of
'*** dates pre-loaded.  This HTML should be embedded
'*** in a form for user-input.
'==================================================
'*** Input parameters
'*** 1. StartDate As String - date at start end of range.
'*** 2. EndDate As String - date at end of range.
'*** Result - HTML Text As String
'***************************************************
Function GenerateDateOptions(StartDate, EndDate)
	Dim HTML
    Dim Q
    Q = Chr(34)
	HTML = "<select name='optDate'>"
	Dim D
	D = StartDate
	Do While D <= EndDate
		HTML = HTML & "<option value=" & Q & D & Q & ">" & D & "</option>"
		D = DateAdd("d", 1, D)
	Loop
	HTML = HTML & "</select>"
	GenerateDateOptions = HTML
End Function

'**************************************************
'*** GenerateHourOptions function
'**************************************************
'*** Written 13/6/2004
'*** Author: A.McMonnies
'*** Revision 1.0
'*** Purpose - Function returns a complete HTML
'*** tag set for an option button with a range of
'*** hour-intervals pre-loaded.  This HTML should
'*** be embedded in a form for user-input.
'==================================================
'*** Input parameters
'*** 1. startHour As String - hour at start end of range.
'*** 2. endHour As String - end at end of range.
'*** Result - HTML Text As String
'***************************************************
Function GenerateHourOptions(startHour, endHour)
    Dim h, options
    Dim Q
    Q = Chr(34)
    options = "<select name='optHour'>"
    h = startHour
    If endHour < startHour Then
        endHour = endHour + 12
    End If
    Do While h < endHour
        options = options & "<option value='" & h & "'>" & h & ":00 to " & h+1 & ":00</option>" 
        h = h+1
    Loop 
    options = options & "</select>"
    GenerateHourOptions = options
End Function  
%>
