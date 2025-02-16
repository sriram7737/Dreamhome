<!-- This sequence will start every page. -->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->

<HTML>

<HEAD>
<TITLE>Rentals View Page</TITLE>
</HEAD>
<BODY>
<H1>DreamHome Estate Agents</H1>

<p>Rentals that match your criteria...</p>

<%
	Set Conn = Application("Conn")
	' Formulate a SQL Select query, uising # to mark values...
	SQL = "SELECT propertyNo As ID, street, city, postcode, fname+' '+lname " & _
            "As contact FROM PropertyForRent INNER JOIN staff ON PropertyForRent.staffNo=staff.staffNo "
	If Request.Form("RentalArea") = "ANY" Then
		SQL = SQL & "WHERE type = '#2' AND rooms >= #3 & rent <= #4"	
	Else
		SQL = SQL & "WHERE PropertyForRent.branchNo = '#1' AND type = '#2' AND rooms >= #3 AND rent <= #4"
	End If
	' Now replace the values...
	SQL = Replace(SQL, "#1", Request.Form("RentalArea"))
	SQL = Replace(SQL, "#2", Request.Form("RentalType"))
	SQL = Replace(SQL, "#3", Request.Form("RentalSize"))
	SQL = Replace(SQL, "#4", Request.Form("RentalMaxPrice"))
	'SQL Statement debug...
	'Response.Write SQL
	Set RS = Conn.Execute(SQL)
	If Not RS.EOF Then
		Response.Write RSToHTMLTable(RS, 1, "viewProp.asp")
	Else
		Response.Write("No matches found for your criteria.")
	End If
	RS.Close
%>
<P>
</BODY>
</HTML>