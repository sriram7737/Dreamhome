<!-- This sequence will start every page. -->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->

<HTML>
<HEAD>
<TITLE>View Property Page</TITLE>
</HEAD>
<BODY>
<H1>DreamHome Estate Agents</H1>

<p>Your selected rental...</p>

<%
	Dim propID
	propID = Request.QueryString("ID")
	Dim SQL
	SQL = "SELECT * FROM PropertyForRent WHERE propertyNo = '" & propID & "';"
	Set RS = Application("Conn").Execute(SQL)
	If RS.RecordCount <> 0 Then
		' Display the property details...
		Response.Write "<b>Property Code:   </b>" & RS("propertyNo") & "<br>"
		Response.Write "<b>Street Address:  </b>" & RS("street") & "<br>"
		Response.Write "<b>City:            </b>" & RS("city") & "<br>"
		Response.Write "<b>Property Type:   </b>" & RS("type") & "<br>"
		Response.Write "<b>Number of Rooms: </b>" & RS("rooms") & "<br>"
		Response.Write "<b>Monthly Rental:  </b>" & RS("rent") & "<br>"
		Response.Write "<b>Staff Cotntact:  </b>" & RS("staffNo") & "<br>"
		Response.Write "<b>Branch Code:     </b>" & RS("branchNo") & "<br>"
		Response.Write "<hr>"
		' Display picture...
		If Not IsNull(RS("picture")) Then
			Response.Write "<img src='" & RS("picture") & "'>"
		End If
		' and plan...
		If Not IsNull(RS("floorPlan")) Then
			Response.Write "<img src='" & RS("floorPlan") & "'>"
		End If
		Response.Write "<hr>"
		Response.Write "<a href='bookViewing.asp?PropID=" & propID & "'>Arrange a Viewing</a><br>"
		Response.Write "<a href='index.asp'>Home</a><br>"
	Else
		Response.Write("No matches found.")
	End If
	RS.Close
%>
<hr>

</BODY>
</HTML>