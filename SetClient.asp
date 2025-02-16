<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<html>
<head><title>Home.ASP</title></head>
<body>
<%
    If Request.QueryString("ID") = "none" Then
        Response.Cookies("ClientID") = ""
        Response.Cookies("ClientID").Expires = Date
        Session.Abandon
        Response.Write "Client reset"
    ElseIf Request.QueryString("ID") = "who" Then
        Response.Write "'ClientID' Cookie now set to : " & Request.Cookies("ClientID")
    ElseIf Request.QueryString("ID") = "" Or Request.QueryString("TimeOut")= "" Then
        Response.Write "Usage: SetClient.asp?ID=[id_value_to_set]&TimeOut=[number_of_days]"
    Else
        Response.Cookies("ClientID") = Request.QueryString("ID")
        Response.Cookies("ClientID").Expires = DateAdd("d", Request.QueryString("TimeOut"), Date)
        Response.Write "'ClientID' Cookie now set to : " & Request.QueryString("ID") & "<br>"
        Response.Write "Cookie will time out on " & DateAdd("d", Request.QueryString("TimeOut"), Date)
    End If
%>
</body>
