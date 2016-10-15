<%
Option Explicit
dim strUserID, strFname, strLName, strCoy
strUserID = Request.QueryString ("UserID")
strFname = Request.QueryString ("FName")
strLName = Request.QueryString ("LName")
strCoy = Request.QueryString ("Coy")

%>


<html>
<head>
<title><% = strFname & " " & strLName & " (" & strCoy & ")" %></title>
<link rel=stylesheet type="text/css" href="lib/mysheets.css">
</head>
<body bgcolor=CornflowerBlue>
<div align="center">
<form action="logna?addTime" method="post" name="logto"> 
<table cellpadding="5" cellspacing="0" border="5" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="35%"> 
<tr valign="middle"><td><h1>
<% = strFname & " " & strLName & " (" & strCoy & ")" %>
</h1></td></tr>
<tr><td align="center"><h3>
<input type="checkbox" name="radLogIn" onClick="document.logto.radLogOut.checked = false; return true;" size=50> Time In
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="radLogOut" onClick="document.logto.radLogIn.checked = false; return true;" size=50> Time Out</h3>
<tr><td align="center">
<span class="regtext">
<input type="hidden" value="<%=Time()%>" name="TimeNow">
<input type="hidden" value="<%=Date()%>" name="DateNow">
Time: <b> <% =Time() %></b><br>
Date: <b> <% =Date() %></b><br>
</span>
</td></tr>
<tr><td>
<input type="submit" value="Submit!">
</table>
</form>
</html>