<!--#include file="lib/scripts.asp"-->
<%
dim strUID, strPWD
strUID =""
strPWD = ""


strQuery = Request.QueryString 

if strQuery = "login" then
	'get the values from the form
	
	strUID = Request.Form("select1")
	strPWD = Request.Form("usrPasswd")
	
	sqlQuery = "Select * from UserInfo Where UserID = '" & strUID & "' and Password = '" & strPWD & "'"
	'open the database
	set dbRset = Server.CreateObject("ADODB.Recordset")
	dbRset.open sqlQuery, myDSN
	
	'check contents..
	if dbRset.EOF then
		LogFlag = 2
	else
		LogFlag = 3
	end if
	call CloseRset()    
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
