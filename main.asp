<!--#include file="lib/scripts.asp"-->
<% 
'get serverVariables first!!!
strRemHost = Request.ServerVariables("REMOTE_HOST")
strQuery = Request.QueryString 

dim LogFlag
LogFlag = 0		'1=invalid pwd, 2=okey!

if strQuery = "login" then
	'now get the form
	dim strUID, strPWD
	strUID = Request.Form("select1")
	strPWD = Request.Form("usrPasswd")
	
	if strPWD = "" then
		LogFlag = 0
	end if
	
	'authenticate it!!
	sqlQuery = "Select * from UserInfo Where UserID = '" & strUID & "' and Password = '" & strPWD & "'"
	set dbRset = server.CreateObject ("ADODB.Recordset")
    dbRset.Open sqlQuery, myDSN
    
    if dbRset.EOF then
		LogFlag = 1		'invalid password.. ang tanga tanga moooooOOOOOOOOoooooOOoOoOooO.  :p
	else 
		LogFlag = 2		'galing tsong!
	end if
	
	call CloseRset()	
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Welcome to gClock Keeper</title>
<link rel=stylesheet type="text/css" href="lib/mysheets.css">
<script language="Javascript">
<!--
function sendIt(){
	document.frmUID.submit();
	}
	
//-->	
</script>

</head>
<body bgcolor=CornflowerBlue>
<div align="center">
<table cellpadding="5" cellspacing="0" border="5" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="35%"> 
<tr><td>
<h1>gClock Keeper</h1><hr noshade width=300 align=left color=MidnightBlue size=5>
<span class="regtext">Welcome to gClock Keeper</span><br><br>
<span class="regtext">Use this as your time card.<br>
Your IP address is:<%=strRemHost %></span><p>
<form name="frmUID"  method="post" action="logna.asp?elogger">
<input name="strUID" value="<%=strUID%>" type="hidden">
</form> 
<form name="frmLogin" method="post" action="main.asp?login" onload="initForm();">
<table cellpadding="5" cellspacing="0" border="2" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="50" align=center> 
  <TR>
    <TD width="20%"><b>Select Name</b></TD>
    <TD width="80%">
    <%
    sqlQuery = "Select * from UserInfo Order by LastName"
    'connect to the database...
    set dbRset = server.CreateObject ("ADODB.Recordset")
	dbRset.Open sqlQuery, myDSN
    
    'check if empty
    if dbRset.EOF then
		Response.write("No records yet..")
	else
		response.write("<select id=select1 name=select1>" & vbCrLf)
		
		do until dbRset.EOF 
			if strUID = dbRset("UserID") then
				response.write("<option selected value='" & dbRset("UserID") & "'>" & dbRset("LastName") & "," & dbRset("FirstName") & " (" & dbRset("Dept") & ") </option>" & vbcrlf)
			else
				response.write("<option value='" & dbRset("UserID") & "'>" & dbRset("LastName") & "," & dbRset("FirstName") & " (" & dbRset("Dept") & ") </option>" & vbcrlf)
			end if
			dbRset.MoveNext 
		loop
		response.write("</select>" & vbCrLf)
	end if
	call CloseRset()    
 %>
    </TD></TR>
  <TR>
    <TD width="20%"><b>Password:</b> </TD>
    <TD width="80%"><INPUT type=password name=usrPasswd size=20></TD></TR>
  <TR>
    <TD align="center" colspan="2">
		<INPUT name=btnReset type=reset value=Reset> &nbsp; &nbsp; 
  	    <INPUT name=btnSubmit type=submit value=Submit>
	</TD></TR>
  <TR>
    <TD align="center" colspan="2">
		Forgot your password? Click <a href="lostpass.htm">here.</a><br>
		New User? Click <a href="newuser.asp">here.</a>
	</TD></TR>
</table>
<span class="regtext" >
<script language="Javascript">
<!--
var LFlag = <%=LogFlag%>;
if (LFlag == 1) document.write ("Invalid password. Please try again.")
if (LFlag == 2) sendIt();
//-->
</script>
</span>
</form>
</td></tr></table>
<!--#include file="baba.htm"-->
</body>
</html>
