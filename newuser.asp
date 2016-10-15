<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>gClock Keeper Add User</title>
<link rel=stylesheet type="text/css" href="lib/mysheets.css">
</head>
<body bgcolor=cornflowerblue>
<div align="center">
<table cellpadding="5" cellspacing="0" border="5" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="50"> 
<tr><td>
<h1>gClock Keeper</h1><hr noshade width=300 align=left color=midnightblue size=5>
<span class="regtext">Fill-up the form properly. </span>
<form name="frmLogin" method="post" action="addusr.asp?okna">
<table cellpadding="5" cellspacing="0" border="2" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="50" align=center> 
  <TR>
    <TD width="20%"><b>FirstName:</b></TD>
    <TD width="80%"><INPUT name=fname></TD></TR>
  <TR>
    <TD width="20%"><b>LastName:</b> </TD>
    <TD width="80%"><INPUT name=lname></TD></TR>
  <tr>
	<TD width="20%"><b>Password:</b> </TD>
    <TD width="80%"><INPUT type=password name=usrPWD></TD></tr>  
  <tr>
	<TD width="20%"><b>Confirm Password:</b> </TD>
    <TD width="80%"><INPUT type=password name=pwdConf></TD></tr>
  <tr>
	<TD width="20%"><b>Department:</b> </TD>
    <TD width="80%"><select name=dept  size=1 length=15>
		<option value="IT">IT</option>
		<option value="Accounting">Accounting</option>
		<option value="HR">HR</option>
		<option value="Marketing">Marketing</option>
		<option value="CNC">CNC</option>
		<option value="Production">Production</option>
		<option value="Management">Management</option>
		</select>
			
	</TD></tr>
  <TR>
    <TD align="middle" colspan="2">
		<INPUT name=btnReset type=reset value="Reset"> &nbsp; &nbsp; 
  	    <INPUT name=btnSubmit type=submit value="Submit">
	</TD></TR>
</table>
</form>
      <P></P>
</td></tr></table></div>
<!--#include file="baba.htm"-->
</body>
</html>
