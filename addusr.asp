<%
dim strfname
dim strlname
dim struserID
dim strpwd
dim strdept
dim strSQL
'get all of them

strfname = Request.Form("fname")
strlname = Request.Form ("lname")
struserID = Request.Form ("userID")
strpwd = Request.Form ("usrPWD")
strdept = Request.Form ("dept")

%>
<HTML>
<HEAD>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<TITLE>gClock Keeper</TITLE>
<link rel=stylesheet type="text/css" href="lib/mysheets.css">
</HEAD>
<body bgcolor=cornflowerblue>
<table cellpadding="5" cellspacing="0" border="5" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="50%"> 
<%
dim flg
flg = frm_valid8()
if flg = "" then
	if check_it() = "" then
		call write_it()
		Response.Write ("<tr><td><h1> User added </h1></td></tr>")
		Response.write ("<tr><td><span class='regtext'>" & "User " & strfname & " " & strlname & " (" & strdept & ") has been successfully added!</span></td></tr>")
		Response.Write ("</td></tr></table><br><br>")
		Response.Write ("<span class='regtext'><a href='main.asp'>Click here to LogIn now.</a></span>")
	else 
		Response.Write ("<tr><td><h1> User Exists ! </h1></td></tr>")
		Response.write ("<tr><td><span class='regtext'>" & "User <u>" & check_it() & "</u> already exists!</span></td></tr>")
		Response.Write ("</td></tr></table><br><br>")
		Response.Write ("<span class='regtext'><a href='javascript:window.history.go(-1)'>Click here to go back to the form.</a></span>")
	end if
else
	Response.Write ("<tr><td><h1> Form Processing Error </h1></td></tr>")
	Response.Write ("<tr><td><span class='regtext'>")
	if instr(1,flg,"f",1) then Response.Write ("Please Provide FirstName<br>")
	if instr(1,flg,"l",1) then Response.Write ("Please Provide LastName<br>")
	if instr(1,flg,"p",1) then Response.Write ("Please Provide Password<br>")
	if instr(1,flg,"c",1) then Response.Write ("Please confirm password<br>")
	Response.Write ("</td></tr>")
	Response.Write ("</td></tr></table><br><br>")
	Response.Write ("<span class='regtext'><a href='javascript:window.history.go(-1)'>Click here to go back to the form.</a></span>")
end if

function frm_valid8()
	dim flagfrm
	flagfrm = ""
	if strfname = "" then
		flagfrm = flagfrm & "f"
	end if
	if strlname = "" then
		flagfrm = flagfrm & "l"
	end if
	if strpwd = "" then
		flagfrm = flagfrm & "p"
	else 
		if strpwd <> Request.Form("pwdConf") then
			flagfrm = flagfrm & "c"
		end if
	end if
	frm_valid8 = flagfrm
end function
	
function check_it()
	Set rsDB = server.CreateObject("ADODB.Recordset")
	rsDB.Open "UserInfo","DSN=eloggerdb;uid=elogger;pwd=password"
	
	
	do while not rsDB.EOF
		if (LCase(rsDB("LastName")) = LCase(strlname) AND rsDB("Dept")= strdept) AND LCase(rsDB("FirstName")) = LCase(strfname) then
			check_it = strfname & " " & strlname & " (" & strdept & ")"
			exit do
		else
			rsDB.MoveNext
		end if
	loop
	rsDB.close
	set rsDB = nothing
end function
	
sub write_it()
	dim connDB
	Set connDB = server.CreateObject("ADODB.Connection")
	connDB.Open "eloggerdb","elogger","password"
	strSQL = "INSERT INTO UserInfo (UserID, Password, FirstName, LastName, Dept) VALUES ('" & nUID() & "', '" & strpwd & "', '" & strfname & "', '" & strlname & "', '" & strdept & "')"
	'Response.Write (strSQL)
	'connDB.Mode = 3		'3 = adModeReadWrite //6march2015
	connDB.Execute (strSQL)
	connDB.Close
	set connDB = nothing
end sub

function nUID()
	dim ctr, coy
	dim rsInfo
	set rsInfo = server.CreateObject ("ADODB.Recordset")
	rsInfo.Open "UserInfo","DSN=eloggerdb;uid=elogger;pwd=password"
	ctr = 0
	do while not rsInfo.EOF
		ctr = ctr +1
		rsInfo.MoveNext	
	loop
	'YearName = year
	nUID = lcase(trim(strco)) & Year(date) & Cstr(ctr +1) 
	rsInfo.Close
	set rsInfo = nothing
end function 
%>
</BODY>
</HTML>
