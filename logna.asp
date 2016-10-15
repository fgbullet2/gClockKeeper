<!--#include file="lib/scripts.asp"-->
<%

dim sqlRange, strView
sqlRange = ""
strView = "Now viewing all Time Logs..."
Response.Buffer = true
strQuery = Request.QueryString 
dim strError
if strQuery = "" then
	Response.Redirect ("main.asp")
else
	'fill up the variables
	dim strUID, msgStatus
	strUID = Request.Form("strUID")
	if strQuery="elogger" and strUID = "" then
		Response.Redirect ("main.asp")
	end if

	
	'supply all the variables
	dim strFName, strLName, strDept
	
	'connect to the database to supply these
	sqlQuery = "Select * from UserInfo where  UserID = '" & strUID & "'"
	set dbRset = server.CreateObject ("ADODB.Recordset")
	dbRset.Open sqlQuery, myDSN
	
	strFName = dbRset("FirstName")
	strLName = dbRset("LastName")
	strDept = dbRset("Dept")	
	
		
	call CloseRset()
		
 	if strQuery="logthis" then
 		dim strLate, strTime, strDate, strIpHost, mySessID, strNotes, flgTaposNa
 		set dbConn = server.createobject("Adodb.connection")
		dbConn.open myDSN
				
 		strLate = request.form("strLate")
		strTime = request.form("strTime")
		strDate = request.form("strDate")
		strIpHost = request.form("strIpHost")
		mySessID = Request.Form("mySessID")
		strNotes = Request.Form("txtNotes")
		flgTaposNa = Request.Form("taposNa")
	
'=====================Heyaaa I did this for 36 hours continoursly  ang hirap grabe =====================================
' France Gerson Bala
' Systems Engineer
' Personal Computer Specialist
' May 12, 2002

Function gfnISODateTime_3(pData)
	If Cstr(Hour(pData)) <= 9 Then gfnISODateTime_3 = gfnISODateTime_3 + "0"
	gfnISODateTime_3 = gfnISODateTime_3 + Cstr(Hour(pData)) 
	If Cstr(Minute(pData)) <= 9 Then gfnISODateTime_3 = gfnISODateTime_3 + "0"
	gfnISODateTime_3 = gfnISODateTime_3 + Cstr(Minute(pData)) 
	If Cstr(Second(pData)) <= 9 Then gfnISODateTime_3 = gfnISODateTime_3 + "0"
	gfnISODateTime_3 = gfnISODateTime_3 + Cstr(Second(pData))
End Function

dim strHour, strMinute, strSeconds, strAll

Function gfnISODateTime(pData)
	If Cstr(Hour(pData)) <= 9 Then gfnISODateTime = gfnISODateTime + "0"
	strHour = strHour + Cstr(Hour(pData)) 
	If Cstr(Minute(pData)) <= 9 Then strMinute = strMinute + "0"
	strMinute = strMinute + Cstr(Minute(pData)) 
	If Cstr(Second(pData)) <= 9 Then strSeconds = strSeconds  + "0"
	strSeconds =  strSeconds + Cstr(Second(pData))
End Function
gfnISODateTime(now)

Dim h, m, s , varInitial_Time_2

	h = "09"
	'if strDept = dbRset("Support") & Admin & Sales then m = "30" else 
	m = "30"
	s = "00"
	varInitial_Time_2 = h + m + s

Dim varLogin_Time_2
	varLogin_Time_2 = Time

dim h1, m2, s3
h1 = (strHour - h)	
	if strMinute >= 30 then h1 = h1 + 1 

m2 = (strMinute - m)
	if m2 <= 29 then m2 = m2 + 60
	if m2 > 59 then m2 = m2 - 60 
	if m2 <= 9 then m2 = "0" & m2 'else m2 = m2
	
s3 = strSeconds

strAll = h1 & ":" & m2 & ":" &  s3 
if gfnISODateTime_3(now) >=  083000 then strAll = strAll else strAll = "00:00:00"

strLate = strAll
strTime = varLogin_Time_2
'strTimeOut = varLogin_Time_2

'-----------------------------------------------


		sqlQuery = "Select TimeIn,TimeOut From TimeInOut where SessionID = '" & mySessID & "'"
		set dbRset = dbConn.execute (sqlQuery)
		if request.form("radTimeIn") = "on" then
			if dbRset.EOF then
				flgTaposNa = "hindi"
				sqlQuery = "Insert into TimeInOut (DateInOut, TimeIn, Late, UserID, ipAddr, Notes,SessionID) Values ('" & strDate & "', '" & strTime & "', '" & strLate & "', '" & strUID & "','" & strIpHost & "', '" & strNotes &"','" & mySessID &"')"
				dbConn.execute sqlQuery
			else
				strError =	"<br>Logging-in twice is denied!"
			end if
		elseif request.form("radTimeOut") = "on" then
			if dbRset.EOF then
				sqlQuery = "Insert into TimeInOut (DateInOut, TimeOut, Late, UserID, ipAddr, Notes, SessionID) Values ('" & strDate & "', '" & strTime & "', '" & strLate & "', '" & strUID & "','" & strIpHost & "','" & strNotes & "','" & mySessID &"')"              
			else
				if flgTaposNa = "oo" then
					strError="<br>Logging-out twice is denied!"
				else
					sqlQuery = "Update TimeInOut Set TimeOut = '" & strTime & "' Where SessionID = '" & mySessID & "'"
					flgTaposNa="oo"
				end if
			end if
			dbConn.Execute sqlQuery
 		end if 				
		call CloseRset()
		call CloseConn()
 	end if

	if strQuery = "viewthis" then
		if request.form("radSel") = "on" then
			sqlRange = "AND DateInOut Between #" & request.form("frmSel") & "# AND #" & request.form("toSel") & "#"
			strView = "Now viewing records from " & request.form("frmSel")& " to " & request.form("toSel") & "..."
		end if
	end if
end if


%>





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>gClock Keeper for
<%=strFName & " " & strLName & " (" & strDept & ")" %>
</title>
<link rel=stylesheet type="text/css" href="lib/mysheets.css">
</head>
<body bgcolor=CornflowerBlue>
<div align="center">
<table cellpadding="5" cellspacing="0" border="5" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="70%"> 
<tr valign="middle"><td valign="middle" colspan=2><span class='regtext'><i>gClock Keeper for</i></span><br>
<h1>
<%=strFName & " " & strLName & " (" & strDept & ")"%>
</h1>
</td></tr>
<tr><td valign="middle" width="40%">
Today is:<b> <%=Date()%></b><br>The time is: <b><%=varLogin_Time_2%></b><br>
Your IP address is:<b><%=request.servervariables("Remote_Host")%></b><br>

<% 
if strDept = "HR" or strDept = "IT" or strDept = "Accounting" or strDept = "Management" then
	Response.Write ("<span class='regtext'><br>View logs of all personnel: <form name='frmUID'  method='post' action='export.asp?elogger'><input name='strUID' value='" & strUID & "' type='hidden'><input type='submit' name='btnSubmit' value='Click Here!'></form></span>")
end if
%>
<br>
<table cellpadding="5" cellspacing="0" border="3" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="100%"> 
<tr>

<!-- Start Time Logs Filter-->

<td width="60%"><b>Please select:</b></td>
<td width="40%"><b>Log your Time</b></td></tr>
<tr><td>
<form name="frmTimeIO" method="POST" action="logna.asp?viewthis">
<input name="strUID" value="<%=strUID%>" type="hidden">
<%

'open the database
sqlQuery = "Select DateInOut from TimeInOut where  UserID = '" & strUID & "' Order by DateInOut"
set dbRset = server.createobject("ADODB.Recordset")
dbRset.open sqlQuery, myDSN

if dbRset.EOF then
	Response.write ("No entry to display")
else
	Response.write ("<input type='radio' name='radAll' checked onClick='document.frmTimeIO.radSel.checked = false'>&nbsp;&nbsp;View all Time Logs<br>" & vbcrlf)
	Response.write ("<input type='radio' name='radSel' onClick='document.frmTimeIO.radAll.checked = false'>&nbsp;&nbsp;View selected Time Logs<br>" & vbcrlf)
	Response.write ("View from <select name='frmSel'>" & vbcrlf)
	do until dbRset.EOF 

		response.write("<option value='" & dbRset("DateInOut") & "'>" & dbRset("DateInOut") & " </option>" & vbcrlf)
		dbRset.movenext
	loop
	dbRset.moveFirst
	Response.write ("</select> to <select name='toSel'>" & vbcrlf)
	do until dbRset.EOF 
		response.write("<option value='" & dbRset("DateInOut") & "'>" & dbRset("DateInOut") & " </option>" & vbcrlf)				
		dbRset.movenext
	loop
	Response.write ("</select>")
end if

dbRset.close
set dbRset = nothing
%>
<p>
<input type="submit" name="btnSubmit" value="View It!">
</form>
<p>
<b><%=strView %></b>
</td>
<!-- End Time Logs Filter-->
<td width="50%">
<form name="frmTime" method="POST" action="logna.asp?logthis">
<br><div align="center">
<input type="radio" name="radTimeIn" checked onClick="document.frmTime.radTimeOut.checked = false">Time In 
&nbsp;&nbsp;
<input type="radio" name="radTimeOut" onClick="document.frmTime.radTimeIn.checked = false">Time Out<br>
Notes:<br>
<textarea name="txtNotes" rows=3 cols=25 wrap=virtual></textarea>
<input type="submit" name="btnSubmit" value="Submit it!">
<input type="hidden" name="strDate" value="<%=Date()%>">
<input type="hidden" name="mySessID" value="<%=strUID & Cstr(Date())%>">
<input type="hidden" name="taposNa" value="<%=flgTaposNa%>">
<input type="hidden" name="strTime" value="<%=Time()%>">
<!--<input type="hidden" name="strTime" value="<%=varLogin_Time_2%>">/-->
<input type="hidden" name="strIPhost" value="<%=request.servervariables("Remote_Host")%>">
<input name="strUID" value="<%=strUID%>" type="hidden">
</div>
<% 
if not strError = "" then
	Response.Write ("<b>&nbsp;&nbsp;&nbsp;&nbsp;" & strError&"</b>")
end if
%>
</form>
</td></tr>

</table>
</td></tr>
<tr><td align='center' valign='middle'>
<table cellpadding="5" cellspacing="0" border="2" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="100%"> 
<tr>
<% 
	'open the database
	sqlQuery = "Select DateInOut, TimeIn, TimeOut, Late, Notes from TimeInOut where  UserID = '" & strUID & "'" & sqlRange & " order by DateInOut;"
	set dbRset = server.CreateObject ("ADODB.Recordset")
	dbRset.open sqlQuery, myDSN
	if dbRset.EOF then
		Response.Write "<td colspan=4>No records yet."
		Response.Write "</td></tr></table>"
	else
		Response.Write "<td><b>Date</b></td>"
		Response.Write "<td><b>Time In</b></td>"
		Response.Write "<td><b>Time Out</b></td>"
		Response.Write "<td><b>Late</b></td>"
		Response.Write "<td><b>Notes</b></td></tr><tr><td>"
		response.write dbRset.getstring (,,"</td><td>","</td></tr><tr><td>"," - - -")
		Response.Write "</td></tr></table>"
	end if
	
	
	call CloseRset()
%>
</td></tr></table>
</td></tr>
</table>
<!--#include file="baba.htm"-->
</body>
</html>