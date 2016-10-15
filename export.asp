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
	'logthis start	
 	'if strQuery="logthis" then
 	'	dim strLate, strTime, strDate, strIpHost, mySessID, strNotes, flgTaposNa
 	'	set dbConn = server.createobject("Adodb.connection")
	'	dbConn.open myDSN
				
 	'	strLate = request.form("strLate")
	'	strTime = request.form("strTime")
	'	strDate = request.form("strDate")
	'	strIpHost = request.form("strIpHost")
	'	mySessID = Request.Form("mySessID")
	'	strNotes = Request.Form("txtNotes")
	'	flgTaposNa = Request.Form("taposNa")
	


	'	sqlQuery = "Select TimeIn,TimeOut From TimeInOut where SessionID = '" & mySessID & "'"
	'	set dbRset = dbConn.execute (sqlQuery)
	'	if request.form("radTimeIn") = "on" then
	'		if dbRset.EOF then
	'			flgTaposNa = "hindi"
	'			sqlQuery = "Insert into TimeInOut (DateInOut, TimeIn, Late, UserID, ipAddr, Notes,SessionID) Values ('" & strDate & "', '" & strTime & "', '" & strLate & "', '" & strUID & "','" & strIpHost & "', '" & strNotes &"','" & mySessID &"')"
	'			dbConn.execute sqlQuery
	'		else
	'			strError =	"<br>Logging-in twice is denied!"
	'		end if
	'	elseif request.form("radTimeOut") = "on" then
	'		if dbRset.EOF then
	'			sqlQuery = "Insert into TimeInOut (DateInOut, TimeOut, Late, UserID, ipAddr, Notes, SessionID) Values ('" & strDate & "', '" & strTime & "', '" & strLate & "', '" & strUID & "','" & strIpHost & "','" & strNotes & "','" & mySessID &"')"              
	'		else
	'			if flgTaposNa = "oo" then
	'				strError="<br>Logging-out twice is denied!"
	'			else
	'				sqlQuery = "Update TimeInOut Set TimeOut = '" & strTime & "' Where SessionID = '" & mySessID & "'"
	'				flgTaposNa="oo"
	'			end if
	'		end if
	'		dbConn.Execute sqlQuery
 	'	end if 				
	'	call CloseRset()
	'	call CloseConn()
 	'end if
	'logthis end
	if strQuery = "viewthis" then
		if request.form("radSel") = "on" then
			'sqlRange = "AND DateInOut Between #" & request.form("frmSel") & "# AND #" & request.form("toSel") & "#"
			sqlRange = "Between #" & request.form("frmSel") & "# AND #" & request.form("toSel") & "#"
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
<link rel=stylesheet type="text/css" href="lib/mysheets3.css">
</head>
<body bgcolor=CornflowerBlue>
<div align="center">
<table cellpadding="5" cellspacing="0" border="5" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="70%"> 
	<tr valign="middle">
		<td valign="middle" colspan=2><span class='regtext'><i>gClock Keeper for</i></span><br>
			<h1>
				<%=strFName & " " & strLName & " (" & strDept & ")"%>
			</h1>
		</td>
	</tr>
	<tr>
		<td valign="middle" width="40%">
			Today is:<b> <%=Date()%></b><br>
			Your IP address is:<b><%=request.servervariables("Remote_Host")%></b><br>

			<% 
				'if strDept = "HR" or strDept = "IT" or strDept = "Accounting" or strDept = "Management" then
				if strDept = "HR" or strDept = "IT" then
					Response.Write ("<span class='regtext'><br>Now viewing logs of all employees!</span>")
				end if
			%>
			<br>
			<table cellpadding="5" cellspacing="0" border="3" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="100%"> 
				<tr>
					<!-- Start Time Logs Filter-->
					<td width="60%"><b>Please select the date range that you want to export:</b></td>
					<td width="40%"><b>Export logs:</b></td>
				</tr>
				<tr>
					<td>
						<form name="frmTimeIO" method="POST" action="export.asp?viewthis">
						<input name="strUID" value="<%=strUID%>" type="hidden">
						<%
						'open the database
						'sqlQuery = "Select DateInOut from TimeInOut where  UserID = '" & strUID & "' Order by DateInOut"
						'sqlQuery = "Select DateInOut from TimeInOut Order by DateInOut"
						sqlQuery = "Select DISTINCT DateInOut	from TimeInOut Order by DateInOut"
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
					
					<!-- For Export Logs-->
					<td width="50%">
						Export to Excel format:&nbsp;&nbsp;
						<form name='frmUID'  method='post' action='export_xls.asp?elogger'>
							<input name="strUID" value="<%=strUID%>" type="hidden">
							<input name="sqlRange" value="<%=sqlRange%>" type="hidden">
							<input name="strView" value="<%=strView%>" type="hidden">
							<input type='submit' name='btnSubmit' value='Click here!'>
						</form>
						<br>
						Go back to my personal gClock Keeper:&nbsp;&nbsp;
						<form name='frmUID'  method='post' action='logna.asp?elogger'>
							<input name="strUID" value="<%=strUID%>" type="hidden">
							<input name="sqlRange" value="<%=sqlRange%>" type="hidden">
							<input name="strView" value="<%=strView%>" type="hidden">
							<input type='submit' name='btnSubmit' value='Click here!'>
						</form>
						
					</td>
					<!-- END For Export Logs-->
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align='center' valign='middle'>
			<table cellpadding="5" cellspacing="0" border="2" bordercolor="blue" bordercolordark="blue" bordercolorlight="blue" width="100%"> 
				<tr>
						<% 
						'open the database
						'sqlQuery = "Select DateInOut, TimeIn, TimeOut, Late, Notes from TimeInOut where  UserID = '" & strUID & "'" & sqlRange & " order by DateInOut;"
						sqlQuery = "SELECT UserInfo.FirstName, UserInfo.LastName, TimeInOut.TimeIn, TimeInOut.TimeOut, TimeInOut.Late, UserInfo.Dept, TimeInOut.DateInOut, TimeInOut.Notes  FROM (TimeInOut INNER JOIN UserInfo ON TimeInOut.UserID = UserInfo.UserID) where	TimeInOut.DateInOut	" & sqlRange & "  order by TimeInOut.DateInOut;"
						
						set dbRset = server.CreateObject ("ADODB.Recordset")
						dbRset.open sqlQuery, myDSN
						 'This is the the code which tells the page to open Excel and give it the data to display
						' Response.ContentType = "application/vnd.ms-excel"
						 'You can give the spreadsheet a name at the point its produced
						' Response.AddHeader "Content-Disposition", "attachment; filename=gclockkeeper_logs_"& Date &".xls" 
						
						if dbRset.EOF then
							Response.Write "<td colspan=4>No records yet."
							Response.Write "</td></tr></table>"
						else
							Response.Write "<td><b>FirstName</b></td>"
							Response.Write "<td><b>LastName</b></td>"
							Response.Write "<td><b>Time In</b></td>"
							Response.Write "<td><b>Time Out</b></td>"
							Response.Write "<td><b>Late</b></td>"
							Response.Write "<td><b>Department</b></td>"
							Response.Write "<td><b>Date</b></td>"
							Response.Write "<td><b>Notes</b></td></tr><tr><td>"
							response.write dbRset.getstring (,,"</td><td>","</td></tr><tr><td>"," - - -")
							Response.Write "</td></tr></table>"
						end if
						call CloseRset()
						%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!--#include file="baba.htm"-->
</body>
</html>