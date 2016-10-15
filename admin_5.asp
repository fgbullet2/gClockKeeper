
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1252">
<TITLE>gClock Keeper</TITLE>
</HEAD>
<BODY>
<%

'dim sqlRange
'sqlRange = ""
'Response.Buffer = true
'strQuery = Request.QueryString 
'dim strError
'if strQuery = "" then
'	Response.Redirect ("main.asp")
'else
	'fill up the variables
'	dim strUID, msgStatus
'	strUID = Request.Form("strUID")
'	if strQuery="elogger2" and strUID = "" then
	'if not strDept = "Admin" then
'		Response.Redirect ("main.asp")
'	end if

'	end if
%>
 
<%
Param = Request.QueryString("Param")
Data = Request.QueryString("Data")
%>
<%
If IsObject(Session("eloggerdb_conn")) Then
    Set conn = Session("eloggerdb_conn")
Else
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open "eloggerdb","",""
    Set Session("eloggerdb_conn") = conn
End If
%>
<%
    sql = "SELECT UserInfo.FirstName, UserInfo.LastName, TimeInOut.TimeIn, TimeInOut.TimeOut, TimeInOut.Late, UserInfo.Dept, TimeInOut.DateInOut  FROM TimeInOut INNER JOIN UserInfo ON TimeInOut.UserID = UserInfo.UserID  "
    If cstr(Param) <> "" And cstr(Data) <> "" Then
        sql = sql & " WHERE [" & cstr(Param) & "] = " & cstr(Data)
    End If
    sql = sql & " ORDER BY TimeInOut.DateInOut    "
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
%>
<table border="1" cellspacing="0" align="CENTER" valign="MIDDLE" bgcolor="White"><FONT FACE="Arial" COLOR=#000000><CAPTION><B>Elogger Summary Report</B></CAPTION>

<THEAD>
<TR>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>FirstName</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>LastName</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>TimeIn</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>TimeOut</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>Late</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>Dept</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT SIZE=2 FACE="Arial" COLOR=#000000>DateInOut</FONT></TH>

</TR>
</THEAD>
<TBODY>
<%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof
 %>
<TR VALIGN=TOP>
<TD BORDERCOLOR=#c0c0c0 ><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("FirstName").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0 ><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("LastName").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0  ALIGN=RIGHT><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("TimeIn").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0  ALIGN=RIGHT><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("TimeOut").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0 ><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Late").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0 ><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Dept").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0  ALIGN=RIGHT><FONT SIZE=2 FACE="Arial" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("DateInOut").Value)%><BR></FONT></TD>

</TR>
<%
rs.MoveNext
loop%>
</TBODY>
<TFOOT></TFOOT>
</TABLE>
<table align="CENTER" valign="MIDDLE">
<tr>
<td><br>
<a href="e_logger_4_admin.xls" target="_parent"> Export Report </a>
</td>
</tr>
</table>
<!--#include file="baba.htm"-->
</BODY>
</HTML>