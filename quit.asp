<% Session.Abandon 

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
end if

strQuery = ""
sqlRange = ""
strUID = ""
	
%> 
