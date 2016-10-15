
<%
Option Explicit
'variable definitions
dim strRemHost			'remote host's ip address
dim strQuery			'the query string

'database variables
dim myDSN				'the DSN connection string
dim sqlQuery			'the sql query string
dim dbConn				'variable for handling dbconnection
dim dbRset				'variable for handling dbrecordset

'dsn connection string
myDSN = "eloggerdb"
'subs and functions
sub CloseConn()
	dbConn.Close
	set dbConn  = nothing
end sub

sub CloseRset()
	dbRset.Close
	set dbRset = nothing
end sub
%>