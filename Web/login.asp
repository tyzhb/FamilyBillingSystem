<%
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
set rs=server.createobject("adodb.recordset")

username=request.form("username")
passwd=request.form("passwd")

exec="select * from Users where (username='"&username&"'and passwd='"&passwd&"')"
rs.open exec,conn

if not rs.eof then
rs.Close
conn.Close
session("checked")="yes"
session("username")=username
response.Redirect "main.asp"
else
session("checked")="no"
response.Redirect "index.asp"
end If

%>
