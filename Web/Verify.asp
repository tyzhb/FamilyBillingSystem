<%
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")

id=request.form("id")
response.Write("hello"+id)
response.Write(id)
exec="update Record set IsVerify='1' where id="&id&""
conn.execute exec
conn.close
set conn=nothing
response.Redirect"UnVerifyList.asp"
%>