<%
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")

buytime=request.form("buytime")
content=request.form("content")
cash=request.form("cash")
buyer=request.form("buyer")
'filltime=request.form("filltime")
filltime=now()
'verifier=request.QueryString("verifier")
verifier=request.form("verifier")
isdeal=request.form("isdeal")
exec="insert into record (buytime,content,cash,buyer,filltime,verifier,isdeal)values('"&buytime&"','"&content&"','"&cash&"','"&buyer&"','"&filltime&"','"&verifier&"','"&isdeal&"')"
response.Write("alert('"&exec&"')")
conn.execute exec
conn.close
set conn=nothing
response.Redirect"Record.asp"
%>