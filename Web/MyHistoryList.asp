<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckLogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>记录</title>
</head>

<body>
<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
%>

<%
username=Session("username")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select * from Record where buyer='"&username&"'"
rs.open exec,conn,1,3
%>
<table width="100%" border="1">
  <tr>
    <td>id</td>
    <td>购买时间</td>
	<td>内容</td>
	<td>金额(元)</td>
	<td>购买人</td>
	<td>填写时间</td>
	<td>确认人</td>
	<td>是否已处理</td>
	<td>是否已确认</td>
  </tr>
  <%
  	do while not rs.eof
  %>
  <tr>
    <td><%=rs("id")%></td>
    <td><%=rs("BuyTime")%></td>
	<td><%=rs("Content")%></td>
	<td><%=rs("Cash")%></td>
	<td><%=rs("Buyer")%></td>
	<td><%=rs("FillTime")%></td>
	<td><%=rs("Verifier")%></td>
	<td><%=rs("IsDeal")%></td>
	<td><%=rs("IsVerify")%></td>
  </tr>
  <%rs.movenext 
  loop%>
</table>

<%
rs.Close
set rs =nothing
conn.Close
set conn=nothing
%>

<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select sum(Cash) as allresult from  Record where buyer='"&username&"'"
rs.open exec,conn,1,1
Response.write "总额:"
Response.write rs("allresult")
Response.write "  "

%>
<%
rs.Close
set rs =nothing
conn.Close
set conn=nothing
%>


<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select sum(Cash) as allresult from  Record where isdeal=true and buyer='"&username&"'"
rs.open exec,conn,1,1
Response.write "已处理总额:"
Response.write rs("allresult")
Response.write "  "

%>
<%
rs.Close
set rs =nothing
conn.Close
set conn=nothing
%>

<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select sum(Cash) as allresult from  Record where isdeal=0 and buyer='"&username&"'"
rs.open exec,conn,1,1
Response.write "未处理总额:"
Response.write rs("allresult")
Response.write "  "

%>
<%
rs.Close
set rs =nothing
conn.Close
set conn=nothing
%>


</body>
</html>
