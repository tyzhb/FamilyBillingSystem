<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckLogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��¼</title>
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
    <td>����ʱ��</td>
	<td>����</td>
	<td>���(Ԫ)</td>
	<td>������</td>
	<td>��дʱ��</td>
	<td>ȷ����</td>
	<td>�Ƿ��Ѵ���</td>
	<td>�Ƿ���ȷ��</td>
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
Response.write "�ܶ�:"
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
Response.write "�Ѵ����ܶ�:"
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
Response.write "δ�����ܶ�:"
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
