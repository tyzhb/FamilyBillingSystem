<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckLogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>未处理记录</title>
</head>

<body>
<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
%>

<%
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select * from Record where Isdeal=false"
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
	<% if (Session("username")="张恒斌") Then 
	response.Write("<td>处理</td>")
	 end if %>
	
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
	<% if (Session("username")="张恒斌") Then
		id=rs("id")
		response.Write("<td><form action='dealrecord.asp' name='dealrecord' method='post'><input type='hidden' value='"&id&"' id='id' name='id'/><input type='submit'value='处理'/></form></td>") 
	%>
	<% end if %>
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
exec="select sum(Cash) as allresult from  Record where isdeal=0"
rs.open exec,conn,1,1
Response.Write "未处理总额:"
Response.write rs("allresult")
Response.Write "<br/>"

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
exec="select buyer,sum(Cash) as allresult from  Record where isdeal=0 group by buyer" '每个人未处理总额
rs.open exec,conn,1,1
%>
 <%
  	do while not rs.eof
	
	Response.Write rs("buyer")
	Response.Write ":"
	Response.write rs("allresult")
	Response.Write "  "
	
  %>

 <%rs.movenext 
  loop%>
<%
rs.Close
set rs =nothing
conn.Close
set conn=nothing
%>

</body>
</html>
