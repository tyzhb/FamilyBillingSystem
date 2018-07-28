<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
%>

<%
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select * from users"
rs.open exec,conn,1,1
%>
<table width="30%" border="1">
  <tr>
    <td>姓名</td>
    <td>密码</td>
  </tr>
  <%do while not rs.eof%>
  <tr>
    <td><%=rs("UserName")%></td>
    <td><%=rs("Passwd")%></td>
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
</body>
</html>
