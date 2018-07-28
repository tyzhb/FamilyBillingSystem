<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
%>

<%
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.mdb")
exec="select * from users"
rs.open exec,conn,1,1
%>
<table width="30%" border="1">
  <tr>
    <td>аеУћ</td>
    <td>УмТы</td>
  </tr>
  <%do while not rs.eof%>
  <tr>
    <td><%=rs("UserName")%></td>
    <td><%=rs(Passwd)%></td>
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