<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckLogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ҳ</title>
</head>

<body>
<%Response.Write(Session("username"))%>
<h1>��ӭʹ��</h1>
<a href="UnVerifyList.asp"target="rightFrame">����
<%
username=Session("username")
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select count(*) as allresult from  Record where  verifier='"&username&"' and isverify=false"
rs.open exec,conn,1,1
Response.write rs("allresult")

rs.Close
set rs =nothing
conn.Close
set conn=nothing
%>
��δȷ�ϼ�¼,�뼰ʱ����!!!
</a>
</body>
</html>
