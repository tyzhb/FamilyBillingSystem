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
<h2>��ӭʹ��</h2>
<br>
<a href="Record.asp" target="rightFrame">�鿴��¼</a>
<br>
<a href="FillRecord.asp" target="rightFrame">��д��¼</a>
<br />
<a href="UnVerifyList.asp"target="rightFrame">�ҵ�δȷ�ϼ�¼</a>
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
��
<br />
<% if (Session("username")="�ź��") Then 
	response.Write("<a href='UndealRecord.asp' target='rightFrame'>�����¼</a>")
	end if 
%>
<br />
<a href="MyHistoryList.asp"target="rightFrame">�ҵļ�¼�鿴</a>
<br />
<a href="HistoryList.asp"target="rightFrame">��ʷ��¼�鿴</a>
</body>
</html>
