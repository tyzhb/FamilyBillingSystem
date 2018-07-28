<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckLogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>填写账单</title>
</head>
<body>
<script type="text/javascript">
	function check(){
		var verifier = document.getElementById('verifier').value;
		if(verifier=='--请选择--'||verifier==''){
			alert("请选择确认人!!!");
			return false;
		}
		return true;
	}
</script>

<form action="add.asp" method="post" name="form1" id="form1" onsubmit="return check()">
<p>购买时间:<input type="text" name="buytime"  />yyyy-mm-dd hh:mm:ss</p>
<p>内容:<input type="text" name="content"  /></p>
<p>金额:<input type="text" name="cash"  /></p>
<p>购买人:<input type="text" name="buyer" value=<%=Session("username")%>></p>
<%
set conn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
%>

<%
username=Session("username")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("history.MDB")
exec="select * from Users where username <> '"&username&"'"
rs.open exec,conn,1,3
%>
<input type="hidden" name="verifier" value="" id="verifier">
<p>确认人:<select name="select1" onBlur="JavaScript:document.form1.verifier.value=document.form1.select1.options[selectedIndex].text" style="width:120px">
<option value="0" selected="selected">--请选择--</option>
  <%
  	do while not rs.eof
  %>
<option value="<%=rs("username")%>"><%=rs("username")%></option>
<%rs.movenext 
  loop%>
</select></p>
<p>是否处理:<input type="text" name="isdeal" value='0' readonly="readonly" />请填写0,代表未处理</p>
<input name="submit" type="submit" value="提交" />
</form>
</body>
</html>
