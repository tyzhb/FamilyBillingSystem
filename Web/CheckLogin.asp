
<%
	if Session("checked") ="yes" Then
%>
<%
	'do nothing
%>
<%
	Else
		Response.Write("<script language='javascript'>alert('���ȵ�½!!!');window.location = 'index.asp';</script>")
	End If
%>

