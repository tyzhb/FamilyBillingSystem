
<%
	if Session("checked") ="yes" Then
%>
<%
	'do nothing
%>
<%
	Else
		Response.Write("<script language='javascript'>alert('гКох╣гб╫!!!');window.location = 'index.asp';</script>")
	End If
%>

