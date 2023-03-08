<!-- #include file = "../inc/getmenu.asp"-->
<%
	Call freeAdmininput
	Call freeRole
	Call freeRoleAss
	Call freelistRole
	session.Abandon()
	Response.Redirect("login.asp")
%>