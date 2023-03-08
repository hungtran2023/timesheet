<!-- #include file = "inc/getmenu.asp"-->
<%
	Response.Cookies("assignID") = ""
	Response.Cookies("introw") = ""
    Response.Cookies("mySession") = ""
	Response.Cookies("SessionExpired") = ""
	 
	Call freeListpro
	Call freeProInfo
	Call freeAssignment
	Call freeAssignRight
	Call freeListEmp
	Call freeShort
	Call freeSinglepro
	Call freeSumpro
	Dim SessionSharing
    Set SessionSharing = server.CreateObject("SessionMgr.Session2")
    SessionSharing("USERID") =  ""
    session.Abandon()

	Response.Redirect("initial.asp")
%>