<!-- #include file = "class/CEmployee.asp"-->
<!-- #include file = "inc/createtemplate.inc"-->
<!-- #include file = "inc/constants.inc"-->
<!-- #include file = "inc/getmenu.asp"-->
<!-- #include file = "inc/library.asp"-->
<!-- #include file="class/clsSHA-1.asp" -->
<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("message.htm")
	End If
	
	Call freeListpro
	Call freeProInfo
	Call freeAssignment
	Call freeAssignRight
	Call freeListEmp
	Call freeShort
	Call freeSinglepro
	Call freeSumpro
	
	'Get Full Name
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"	
	Set objEmployee = Nothing
	
	'Make list of menu
	If isEmpty(session("Menu")) then
		getRes = getarrMenu(session("USERID"))
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End if



	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
'Response.Write UBound(getRes,2)

	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@content", "")
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="timesheet.css">

<script language="javascript" src="library/library.js"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>

			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(1))

			%>		

			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
	<form method="post" name="temp">
		<input type="hidden" value="" name="temp">
	</form>
</body>
</html>