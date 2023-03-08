<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->
<!-- #include file = "../../../inc/library.asp"-->

<%
	Dim strUserName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError, intPageSize, fgRight 'view all or Not
	Dim varEmp, varFullName

'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../../message.htm")
	End If					

	intUserID = session("USERID")
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 
	
	If objDatabase.dbConnect(strConnect) Then

'--------------------------------------------------
' Check right on page
'--------------------------------------------------
			
		strSQL = "exec sp_CheckRight " & intUserID
		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.noRecord Then				 
				strError1 = "You have no access rights."
			End If
			objDatabase.closeRec
		Else
			strError = "1. " & objDatabase.strMessage
		End If
	Else
		strError = "2. " & objDatabase.strMessage
	End If
	
'--------------------------------------------------
' End of checking right on page
'--------------------------------------------------

'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
	
'--------------------------------------------------
' Make list of menu
'--------------------------------------------------

	If isEmpty(session("Menu")) Then 
		getRes = getarrMenu(intUserID)
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End If	
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AD"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", "&nbsp;")
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", "&nbsp;")	
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../../timesheet.css">
<script language="javascript" src="../../../library/library.js"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>

<%	If strError1 = "" Then%>        

<SCRIPT language="javascript">
	window.document.frmreport.action = "rpt_sal_staff.asp";
	window.document.frmreport.submit();
</SCRIPT>

<%	Else%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<%          
		If strError <> "" Then
%>               
		  <tr>
			<td class="red">&nbsp;<%=strError%></td>
		  </tr>
<%		End If%>				

		  <tr>
         	<td class="red" align="center" valign="middle"><b><%=strError1%></b></td>
		  </tr>	          
        </table>
<%	End If
	Set objDatabase = Nothing

'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------

	Response.Write(arrTmp(1))

'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(2))    
%>

</form>
</body>
</html>