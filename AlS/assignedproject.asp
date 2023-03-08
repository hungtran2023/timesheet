<!-- #include file = "class/CEmployee.asp"-->
<!-- #include file = "inc/createtemplate.inc"-->
<!-- #include file = "inc/getmenu.asp"-->
<!-- #include file = "inc/constants.inc"-->
<!-- #include file = "inc/library.asp"-->
<%
'**************************************************
' Sub: CopyDataPrivate
' Description: 
' Parameters: 
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Sub CopyDataPrivate(ByRef rsSrc, ByRef rsDes)
	rsSrc.MoveFirst
	Do While not rsSrc.EOF
	    rsDes.AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID"), _
	 					Array(rsSrc(0), rsSrc(1), rsSrc(2), rsSrc(3), rsSrc(4), rsSrc(5))
		rsSrc.MoveNext
	Loop
End Sub
'**************************************************
' Sub: Extract
' Description: make a recordset, then copy data into
' Parameters: 
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Sub Extract(ByRef rsSrc, ByRef rsDes)
	if rsDes.RecordCount>0 then
	  set rsDes = Server.CreateObject("ADODB.Recordset")
	  Call SetAttRsPrivate(rsDes)
	end if
	Call CopyDataPrivate(rsSrc, rsDes)
End Sub
'**************************************************
' Sub: SetAttRsPrivate
' Description: initial a recordset
' Parameters: 
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Sub SetAttRsPrivate(ByRef rsSrc)
	rsSrc.CursorLocation = adUseClient     ' Set the Cursor Location to Client

  ' Append some Fields to the Fields Collection
	rsSrc.Fields.Append "pID", advarChar, 20
	rsSrc.Fields.Append "pName", adVarChar, 120
	rsSrc.Fields.Append "sID", adInteger
	rsSrc.Fields.Append "sName", adVarChar, 150
	rsSrc.Fields.Append "sParentID", adInteger,,adFldIsNullable
	rsSrc.Fields.Append "chainID", adVarChar, 100,adFldIsNullable
	rsSrc.CursorType = adOpenStatic
	rsSrc.Open
End Sub
'**************************************************
' Function: AppendTree
' Description: 
' Parameters: projectID, ProjectName, TaskName, level
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Function AppendTree (ByVal strpID, ByVal strpName, ByVal strsName, ByVal intLevel)
Dim strTmp, i
	strTmp = ""
	strOut = Showlabel(strsName)
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "<IMG alt='' border='0' height='18' src='images/t_dot.gif' width='36'>"
		Next
		strTmp = strTmp & "<IMG alt='' border='0' src='images/dot1.gif'>"
		strTmp = strTmp & "<IMG alt='' border='0' height='10' width='12' src='images/nosign.gif'>"
	Else
		strOut = Showlabel(strpID) & " - " & Showlabel(strpName)
	End If
	AppendTree = strTmp & strOut & "<BR>"
End Function
'**************************************************
' Sub: FetchChild
' Description: this is a recursive function
' Parameters: 
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Sub FetchChild(ByRef rsGet, ByRef strTree, ByVal intLevel)
Dim strName, intContinue
	Do Until rsGet.EOF
	  strTree = strTree & AppendTree(rsGet("pID"), rsGet("pName"), rsGet("sName"), intLevel)

	  objData.Filter = "sParentID = " & rsGet("sID")

	  intContinue = 0
	  If objData.RecordCount > 0 then
		intContinue = 1
		Call Extract(objData, arrRs(intLevel+1))
		 arrRs(intLevel+1).MoveFirst
	  End If
	  objData.Filter = ""
	  If intContinue = 1 Then
		FetchChild arrRs(intLevel+1), strTree, intLevel+1
	  End if
	  rsGet.MoveNext
	Loop
End Sub

	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage
	
	gMessage = ""
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("message.htm")
	End If					
	
'-------------------------------
' Calculate pagesize
'-------------------------------
	if not isEmpty(session("Preferences")) then
		arrPre = session("Preferences")
		if arrPre(1, 0)>0 then PageSize = arrPre(1, 0) else PageSize = PageSizeDefault
		set arrPre = nothing
	else
		PageSize = PageSizeDefault
	end if
'--------------------------------------
' Get Full Name
'--------------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID")) 'strUserName)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
'--------------------------------------
' Make list of menu
'--------------------------------------
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
	
	If IsEmpty(Session("strHTTP")) Then 
		Call MakeHTTP
	End if
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "")


'-----------------------------------------------------------------------------
'making custom record set, get all Sub Task and Parent Task of one employee
'-----------------------------------------------------------------------------
Dim objData
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
	if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
		Call freeListpro
		Call freeListEmp
		Call freeProInfo
		Call freeAssignment
		Call freeAssignRight
		Call freeShort
		Call freeSinglepro
		Call freeSumpro
	end if
end if

if fgExecute and isEmpty(session("arrBookmark")) then
	set objData = Server.CreateObject("ADODB.Recordset") ' Create the ADO Object
	Call SetAttRsPrivate(objData)

	strAncestor = ""
	strConnect = Application("g_strConnect")
	  
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "select b.projectID, c.ProjectName, b.SubTaskID, b.SubTaskName, ISNULL(b.taskID, 0), ISNULL(b.ChainID, '') FROM " &_
					"((select * from ATC_Assignments where fgDelete = 0 AND StaffID = " & session("USERID") & ") a " &_
					"INNER JOIN ATC_Tasks b ON b.SubTaskID = a.SubTaskID) " &_
	 				"INNER JOIN ATC_Projects c ON c.ProjectID = b.ProjectID WHERE c.fgDelete = 0"

	  If objDb.runQuery(strQuery) Then
	    If not objDb.noRecord Then
			'-----Copy data------------------------
			objDb.MoveFirst

			Do While not objDb.rsElement.EOF

			  objData.AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID"), _
							Array(objDb.rsElement(0), objDb.rsElement(1), objDb.rsElement(2), _
							objDb.rsElement(3), objDb.rsElement(4), objDb.rsElement(5))
			  strAncestor = strAncestor & objDb.rsElement(5)
			  
			  objDb.MoveNext
			Loop
'response.write strAncestor					
'response.end			
			'--------------------------------------
			If strAncestor <> "" Then
				strAncestor = Mid(strAncestor, 1, Len(strAncestor)-1)
				strQuery = "select a.projectID, b.ProjectName, a.SubTaskID, a.SubTaskName, ISNULL(a.taskID, 0), ISNULL(a.ChainID, '') FROM " &_
							"ATC_Tasks a INNER JOIN ATC_Projects b ON b.ProjectID = a.ProjectID " &_
	 						"WHERE a.SubTaskID IN (" & strAncestor & ") order by a.SubTaskID" 

				If objDb.runQuery(strQuery) Then
				  Call CopyDataPrivate(objDb.rsElement, objData)
				Else
				  gMessage = objDb.strMessage
				End if
			End If
		Else
		  gMessage = "No assignment."
		End if
	  Else
	    gMessage = objDb.strMessage
	  End if
	  objDb.dbDisConnect
	Else
	  gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
End if


'----------------Begin analyse-------------------
strLast = ""
If fgExecute and gMessage="" and isEmpty(session("arrBookmark")) then
	Dim arrRs(4)
	  '-- Create the ADO Objects
	For i = 0 to 4
	  set arrRs(i) = Server.CreateObject("ADODB.Recordset")
	  Call SetAttRsPrivate(arrRs(i))
	Next

	objData.Sort = "pID"
	objData.Filter = "sParentID = 0"
	set objRoot = Server.CreateObject("ADODB.Recordset")
	Call SetAttRsPrivate(objRoot)
	Call CopyDataPrivate(objData, objRoot)
	objData.Filter = ""
	k = 0
	strTree = ""
	objRoot.MoveFirst
	'--------------------------------------
	' loop for every project
	'--------------------------------------
	strHead = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & chr(13) & _
		"  <tr><td bgcolor='#617DC0'>" &_
		"<table width='100%' border='0' cellspacing='1' cellpadding='5'>" & chr(13)
	intNumofLine = 0
	CurPage = -1
	Dim arrBookmark()
	strLast = ""
	Do Until objRoot.EOF
		k = k + 1
		arrRs(0).AddNew Array("pID", "pName", "sID", "sName", "sParentID", "chainID"), _
						Array(objRoot(0), objRoot(1), objRoot(2), objRoot(3), objRoot(4), objRoot(5))
		FetchChild arrRs(0), strTree, 0
	'--------------------------------------	
	' Reset all of recordset
	'--------------------------------------
		For i = 0 to 4
		  arrRs(i).Close
		  Set arrRs(i) = Nothing
		  set arrRs(i) = Server.CreateObject("ADODB.Recordset")
		  Call SetAttRsPrivate(arrRs(i))
		Next
		
		If k mod 2 = 1 Then
			strColor = "#E7EBF5"
		Else
			strColor = "#FFF2F2"
		End If
		
		intNumofLine = intNumofLine + LineCount(strTree)
		strTmp = "<tr bgcolor='" & strColor & "'><td valign='top' class='blue'>" & strTree & "</td></tr>"
		strLast = strLast & strTmp & chr(13)
		if intNumofLine > PageSize then 'go to a new page
			strLast = strHead & strLast & "</table></td></tr></table>"
			Curpage = Curpage + 1
			Redim Preserve arrBookmark(Curpage)
			arrBookmark(Curpage) = strLast
			strLast = ""
			intNumofLine = 0
		end if		
		strTree = ""
	  	objRoot.MoveNext
	Loop
	'Last page or only one page
	if strLast<>"" then
		strLast = strHead & strLast & "</table></td></tr></table>"
		Curpage = Curpage + 1
		Redim Preserve arrBookmark(Curpage)
		arrBookmark(Curpage) = strLast
	end if
	session("arrBookmark") = arrBookmark
	session("CurpageAssigned") = 1
	session("NumPageAssigned") = CurPage + 1
	
	objData.Close
	objRoot.Close
	For i = 0 to 4
		arrRs(i).Close
		Set arrRs(i) = Nothing
	Next
	set objData = nothing
	set objRoot = nothing
else
	arrBookmark = session("arrBookmark")
End if

if fgExecute then
	varNavi = Request.QueryString("navi")
	if varNavi<>"" then
		tmpi = Session("CurPageAssigned")
		select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPageAssigned") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPageAssigned")
			end if
		End select
		Session("CurPageAssigned") = tmpi
	end if

	varGo = Request.QueryString("Go")
	if varGo <> "" then Session("CurPageAssigned") = CInt(varGo)
end if 'fgexecute

strLast = "&nbsp;"
if isArray(arrBookmark) then
	strLast = arrBookmark(session("CurpageAssigned") - 1)
else
	Session("CurPageAssigned") = 0
	Session("NumPageAssigned") = 0
end if
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="timesheet.css">
<script language="javascript" src="library/library.js"></script>
<script>
function next() {
var curpage = <%=session("CurPageAssigned")%>
var numpage = <%=session("NumPageAssigned")%>
	if (curpage < numpage) {
		document.frmAssigned.action = "assignedproject.asp?navi=NEXT"
		document.frmAssigned.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageAssigned")%>
var numpage = <%=session("NumPageAssigned")%>
	if (curpage > 1) {
		document.frmAssigned.action = "assignedproject.asp?navi=PREV"
		document.frmAssigned.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageAssigned")%>
	var curpage = <%=session("CurPageAssigned")%>
	var intpage = frmAssigned.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmAssigned.action = "assignedproject.asp?Go=" + intpage
		document.frmAssigned.target = "_self"
		document.frmAssigned.submit();		
	}
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="frmAssigned">
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
			Response.Write(arrTmp(0))
			%>

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
  <td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	  <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
        <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
      </tr>
	  <tr align="center"> 
	    <td class="title" height="10" align="center">&nbsp;</td>
	  </tr>
	  <tr valign="millde">
	    <td class="title" height="50" align="center">Activities 
	      List </td>
	  </tr>
	</table>
  </td>
  </tr>
 <tr valign="top"> 
    <td height="100%" valign="top">
	  	<table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365" >
			<tr> 
			  <td bgcolor="#FFFFFF" valign="top">
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(strLast)
			%>
			 </td>
		   </tr>
		</table>
    </td>
 </tr>
 <tr align="right" bgcolor="#99A89D"> 
    <td height="20" align="right" bgcolor="#E7EBF5">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
		  <tr> 
		    <td align="right" bgcolor="#E7EBF5"> 
		      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
		        <tr class="black-normal"> 
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageAssigned")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPageAssigned")%>/<%=session("NumPageAssigned")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Go to previous page'; return true;" onMouseOut="self.status=''">Previous</a> /
					<a href="javascript:next();" onMouseOver="self.status='Go to next page'; return true;" onMouseOut="self.status=''"> Next</a>&nbsp;&nbsp;&nbsp;</td>
		        </tr>
		      </table>
		    </td>
		  </tr>
		</table>
    </td>
 </tr>
</table>
			<%
			Response.Write(arrTmp(1))
			%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>
</body>
</html>