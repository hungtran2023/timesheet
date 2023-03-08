<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' function: BuildTaskList
' Description: draw tree of subtask
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub BuildTaskList(byval rsTaskClone,byval blnRightParent,byval SubTaskID,byval intLevel, byref strReturn)
	
	dim blnRight

	rsTaskClone.Filter = "ParentID = " & SubTaskID	
	If rsTaskClone.recordCount>0 then
		Do while not rsTaskClone.EOF
			blnRight=blnRightParent OR (rsTaskClone("OwnerID")=Session("UserID")) OR (rsTaskClone("StaffID")=Session("UserID")) OR fgDelegate
			strReturn=Replace(strReturn,"©","")
			strReturn=strReturn & AppendTree(rsTaskClone("SubtaskName"),intlevel,blnRight and (intLevel<>0))
			call BuildTaskList(rsTaskClone.Clone,blnRight,rsTaskClone("SubTaskID"),intLevel + 1, strReturn)
			rsTaskClone.MoveNext
		loop
	else
		strButton="<a href='javascript:assignment(" & SubTaskID & ");'>Assign</a>"
		if blnRightParent then 
			strReturn=Replace(strReturn,"©",strButton)
		else
			strReturn=Replace(strReturn,"©","")
		end if
		
		strReturn=strReturn & AppendUser(rsParticipant.Clone,SubTaskID,intLevel,blnRightParent)
	end if
end Sub

'****************************************
' function: BuildListCombobox
' Description: draw tree of subtask
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub BuildListCombobox(byval rsTaskClone,byval blnRightParent,byval SubTaskID,byval intLevel, byref strListTask)
	
	dim blnRight

	rsTaskClone.Filter = "ParentID = " & SubTaskID	
	If rsTaskClone.recordCount>0 then
		Do while not rsTaskClone.EOF
			blnRight=blnRightParent OR (rsTaskClone("OwnerID")=Session("UserID")) OR (rsTaskClone("StaffID")=Session("UserID")) OR fgDelegate
			strListTask = strListTask & AppendListAssignment(rsTaskClone("SubtaskName"),intLevel , rsTaskClone("SubTaskID"))
			call BuildListCombobox(rsTaskClone.Clone,blnRight,rsTaskClone("SubTaskID"),intLevel + 1, strListTask)
			rsTaskClone.MoveNext
		loop
	
	end if
end Sub
'****************************************
' function: AppendListAssignment
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function AppendListAssignment (ByVal strsName, ByVal intLevel, ByVal intValue)
Dim strTmp, i, strColor
	strTmp = ""
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "&nbsp;&nbsp;"
		Next
	End If
	
	AppendListAssignment = "<option value='" & intValue & "'>" & strTmp & "* " & showlabel(strsName) & "</option>"
	
End Function
'****************************************
' function: appendUser
' Description: draw tree of user
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function AppendUser (Byval Rs, byval TaskID, ByVal intLevel, ByVal blnShow)
Dim strTmp, strIndent, strColor
	strTmp = ""
	If intLevel > 0 Then		
		strIndent = "<IMG alt='' vspace='0' border='0' height='10' src='../../images/t_dot.gif' width='" & (intLevel*36) & "'>"
	Else
		strIndent = "<IMG alt='' vspace='0' border='0' height='10' src='../../images/t_dot.gif' width='5'>"
	End If	

	Rs.Filter = "SubtaskID = " & TaskID
	Do Until Rs.EOF
	  
		strColor = "#E7EBF5"
		if blnShow then
		  strTmp = strTmp & "<tr bgcolor='" & strColor & "'><td class='blue-normal'>" & strIndent & Showlabel(rs("FullName")) &_
		  			"</td><td class='blue-normal' width='200'>" & Showlabel(rs("JobTitle")) & "</td>" &_
		  			"<td class='blue-normal' width='24'><input type='checkbox' name='chkpar' value='" & Rs("AssignmentID") & "' " & strCHK & "></td></tr>" & chr(13)
		else
		  strTmp = strTmp & "<tr bgcolor='" & strColor & "'><td class='blue-normal'>" & strIndent & Showlabel(rs("FullName")) &_
		  			"</td><td class='blue-normal' width='200' colspan='2'>" & Showlabel(rs("JobTitle")) & "</td>" &_
		  			"</tr>" & chr(13)
		end if
		Rs.MoveNext
	Loop
	Rs.Filter=""
	AppendUser = strTmp
End Function

'****************************************
' function: appendtree
' Description: draw tree of user
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function AppendTree (ByVal strsName, ByVal intLevel, ByVal blnShow)
Dim strTmp, i, strColor
	strTmp = ""
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "<IMG alt='' vspace='0' border='0' height='10' src='../../images/t_dot.gif' width='36'>"
		Next
		strTmp = strTmp & "<IMG alt='' vspace='0' border='0' src='../../images/dot1.gif'>"
		strTmp = strTmp & "<IMG alt='' vspace='0' border='0' height='10' width='12' src='../../images/nosign.gif'>"
	End If	
	LineOnPage = LineOnPage + 1
	strColor = "#FFF2F2"
	if blnShow then
		AppendTree = "<tr bgcolor='" & strColor & "'><td colspan='2' class='blue'>" & strTmp & Showlabel(strsName) & "</td><td>©</td></tr>" & chr(13)
	else
		AppendTree = "<tr bgcolor='" & strColor & "'><td colspan='2' class='black'>" & strTmp & Showlabel(strsName) & "</td><td align='center'>©</td></tr>" & chr(13)
	end if
End Function

'****************************************
' function: GetData
' Description: draw tree of user
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub GetData (ByVal strQuery, byref rs)
Dim objDb,strConnect
	
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
			
	If objDb.openRec(strQuery) Then
		objDb.recDisConnect
		set rs = objDb.rsElement.Clone

		objDb.CloseRec
	Else
		gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
	
End sub 

'****************************************
' function: Get TMS
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function GetTMS_Sql 
	Dim rsIndex
	dim strSql
	strSql="SELECT AssignmentID FROM ATC_Timesheet"
	call GetData("SELECT TMS_Table FROM ATC_Index",rsIndex)
	if rsIndex.RecordCount>0 then
		Do while not rsIndex.EOF
			strSql = strSql & " UNION ALL SELECT AssignmentID from " & rsIndex(0)
			rsIndex.MoveNext
		loop	
	end if
	set rsIndex=nothing
	GetTMS_Sql=strSql
End Function

'****************************************
' function: Check Del
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function CheckDel (byval AssignmentID)
	Dim rsIndex
	dim strSql
	strSql="SELECT Count(*) FROM (" & GetTMS_Sql() & ") as a WHERE AssignmentID=" & AssignmentID
	call GetData(strSql,rsIndex)
	
	CheckDel=(rsIndex(0)=0)

	set rsIndex=nothing
End Function

'***************************************************************************************
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, rsParticipant, objUser, rsTask
	Dim arrRs(4), varBookMark, varBookPro, LineOnPage, varSubID, flagShow, gMessage, PageSize
	Dim fgUpdate,TaskIDforView
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	'get all Sub Task and Parent Task of one project
	proID = Request.Form("txthiddenstrproID")
	proName = Request.Form("txthiddenstrproName")
	'Response.Write proID & "--" & proName
'-----------------------------------
'Check ACCESS right
'-----------------------------------

	tmp = Request.Form("txtpreviouspage")
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgDelegate=false
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
			end if
			
			if getRight(0, ii) = "Project Delegator" then fgDelegate=true
		next
		
		if fgUpdate = false then
			strSql = "select * from atc_rightontasks a left join atc_tasks b on a.SubTaskID = b.SubTaskID " &_
					"where staffid = " & session("USERID") & " and b.Projectid = '" & proID &"'"
			Set objDb = New clsDatabase
			strConnect = Application("g_strConnect")
			if objDb.dbConnect(strConnect) then
				if objDb.runQuery(strSql) then
					if not objDb.noRecord then
						fgUpdate = true
						objDb.Closerec
					end if
				else
					gMessage = objDb.strMessage
				end if
				objDb.dbDisconnect
			else
				gMessage = objDb.strMessage
			end if
			set objDb = nothing
		end if
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if


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
	
'------------------------------------	
' Get Full Name
'------------------------------------
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
	
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'----------------------------------
' Main procedure
'----------------------------------

gMessage = ""
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
end if

TaskIDforView=Request.Form("lsttaskView")
if TaskIDforView="" then TaskIDforView=0

'--------------------------

'--------------------------
strAct = Request.QueryString("act")
if strAct = "REMOVE" then
	countU = Request.Form("chkpar").Count
	if countU>0 then
		strUpdate=""
		strDelete=""
		
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			objDb.cnDatabase.BeginTrans
			For i = 1 to countU
				varBook = int(Request.Form("chkpar")(i))
				If CheckDel(varBook) then 'delete
					strDelete = strDelete & varBook & ","
				Else 'update fgDelete
					strUpdate = strUpdate & varBook & ","
				End if
			Next
			'Response.Write strDelete & "<br>" & strUpdate & "<br>"
			if strDelete<>"" then 
				strDelete="DELETE FROM ATC_Assignments WHERE AssignmentID IN (" & Left(strDelete,len(strDelete)-1) & ")"
				if not objDb.runActionQuery(strDelete) then gMessage = objDb.strMessage
			end if
			if strUpdate<>"" then 
				strUpdate="UPDATE ATC_Assignments SET fgDelete = 1 WHERE AssignmentID IN (" & Left(strUpdate,len(strUpdate)-1) & ")"
				if not objDb.runActionQuery(strUpdate) then gMessage = gMessage & " " & objDb.strMessage
			end if
			
			if gMessage<>"" then 
				objDb.cnDatabase.RollbackTrans
			else
				objDb.cnDatabase.CommitTrans
			end if
			objDb.dbdisConnect
		else
			gMessage = objDb.strMessage 'error in connection
		end if
		set objDb = nothing
	end if  
end if
'--------------------------

'--------------------------

if proID <>"" then 'draw sub task
	strQuery="SELECT a.SubTaskID,ProjectID,SubTaskName,ISNULL(TaskID,0)as parentID,ChainID,OwnerID , ISNULL(b.StaffID ,0) as StaffID " & _
			"FROM ATC_Tasks a LEFT JOIN ATC_RightonTasks b ON (a.SubtaskID=b.SubtaskID AND b.StaffID=" & session("USERID") & ") " & _
			"WHERE projectID='" & proID & "' AND Right(SubTaskName,3)<>'-TP' ORDER BY TaskID,a.SubtaskID"									
	call GetData(strQuery,rsTask)

        strQuery="SELECT FullName,JobTitle,a.* " & _
			"FROM ATC_Assignments a " & _
				"INNER JOIN HR_Employee b ON a.StaffID=b.PersonID    " &_
			"WHERE a.fgDelete=0 AND subTaskID IN (SELECT subTaskID FROM ATC_Tasks WHERE projectID='" & proID & "') ORDER BY FullName"
							
	call GetData(strQuery,rsParticipant)
		
	if gMessage="" then
		strLast = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & chr(13) & _
						"<tr><td bgcolor='#DDDDDD'><table width='100%' border='0' cellspacing='1' cellpadding='5'>" & chr(13) &_
						"<tr bgcolor='#8CA0D1'><td align='center' class='blue' width='55%'>Fullname</td><td align='center'  class='blue' width='35%'>" &_
						"Job Title</td><td width='10%'>&nbsp;</td></tr>"	
		gLevel=0
		gRight=false

		if cint(TaskIDforView)<>0 then 
			rsTask.Find "SubtaskID =" & TaskIDforView
			gRight=((rsTask("OwnerID")=Session("UserID")) OR (rsTask("StaffID")=Session("UserID"))) OR fgDelegate
				
			strLast=strLast & AppendTree(rsTask("SubtaskName"),gLevel,gRight)
			gLevel=gLevel+1
		end if
			
		Call BuildTaskList(rsTask.Clone,gRight,cint(TaskIDforView),gLevel,strLast)	
		Call BuildListCombobox(rsTask.Clone,false,0,0,strList)

		strLast = strLast & "</table></td></tr></table>"
	end if

end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

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

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>
<script>
var objNewWindow, objSEWindow;

function assignment(taskid) { //v2.0
	
	document.frmassign.action = "selectemployee_ass.asp?outside=1&kind=1&taskid=" + taskid
	document.frmassign.target = "_self"
	document.frmassign.submit();
}

function ViewAssignment() { //v2.0
	
	//document.frmassign.action = "assignment.asp?outside=1";
	//document.frmassign.target = "_self"
	//document.frmassign.submit();
}

function checklist() {
	var tmp = document.frmassign.lsttask.selectedIndex
	if (tmp == -1) {
			alert("Please choose an item.");
			document.frmassign.lsttask.focus();
			return("");
	}
	if (document.frmassign.lsttask.options[tmp].value == "") {
			alert("This item can't be chosen.");
			document.frmassign.lsttask.focus();
			return("");
	}
	return(document.frmassign.lsttask.options[tmp].value);
}

function window_onunload() {
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.close();
	if((objSEWindow) && (!objSEWindow.closed))
		objSEWindow.close();
}

function setchecked(val) {
  with (document.frmassign) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkpar") {
			elements[ii].checked = val
		}
	}
  }
}

function chkremove() {
  fg = false;
  with (document.frmassign) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chkpar") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 if (fg == false) alert("No participant selected.")
 return(fg)
}

function remove() {
  if (chkremove()==true) {
  	document.frmassign.action = "assignment.asp?act=REMOVE"
	document.frmassign.target = "_self"
	document.frmassign.submit();
  }
}

function subtask() {
	document.frmassign.action = "subtask.asp";	//&proid=" + proid + "&proname=" + proname;
	document.frmassign.target = "_self";
	document.frmassign.submit();
}

function assignright() {
	document.frmassign.action = "assignright.asp?outside=1";	//&proid=" + proid + "&proname=" + proname;
	document.frmassign.target = "_self";
	document.frmassign.submit();
	}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript">
<form name="frmassign" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			'begin of @@Conntent
			%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
		 <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
        <tr>
          <td class="blue" height="20" align="left">&nbsp;&nbsp;
			<a href="../../Management/project/listofproject.asp?act=BACK" onMouseOver="self.status='Show the list of projects'; return true;" onMouseOut="self.status=''">Project List</a> 
			&nbsp;|&nbsp; 
			<a href="javascript:subtask();" onMouseOver="self.status='Show the list of subtasks'; return true;" onMouseOut="self.status=''">Sub tasks</a>
			&nbsp;|&nbsp; Assignment
			&nbsp;|&nbsp; 
			<a href="javascript:assignright();" onMouseOver="self.status='Right on Project'; return true;" onMouseOut="self.status=''">Right for subtask</a></td>
        </tr>
        <tr valign="middle"> 
          <td class="title" height="50" align="center"> Projects Assignment <span class="blue-normal"><br> <%=proID & " (" & proName & ")"%></span></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#617DC0"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#617DC0"> 
	<%
	'--------------------------------------------------
	' Write the body of HTML page (menu)
	'--------------------------------------------------	
	Response.Write(strLast)
	%>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
							<tr> 
	<%if fgUpdate and Instr(strLast, "<input type='checkbox'")>0 then%>
							  <td class="blue-normal" align="left" height="20" width="69%">&nbsp;&nbsp;*Choose 
							    the checkbox, then click &quot;remove&quot; to remove assignment.</td>
							  <td class="blue" align="right" height="20" width="31%">&nbsp;<a href="javascript:setchecked(1);" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
							    All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''">Clear 
							    All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:remove();" onMouseOver="self.status='Remove assignment'; return true;" onMouseOut="self.status=''"> Remove</a> 
							    &nbsp;</td>
	<%end if%>
							</tr>
						</table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%'end of @@content
  Response.Write(arrTmp(1))
%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			%>	
<input type="hidden" name="txthiddenstrproName" value="<%=proName%>">
<input type="hidden" name="txthiddenstrproID" value="<%=proID%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>