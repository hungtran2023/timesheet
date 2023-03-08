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
			blnRight=blnRightParent OR (rsTaskClone("OwnerID")=Session("UserID")) OR fgDelegate
			strButton="<a href='javascript:assignment(" & rsTaskClone("SubTaskID") & ");'>Assign</a>"

			strReturn=strReturn & AppendTree(rsTaskClone("SubtaskName"),intlevel,blnRight and (intLevel<>0))
			
			if blnRight then 
				strReturn=Replace(strReturn,"©",strButton)
			else
				strReturn=Replace(strReturn,"©","")
			end if
			
			strReturn=strReturn & AppendUser(rsParticipant.Clone,rsTaskClone("SubtaskID"),intLevel,blnRight)
			
			blnRight =blnRight OR (rsTaskClone("StaffID")=Session("UserID"))

			call BuildTaskList(rsTaskClone.Clone,blnRight,rsTaskClone("SubTaskID"),intLevel + 1, strReturn)
			rsTaskClone.MoveNext
		loop
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
		strIndent = "<IMG alt='' vspace='0' border='0' height='10' src='../../images/t_dot.gif' width='" & (intLevel*36+36+20) & "'>"
	Else
		strIndent = "<IMG alt='' vspace='0' border='0' height='10' src='../../images/t_dot.gif' width='5'>"
	End If	
	
	Rs.Filter = "SubtaskID = " & TaskID
	'Response.Write "<br>SubtaskID = " & TaskID & " - Count:" & Rs.Recordcount
	Do Until Rs.EOF
		strColor = "#E7EBF5"
		if blnShow then
		  strTmp = strTmp & "<tr bgcolor='" & strColor & "'><td height='18' class='blue-normal'>" & strIndent & showlabel(rs("FullName")) &_
		  			"</td><td height='18' class='blue-normal' width='200'>" & Showlabel(rs("JobTitle")) & "</td>" &_
		  			"<td align='center' class='blue-normal' width='24' ><input type='checkbox' name='chkjunior' value='" & rs("StaffID") & "#" & rs("SubtaskID") & "' " & strCHK & "></td></tr>" & chr(13)
		else
		  strTmp = strTmp & "<tr bgcolor='" & strColor & "'><td height='18' class='blue-normal'>" & strIndent & showlabel(rs("FullName")) &_
		  			"</td><td height='18' class='blue-normal' width='200' colspan='2'>" & Showlabel(rs("JobTitle")) & "</td>" &_
		  			"</tr>" & chr(13)
		end if
		Rs.MoveNext
	Loop
	Rs.Filter=""
	AppendUser = strTmp
End Function
'****************************************
' function: append3Dot
' Description: draw ...
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function Append3Dot (ByVal intLevel)
Dim strTmp, strIndent, strColor
	strTmp = ""
	If intLevel > 0 Then		
		strIndent = "<IMG alt='' border='0' height='18' src='../../images/t_dot.gif' width='" & (intLevel*36+36+20) & "'>"
	Else
		strIndent = "<IMG alt='' border='0' height='18' src='../../images/t_dot.gif' width='5'>"
	End If	
	LineOnPage = LineOnPage + 1
	strColor = "#E7EBF5"
	strTmp = "<tr bgcolor='" & strColor & "'><td colspan='3'>" & strIndent & "..." & "</td></tr>" & chr(13)
	Append3Dot = strTmp
End Function
'****************************************
' function: appendTree
' Description: draw tree of subtask
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
		AppendTree = "<tr bgcolor='" & strColor & "'><td colspan='2' class='blue'>" & strTmp & Showlabel(strsName) & "</td><td align='center'>©</td></tr>" & chr(13)
	else
		AppendTree = "<tr bgcolor='" & strColor & "'><td colspan='2' class='black'>" & strTmp & Showlabel(strsName) & "</td><td align='center'>©</td></tr>" & chr(13)
	end if
End Function

'****************************

	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	Dim rsJunior, objUser, fgUpdate
	Dim arrRs(4), varBookMark, varBookPro, LineOnPage, varSubID, flagShow, gMessage, PageSize
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.Form("txtpreviouspage")
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		fgDelegate=false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
			end if
			
			if getRight(0, ii) = "Project Delegator" then fgDelegate=true
		next
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
	
'--------------------------------------
' Get Full Name
'--------------------------------------
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
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'-------------------------------------
gMessage = ""
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
end if

'get all Sub Task and Parent Task of one project
proID =	Request.Form("txthiddenstrproID")	'Request.QueryString("proid")
proName = Request.Form("txthiddenstrproName")	'Request.QueryString("proname")

if fgExecute and fgUpdate then
	fgOutside = Request.QueryString("outside")
	strAct = Request.QueryString("act")
	if strAct = "REMOVE" then
		countU = Request.Form("chkjunior").Count
		if countU>0 then
			Set objDb = New clsDatabase
			strConnect = Application("g_strConnect")
			
			if objDb.dbConnect(strConnect) then
				objDb.cnDatabase.BeginTrans
				For i = 1 to countU
					varBook = Request.Form("chkjunior")(i)
								
					varBook=replace(varBook,"#"," AND SubtaskID=")
					strQuery = "DELETE FROM ATC_RightOnTasks WHERE StaffID = " & varBook
					ret = objDb.runActionQuery(strQuery)
					if not ret then 
						gMessage = objDb.strMessage
						exit for
					end if
				Next
						  
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
End if	'fgExecute
'-------------------------------------------------------------
'get infomation of participant
'-------------------------------------------------------------
if proID <>"" then 'draw sub task
	strQuery="SELECT a.SubTaskID,ProjectID,SubTaskName,ISNULL(TaskID,0)as parentID,ChainID,OwnerID , ISNULL(b.StaffID ,0) as StaffID " & _
			"FROM ATC_Tasks a LEFT JOIN ATC_RightonTasks b ON (a.SubtaskID=b.SubtaskID AND b.StaffID=" & session("USERID") & ") " & _
			"WHERE projectID='" & proID & "' AND Right(SubtaskName,3)<>'-TP' ORDER BY TaskID,a.SubtaskID"									
	call GetData(strQuery,rsTask)
	'Response.Write strQuery

	strQuery="SELECT FullName,JobTitle,a.* " & _
			"FROM ATC_RightOnTasks a " & _
				"INNER JOIN HR_Employee b ON a.StaffID=b.PersonID    " &_
			"WHERE  subTaskID IN (SELECT subTaskID FROM ATC_Tasks WHERE projectID='" & proID & "') ORDER BY FullName"

	call GetData(strQuery,rsParticipant)
	'Response.Write strQuery
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
var fgWhich;

function assignment(taskid) { //v2.0
	
	document.frmassign.action = "selectemployee_ass.asp?outside=1&kind=2&taskid=" + taskid
	document.frmassign.target = "_self"
	document.frmassign.submit();
}

function ViewAssignment() { //v2.0
	
	//document.frmassign.action = "assignment.asp?outside=1";
	//document.frmassign.target = "_self"
	//document.frmassign.submit();
}


function setchecked(val) {
  with (document.frmassign) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkjunior") {
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
		if ((elements[ii].name == "chkjunior") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 if (fg == false) alert("No person selected.")
 return(fg)
}

function remove() {
  if (chkremove()==true) {
  	document.frmassign.action = "assignright.asp?act=REMOVE";
	document.frmassign.target = "_self";
	document.frmassign.submit();
  }
}

function subtask() {
	document.frmassign.action = "subtask.asp";	//&proid=" + proid + "&proname=" + proname;
	document.frmassign.target = "_self";
	document.frmassign.submit();
}
function assign() {
	document.frmassign.action = "assignment.asp?outside=1";	//&proid=" + proid + "&proname=" + proname;
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
			%>
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			'Response.Write(arrPageTemplate(1))
			Response.Write(arrTmp(0))
			'begin of @@Conntent
			%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
		 <td class="red" colspan="4" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
        <tr> 
          <td class="blue" height="20" align="left">&nbsp;&nbsp;
			<a href="../../Management/project/listofproject.asp?act=BACK" onMouseOver="self.status='Show the list of projects'; return true;" onMouseOut="self.status=''">Project List</a> 
			&nbsp;|&nbsp; 
			<a href="javascript:subtask();" onMouseOver="self.status='Show the list of subtasks'; return true;" onMouseOut="self.status=''">Sub tasks</a>
			&nbsp;|&nbsp; <a href="javascript:assign();" onMouseOver="self.status='Project assignments'; return true;" onMouseOut="self.status=''">Assignment</a>
			&nbsp;|&nbsp; Right for subtask</td>
        </tr>
        <tr valign="middle"> 
          <td class="title" height="50" align="center"> Right on Projects<span class="blue-normal"><br> <%=proID & " (" & proName & ")"%></span></td>
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
<input type="hidden" name="txttemp" value="<%=fgUpdate%>">
</form>
</body>
</html>