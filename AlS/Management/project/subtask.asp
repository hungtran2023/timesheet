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
Sub BuildTaskList(byval rsTaskClone,byval blnRightParent,byval SubTaskID,byval intLevel, byref strReturn,byref strListTask)
	
	dim blnRight

	rsTaskClone.Filter = "ParentID = " & SubTaskID	
	If rsTaskClone.recordCount>0 then
		Do while not rsTaskClone.EOF
			blnRight=blnRightParent OR (rsTaskClone("OwnerID")=Session("UserID")) OR (rsTaskClone("StaffID")=Session("UserID")) OR fgDelegate
			
			strReturn=strReturn & AppendTree(rsTaskClone("SubtaskName"),rsTaskClone("fgBillable"),intlevel,blnRight and (intLevel<>0),rsTaskClone("SubTaskID"))
			strListTask = strListTask & AppendList (rsTaskClone("SubtaskName"),intLevel , blnRight , rsTaskClone("SubTaskID"), rsTaskClone("ChainID"))
			
			call BuildTaskList(rsTaskClone.Clone,blnRight,rsTaskClone("SubTaskID"),intLevel + 1, strReturn,strListTask)
			rsTaskClone.MoveNext
		loop
	end if
end Sub
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
Sub GetTaskInformation(byval TaskID,byref TaskParent,byref TaskName)
	Dim strConn,objDbase,strQuery
	strConn = Application("g_strConnect")
	Set objDbase = New clsDatabase
	objDbase.recConnect(strConn)
	strQuery="SELECT * FROM ATC_Tasks WHERE subtaskID=" & TaskID
	
	If objDbase.openRec(strQuery) Then
		objDbase.recDisConnect
		if not objDbase.noRecord then
			TaskParent=objDbase.rsElement("TaskID")
			TaskName=objDbase.rsElement("SubTaskName")
			fgBillable=objDbase.rsElement("fgBillable")
		else
			gMessage = "No data available."
		end if
		objDbase.CloseRec
	Else
		gMessage = objDbase.strMessage
	End if
	
	set objDbase=nothing
End Sub
'****************************************
' function: SaveTaskInformation
' Description: Update
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
sub UpdateTaskInformation(byref taskID)
	
	dim varTask,objDb,strConnect
				
	varTask = Request.Form("txtname")
	varTask = replace(varTask, "'", "''")
	varTask = replace(varTask, chr(34), "''")
	
		
	fgBillable=Request.Form("lstTaskType")
	if fgBillable="" then fgBillable=0
	
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
		
	if objDb.dbConnect(strConnect) then
		'testing for this task have assignment or not		
		if gMessage = "" then
			strQuery = "UPDATE ATC_Tasks SET SubtaskName='" & varTask & "', fgBillable=" & fgBillable & " WHERE subTaskID=" & taskID

			if not objDb.runActionQuery(strQuery) then 
				gMessage = objDb.strMessage
			else
				gMessage="Updated successfully."
				taskID=-1
			end if
		end if
		objDb.dbdisConnect
	else
		gMessage = objDb.strMessage
	end if
	
	set objDb = nothing
	
end sub

'****************************************
' function: InsertTaskInformation
' Description: Update
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
sub InsertTaskInformation
	
	dim varTask,objDb,strConnect,varTaskParent
	dim strChainID
			
	varTaskParent = Request.Form("lsttask")
	varTask = Request.Form("txtname")
	varTask = replace(varTask, "'", "''")
	varTask = replace(varTask, chr(34), "''")
	
	fgBillable=Request.Form("lstTaskType")
	if fgBillable="" then fgBillable=0
	
	strChainID = Mid(varTaskParent, InStr(varTaskParent,"@") + 1,len(varTaskParent))
	varTaskParent = Mid(varTaskParent, 1, InStr(varTaskParent,"@") - 1)
	strChainID = strChainID & varTaskParent & ","
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	if objDb.dbConnect(strConnect) then
		'testing for this task have assignment or not
		strQuery="SELECT * FROM ATC_Tasks WHERE SubtaskID IN " & _
					"(SELECT SubtaskID FROM ATC_Assignments WHERE subtaskID=" & varTaskParent & ")"
		
		if objDb.runQuery(strQuery) then
			if not objDb.noRecord then
				gMessage = "Task '" & objDb.rsElement("SubTaskName") & "'  was assigned."
			else
				strQuery = "INSERT INTO ATC_Tasks(ProjectID, SubtaskName, TaskID, ChainID, OwnerID,fgBillable) " &_
						"VALUES('" & proID & "', '" & varTask & "', " & varTaskParent & ", '" & strChainID & "', " & session("USERID") & "," & fgBillable & ")"		
				
				if not objDb.runActionQuery(strQuery) then gMessage = objDb.strMessage
				taskID=-1
			end if
		else
			gMessage = objDb.strMessage
		end if
		objDb.dbdisConnect
	else
		gMessage = objDb.strMessage
	end if		
	set objDb = nothing
end sub
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
Function AppendTree (ByVal strsName, byval blnBillable,ByVal intLevel, ByVal blnShow, ByVal intValue)
Dim strTmp, i, strColor
dim strLink,strBillable
	strTmp = ""
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "<IMG alt='' border='0' height='18' src='../../images/t_dot.gif' width='36'>"
		Next
		strTmp = strTmp & "<IMG alt='' border='0' src='../../images/dot1.gif'>"
		strTmp = strTmp & "<IMG alt='' border='0' height='10' width='12' src='../../images/nosign.gif'>"
	End If

	strBillable="&nbsp;"
	if cint(blnBillable)=1 then 
		strBillable="<img src='../../images/yes.gif' alt='Billable'>"
	elseif cint(blnBillable)=2 then
		strBillable="<img src='../../images/notyet.gif' alt='Non Recovering'>"
	end if
	
	LineOnPage = LineOnPage + 1
	strColor = "#FFF2F2"
	if blnShow = true and intLevel>0 then
		strLink="<a href='javascript:edit(" & intValue & ")' class='c' onMouseOver='self.status=&quot;Edit subtask&quot; ; return true;' onMouseout='self.status=&quot;&quot;'>" & Showlabel(strsName) & "</a>"
		AppendTree = "<tr bgcolor='" & strColor & "'><td valign='top' class='blue'>" & strTmp & strLink & "</td>" &_
					"<td valign='top' width='5%' align='center' class='blue'>" & strBillable &_
					"</td><td valign='top' width='5%' align='center' class='blue'><input type='checkbox' name='chkrem' value='" & intValue & "@" & showlabel(strsName) & "'>" &_
		            "</td></tr>" & chr(13)
	else
		AppendTree = "<tr bgcolor='" & strColor & "'><td valign='top' class='black'>" & strTmp & Showlabel(strsName) & "</td>" &_
					"<td valign='top' width='5%' align='center' class='black'>&nbsp; " & strBillable &_
					"</td><td valign='top' width='5%' align='center' class='black'>&nbsp;" &_
		            "</td></tr>" & chr(13)
	end if
End Function

'****************************************
' function: task_remove
' Description: removing assignment
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_remove
	countU = Request.Form("chkrem").Count
	if countU>0 then
	  Set objDb = New clsDatabase
	  strConnect = Application("g_strConnect")
	  ret = objDb.dbConnect(strConnect)
	  if ret then
	    strDonot = ""
	    for ii = 1 to countU
	  		strTaskname = Request.Form("chkrem")(ii)
	  		strTask = Mid(strTaskname, 1, Instr(strTaskname, "@") - 1)
	  		strTaskname = Mid(strTaskname, Instr(strTaskname, "@") + 1, len(strTaskname))			
		
			set myCmd = Server.CreateObject("ADODB.Command")
			set myCmd.ActiveConnection = objDb.cnDatabase
			myCmd.CommandType = adCmdStoredProc
			myCmd.CommandText = "Checkanddeltask"
			set myParam = myCmd.CreateParameter("result", adTinyInt, adParamReturnValue)
			myCmd.Parameters.Append myParam
			set myParam = myCmd.CreateParameter("taskID", adInteger, adParamInput)

			myCmd.Parameters.Append myParam
			myCmd("taskID") = cdbl(strTask)
			myCmd.Execute , , adExecuteNoRecords
			if myCmd("result") = 0 then
				strDonot = strDonot & " " & strTaskname & ","
			end if
			set myCmd = nothing
	    next
	    if strDonot<>"" then
	      strDonot = Mid(strDonot, 1, len(strDonot)- 1)
	      gMessage = "Can not remove '" & strDonot & "'."
	      fgRefresh = "0"
	    else
	  	  gMessage = "Removed successfully."
	  	  fgRefresh = "1"
	    end if
	    objDb.dbDisConnect
	  else
	    gMessage =  objDb.strMessage
	  end if
	  Set objDb = Nothing
	end if
End Sub

'****************************************
' function: task_remove
' Description: removing assignment
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function Add3tasks(byval proID)
    
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "Add3Subtasks"
	
		Set myParam = myCmd.CreateParameter("ProjectID", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam		
		
		myCmd("ProjectID")	= proID
		
		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
					
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	Add3tasks=strError
	
end Function
'--------------------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb,rsTask
	Dim objUser, gMessage
	Dim ProName, proID 'value that is got from query string
	Dim strOut, fgRefresh, fgUpdate,strLast,strList,fgBillable
	Dim TaskID,TaskParent,TaskName
	
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
	tmp = "listofproject.asp"
	
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

'---------------------------	
' Get Full Name
'---------------------------
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
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")
'------------------------------
' main procedure
'------------------------------
	if Request.QueryString("fgMenu") <> "" then
		fgExecute = false
	else
		fgExecute = true
	end if
	
	gMessage = ""
	'get all Sub Task and Parent Task of one project
	proID = Request.Form("txthiddenstrproID")
	proName = Request.Form("txthiddenstrproName")
	if fgExecute then strAct = Request.QueryString("act")
	TaskID=-1
	TaskParent=-1
	TaskName=""
	fgBillable=1
	select case strAct
	case "REMOVE"
		Call task_remove
	case "EDIT"
		TaskID=Request.QueryString("TaskID")
		Call GetTaskInformation(TaskID,TaskParent,TaskName)
	case "SAVE"
		TaskID=Request.QueryString("TaskID")
	
		if cdbl(TaskID)<>-1 then 

			call UpdateTaskInformation(taskID)
		else
			call InsertTaskInformation()
		end if
	case "ADD3"
	    Call Add3tasks(proID)
	end select

	if proID <>"" then 'draw sub task
		strConnect = Application("g_strConnect")
		Set objDb = New clsDatabase
		objDb.recConnect(strConnect)

		strQuery="SELECT a.SubTaskID,ProjectID,SubTaskName,fgBillable,ISNULL(TaskID,0)as parentID,ChainID,OwnerID , b.StaffID " & _
				"FROM ATC_Tasks a LEFT JOIN ATC_RightonTasks b ON (a.SubtaskID=b.SubtaskID AND b.StaffID=" & session("USERID") & ") " & _
				"WHERE projectID='" & proID & "' ORDER BY TaskID,a.SubtaskID"		
								
'Response.Write strQuery		

		If objDb.openRec(strQuery) Then
			objDb.recDisConnect
			if not objDb.noRecord then
				set rsTask = objDb.rsElement.Clone
				rsTask.MoveFirst
				strLast = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & chr(13) & _
						  "  <tr><td bgcolor='#DDDDDD'>" &_
						  "<table width='100%' border='0' cellspacing='1' cellpadding='5'>"		
					
				Call BuildTaskList(rsTask.Clone,false,0,0,strLast,strList)					
				strLast = strLast & "</table></td></tr></table>"
				rsTask.Filter= "ParentID = 0"
				
				intOwner=rsTask("OwnerID")
			else
				gMessage = "No data available."
			end if
			objDb.CloseRec
		Else
			gMessage = objDb.strMessage
		End if
		Set objDb = Nothing
	end if

fgUpdate= fgUpdate and (Instr(strLast, "<input type='checkbox'")>0 or intOwner=Session("UserID") or fgDelegate)

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
<script LANGUAGE="JavaScript">

function CheckData() {
	var strType="<%=strAct%>"
	if (document.proinfo.txtname.value == "") {
		alert("Please enter value for this field.");
		document.proinfo.txtname.focus();
		return false;
	}
	if (strType==""){
		var tmp = document.proinfo.lsttask.selectedIndex;
		if (document.proinfo.lsttask.options[tmp].value == "") {
			alert("You have no permission to make changes to this task.");
			document.proinfo.lsttask.focus();
			return false;
		}
	}
	return true;
}

function savedata() {
	var TaskID=<%=TaskID%>
	if (CheckData()==true) {
		document.proinfo.action = "subtask.asp?act=SAVE&taskID=" + TaskID;	//&proID=" + proid;
		document.proinfo.target = "_self";
		document.proinfo.submit();
	}
}

function add3Tasks() {
		
	document.proinfo.action = "subtask.asp?act=ADD3";
	document.proinfo.target = "_self";
	document.proinfo.submit();
	
}


function edit(TaskID) {
	document.proinfo.action = "subtask.asp?act=EDIT&taskID=" + TaskID;	//&proID=" + proid;
	document.proinfo.target = "_self";
	document.proinfo.submit();
}

function add() {
	document.proinfo.action = "subtask.asp";
	document.proinfo.target = "_self";
	document.proinfo.submit();
}

function setchecked(val) {
  with (document.proinfo) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkrem") {
			elements[ii].checked = val;
		}
	}
  }
}

function chkremove() {
  fg = false;
  with (document.proinfo) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chkrem") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 if (fg == false) alert("No task selected.")
 return(fg)
}

function remove() {
	if(chkremove()==true) {
		document.proinfo.action = "subtask.asp?act=REMOVE";	
		document.proinfo.target = "_self";
		document.proinfo.submit();
	}
}

function assign() {
	document.proinfo.action = "assignment.asp?outside=1";	//&proid=" + proid + "&proname=" + proname;
	document.proinfo.target = "_self";
	document.proinfo.submit();
}

function assignright() {
	document.proinfo.action = "assignright.asp?outside=1";	//&proid=" + proid + "&proname=" + proname;
	document.proinfo.target = "_self";
	document.proinfo.submit();
}


</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript">
<form name="proinfo" method="post">
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
          <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
        </tr>
        <tr> 
          <td class="blue" align="left" width="90%">&nbsp;&nbsp;
		<a href="listofproject.asp?act=BACK" onMouseOver="self.status='Show the list of projects'; return true;" onMouseOut="self.status=''">Project List</a> 
		&nbsp;|&nbsp; Sub tasks&nbsp;
		|&nbsp; <a href="javascript:assign();" onMouseOver="self.status='Project assignments'; return true;" onMouseOut="self.status=''">Assignment</a> &nbsp;
		|&nbsp; <a href="javascript:assignright();" onMouseOver="self.status='Right on Project'; return true;" onMouseOut="self.status=''">Right for subtask</a></td>
          <td class="blue" width="10%">&nbsp;</td>
        </tr>
        <tr align="center"> 
          <td class="title" height="70" align="center" colspan="2">Sub-tasks <span class="blue-normal"><br> <%=proID & " (" & proName & ")"%></span></td>
        </tr>
      </table>
    </td>
  </tr>
  
<%'Response.Write fgUpdate
if fgUpdate then%>    
  <tr>
	<td valign="top">
	<table width="60%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#003399">
		<tr> 
		  <td >
			<table width="100%" border="0" align="center" cellpadding="20" cellspacing="0" >
			  <tr> 
				        <td bgcolor="#C0CAE6" ><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td>
							  	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
                                  <tr> 
										
                                    <td width="30%" align="right" class="blue-normal">  Name&nbsp; </td>
                                    <td width="70%" class="text-blue01" bgcolor="C0CAE6"><input type="text" name="txtname" maxlength="50" class="blue-normal" style='HEIGHT: 22px; WIDTH:100%' value="<%=TaskName%>"> 
                                    </td>
									</tr>
									
									<%if strAct<>"EDIT" Then%>
									<tr> 
    									
                                    <td width="30%" align="right" class="blue-normal"> Sub-task of&nbsp; </td>		                                
                                    <td width="70%" bgcolor="C0CAE6" class="text-blue01" ><select name='lsttask' class='blue-normal' style='HEIGHT: 22px; WIDTH: 100%'>
                                        <%=strList%>
                                      </select> </td>
									</tr>
									
									
									<%end if%>
									
									<tr> 
    									
                                    <td width="30%" align="right" class="blue-normal"> Type&nbsp; </td>		                                
                                    <td width="70%" bgcolor="C0CAE6" class="blue-normal" >
										<select name='lstTaskType' class='blue-normal' style='HEIGHT: 22px; WIDTH: 100%'>
											<option value='1' <%if cint(fgBillable)=1 then %>selected<%end if%>>Billable</option>
									<%if fgDelegate then%>											
											<option value='0' <%if cint(fgBillable)=0 then %>selected<%end if%>>None Billable</option>
											<option value='2' <%if cint(fgBillable)=2 then %>selected<%end if%>>Risked Billable</option>
									<%end if%>
										</select>
										
									</td>
									</tr>
								</table></td>
                            </tr>
                            <tr>
                              <td><table width="120" height="30" border="0" align="right" cellpadding="0" cellspacing="5" name="aa">
                                  <tr> 
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" align="center" class="blue"> 
                                      <a href="javascript:savedata();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''"><%if strAct="EDIT" Then%>Save<%else%>Add<%end if%></a> 
                                    </td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table> </td>
			  </tr>
			</table>
		  </td>
		</tr>
	</table>
			</td>
  </tr>
<%end if%>  
  <tr> 
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="blue" align="left" height="20">&nbsp;&nbsp;
	<%if fgUpdate then%><a href="javascript:add();" onMouseOver="self.status='Add a sub-task'; return true;" onMouseOut="self.status=''">Add Sub-Task</a>
	&nbsp;|&nbsp; <a href="javascript:add3Tasks();" onMouseOver="self.status=''; return true;" onMouseOut="self.status=''">Add 4 Sub-Tasks</a>
	<%end if%>&nbsp;</td>
              </tr>
            </table>
	<%Response.Write strLast
	if fgUpdate then%>          
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <tr>
	
                 <td class="blue-normal" align="left" height="20" width="69%">&nbsp;&nbsp;*Choose 
                   the checkbox, then click &quot;remove&quot; to remove sub-task.</td>
                 <td class="blue" align="right" height="20" width="31%">&nbsp;<a href="javascript:setchecked(1);" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
                   All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''">Clear All</a>&nbsp;&nbsp;&nbsp; 
                   <a href="javascript: remove();" onMouseOver="self.status='Remove sub-task'; return true;" onMouseOut="self.status=''"> Remove</a> &nbsp;</td>

               </tr>
             </table>
	<%end if%>             
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
			<%
			Response.Write(arrTmp(1))
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			%>
<input type="hidden" name="txthiddenstrproID" value="<%=proID%>">
<input type="hidden" name="txthiddenstrproName" value="<%=proName%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>