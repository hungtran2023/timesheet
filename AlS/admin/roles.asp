<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../inc/library.asp"-->
<%
'****************************************
' function: outBody
' Description: table of list function, have 7 column (description, parentname, updateable, fgExec, FunctionID, selected, fgEnable)
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef arrSrc, ByVal psize, ByVal whichpage)
  strOut = ""
  i = (whichpage-1)*psize
  if i >= 0 then
	lastU = Ubound(arrSrc, 2)
	strParentRes = ""
	cnt = 0
	if strMode<>"EDIT" then 
		strRestrict = "onClick='return false;'"
	else
		strRestrict = ""
	end if
	Do Until i > lastU
		cnt = cnt + 1
'		if strMode<>"EDIT" then 
'			strRestrictUpdate = ""
'		else
'			strRestrictUpdate = "onClick='checkright(" & cnt & ");'"
'		end if
		if strParentRes <> arrSrc(1, i) then
			strParent = Showlabel(arrSrc(1, i))
			strParentRes = arrSrc(1, i)				
		else
			strParent = "_"
		end if
		strUpdate = ""
		if arrSrc(3, i) = 1 then
			if arrSrc(2, i) = 1 then strUpdate = "checked"
		end if
		strSelected = ""
		if arrSrc(5, i) <> -1 then strSelected = "checked"
		if i mod 2 = 0 then
			strColor = "#E7EBF5"
		else
			strColor = "#FFF2F2"
		end if
'Response.Write 	arrSrc(0, i) & "- Update:" & arrSrc(2, i) & "- Enable:" & arrSrc(6, i)  & "<br>"			
		strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
		        "<td valign='top' width='23%' class='blue'>" & strParent & "</td>" &_
		        "<td valign='top' width='54%' class='blue'>" & Showlabel(arrSrc(0, i)) & "</td>" &_
		        "<td valign='top' width='11%' class='blue-normal' align='center'>"
		if arrSrc(6, i) = 1 then
			strOut = strOut &_
				"<input type='checkbox' name='chkupdate' value='" & i & "' " & strUpdate & " " & strRestrict & "></td>"
		else
			strOut = strOut &_
				"&nbsp;</td>"
		end if
		strOut = strOut &_
				"<td valign='top' width='11%' class='blue-normal' align='center'>" &_
		         "<input type='checkbox' name='chkfunc' value='" & i & "' " & strSelected & " " & strRestrict & ">" &_
		         "</td></tr>" & chr(13)
		i = i + 1
		if cnt = pSize then exit do
	Loop
  End if
  Outbody = strOut
end function
'**************************************************
' Function: GetDetail
' Description: 
' Parameters: DetailofID is RoleID need to get details
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Function GetDetail(ByVal DetailofID)
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	if objDb.dbConnect(strConnect) then
		if DetailofID <> "" then
			strQuery = "SELECT C.Description, C.ParentName, ISNULL(D.Updateable, 0) Updateable, E.fg fgExec, " &_
					"C.FunctionID, isnull(D.FunctionID, -1) Selected, C.fgUpdateable fgEnable " &_
					"FROM (SELECT a.Description, b.Description ParentName, a.FunctionID, a.GroupID, a.fgUpdateable " &_
					"FROM ATC_Functions a LEFT JOIN ATC_Functions b ON a.GroupID = b.FunctionID " &_
					"WHERE a.GroupID is not null and (a.Form like '%/%'  or a.Form like '%%'  or a.fgAttribute = 1)) C " &_
					"LEFT JOIN (select * from ATC_Permissions where GroupID = " & DetailofID & ") D " &_
					"ON C.FunctionID = D.FunctionID " &_
					"LEFT JOIN (select 0 as fg, FunctionID from ATC_Functions Where Form not like '%/%' and Form not " &_
					"like '%.htm') E ON C.FunctionID = E.FunctionID ORDER BY C.groupID"
		else
			strQuery = "SELECT C.Description, C.ParentName, 0 Updateable, E.fg fgExec, C.FunctionID, -1 Selected, " &_
					"C.fgUpdateable fgEnable FROM (SELECT a.Description, b.Description ParentName, a.FunctionID, a.GroupID, " &_
					"a.fgUpdateable FROM ATC_Functions a LEFT JOIN ATC_Functions b ON a.GroupID = b.FunctionID " &_
					"WHERE a.GroupID is not null and (a.Form like '%/%'  or a.fgAttribute =1)) C " &_
					"LEFT JOIN (select 0 as fg, FunctionID from ATC_Functions Where Form not like '%/%' and Form not " &_
					"like '%.htm') E ON C.FunctionID = E.FunctionID ORDER BY C.groupID"
		end if

		If objDb.runQuery(strQuery) Then
		  if not objDb.noRecord then
			arrFunc = objDb.rsElement.GetRows
			for i = 0 to Ubound(arrFunc, 2)
				'Updateable
				if arrFunc(2, i) = true then
					arrFunc(2, i) = 1
				else
					arrFunc(2, i) = 0
				end if
				'Enable
				if arrFunc(6, i) = true then
					arrFunc(6, i) = 1
				else
					arrFunc(6, i) = 0
				end if
				if IsNull(arrFunc(3, i)) then arrFunc(3, i)=1
'Response.Write 	arrFunc(0, i) & "- Update:" & arrFunc(2, i) & "- Enable:" & arrFunc(6, i)  & "<br>"
			next
			session("arrFuncCache") = arrFunc
			set arrFunc = nothing
			fgRet = true
			objDb.CloseRec
		  else
			fgRet=false
		  end if
		Else
		  gMessage = objDb.strMessage
		End if
		objDb.dbDisconnect
	else	'error in connection
		gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	GetDetail = fgRet
End function
'**************************************************
' Function: GetInfo
' Description: 
' Parameters: DetailofID is RoleID need to get info
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Function GetInfo(ByVal DetailofID)
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	if objDb.dbConnect(strConnect) then
		strQuery = "SELECT * FROM ATC_Group WHERE GroupID = " & DetailofID
		If objDb.runQuery(strQuery) Then
		  if not objDb.noRecord then
			strRoleName = objDb.rsElement("GroupName")
			strComment = objDb.rsElement("Comment")
			fgRet = true
			objDb.CloseRec
		  else
			fgRet=false
		  end if
		Else
		  gMessage = objDb.strMessage
		End if
		objDb.dbDisconnect
	else	'error in connection
		gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	GetInfo = fgRet
End function
'****************************************
' function: task_save
' Description:
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function task_save()
Dim fgret
	fgret = false
	gMessage = ""
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	ret = objDb.dbConnect(strConnect)
	if ret then
		objDb.cndatabase.BeginTrans
		strRolenamei = Request.Form("txtrolename")
		strRolenamei = replace(strRolenamei, "'", "''")
		strCommenti = Request.Form("txtcomment")
		strCommenti = replace(strCommenti, "'", "''")
		strQuery = "INSERT INTO ATC_Group(Groupname, Comment) VALUES('" & strRolenamei & "', '" & strCommenti & "')"
		If objDb.runActionQuery(strQuery) Then
			strQuery = "Select @@IDENTITY as myid"
			if objDb.runQuery(strQuery) then
				strRoleID = objDb.rsElement("myid")
				objDb.CloseRec
				'insert into ATC_permission
				arrTmp = session("arrFuncCache")
				for i = 0 to Ubound(arrTmp, 2)
					if arrTmp(5, i)<>-1 then 'selected
						if arrTmp(3, i) = 1 then
							strup = arrTmp(2, i)
						else
							strup = "0"
						end if
						strQuery = "INSERT INTO ATC_Permissions(GroupID, FunctionID, Updateable) VALUES(" &_
									strRoleID & "," & arrTmp(4, i) & "," & strup & ")"
						if not objDb.runActionQuery(strQuery) then
							gMessage = objDb.strMessage
							exit for
						end if
					end if
				next
				'end insert
			else
				gMessage = objDb.strMessage
			end if
		else
		  gMessage = objDb.strMessage
		End if
		if gMessage = "" then
			objDb.cndatabase.CommitTrans
			gMessage = "Added successfully."
			fgret = true
		else
			objDb.cndatabase.RollbackTrans
		end if
		objDb.dbDisConnect
	else 'error in connection
		gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	if fgret then task_save = strRoleID else task_save = ""
end function
'****************************************
' function: task_update
' Description:
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function task_update(ByVal roleid)
Dim fgret
	gMessage = ""
	fgret = false
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	ret = objDb.dbConnect(strConnect)
	if ret then
		objDb.cndatabase.BeginTrans
		strRolename = Request.Form("txtrolename")
		strRolename = replace(strRolename, "'", "''")
		strComment = Request.Form("txtcomment")
		strComment = replace(strComment, "'", "''")
		strQuery = "UPDATE ATC_Group SET Groupname = '" & strRolename & "', Comment='" & strComment & "' WHERE GroupID = " & roleid
	
		If objDb.runActionQuery(strQuery) Then
			strQuery = "DELETE ATC_Permissions WHERE GroupID = " & roleid
			if objDb.runQuery(strQuery) then
				'insert into ATC_permission
				arrTmp = session("arrFuncCache")
				for i = 0 to Ubound(arrTmp, 2)
					if arrTmp(5, i)<>-1 then 'selected
						if arrTmp(3, i) = 1 then
							strup = arrTmp(2, i)
						else
							strup = "0"
						end if
						strQuery = "INSERT INTO ATC_Permissions(GroupID, FunctionID, Updateable) VALUES(" &_
									roleid & "," & arrTmp(4, i) & "," & strup & ")"
'Response.Write	strQuery & "<br>"
						if not objDb.runActionQuery(strQuery) then
							gMessage = objDb.strMessage
							exit for
						end if
					end if
				next
				set arrTmp = nothing
				'end insert
			else
				gMessage = objDb.strMessage
			end if
		else
		  gMessage = objDb.strMessage
		End if
		if gMessage = "" then
			objDb.cndatabase.CommitTrans
			fgret = true
			gMessage = "Saved successfully."
		else
			objDb.cndatabase.RollbackTrans
		end if
		objDb.dbDisConnect
	else 'error in connection
		gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	task_update = fgret
end function
'****************************************
' function: task_delete
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function task_delete(roleid)
Dim fgret
	fgret = false
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
	  objDb.cnDatabase.BeginTrans
	  strQuery = "DELETE ATC_Permissions WHERE GroupID = " & roleid
	  if objDb.runActionQuery(strQuery) then
		strQuery = "DELETE ATC_UserGroup WHERE GroupID = " & roleid
		if objDb.runActionQuery(strQuery) then
			strQuery = "DELETE ATC_Group WHERE GroupID = " & roleid
			if not objDb.runActionQuery(strQuery) then gMessage = objDb.strMessage
		else
			gMessage = objDb.strMessage
		end if
	  else 'error in query 1
		gMessage = objDb.strMessage
	  end if
	  if gMessage<>"" then 
	  	objDb.cnDatabase.RollbackTrans
	  else
	  	objDb.cnDatabase.CommitTrans
	  	gMessage = "Deleted successfully."
	  	fgret = true
	  end if
	  objDb.dbDisConnect
	else 'error in connection
	  gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	task_delete = fgret
End function
'****************************************
' function: writearray
' Description: update arrfunccache for check box
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub writearray(ByVal whichpage, ByVal psize)
	'Reset info of page
	i = (whichpage-1)*psize
	arrTmp = session("arrFuncCache")
	lastU = Ubound(arrTmp, 2)
	cnt = 0
	Do Until i > lastU
		cnt = cnt + 1
		arrTmp(2, i) = 0
		arrTmp(5, i) = -1
		i = i + 1 
		if cnt = pSize then exit do
	Loop
	
	countU = Request.Form("chkfunc").Count
	if countU>0 then
		For i = 1 to countU
			varIdx = int(Request.Form("chkfunc")(i))
			arrTmp(5, varIdx) = 1
		Next
	end if
	
	countU = Request.Form("chkupdate").Count
	if countU>0 then
		For i = 1 to countU
			varIdx = int(Request.Form("chkupdate")(i))
			arrTmp(2, varIdx) = 1
			arrTmp(5, varIdx) = 1
		Next
	end if
	session("arrFuncCache") = arrTmp
	set arrTmp = nothing
End sub
'----------------------------------------------------------------------------------------
	Dim strFunction
	Dim objDb, gMessage, strRolename, strRoleID, strComment, PageSize
	Dim strMode, fgRightDelete 'enable or not delete button
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("Inhouse")) = False Then
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
		PageSize=100
	end if
	PageSize = PageSize - 3
'------------------------------------	
' Get Full Name
'------------------------------------
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
	strtmp1 = Replace(logoff, "XX", session("strHTTP") & "admin/")
	strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"
			
'------------------------------------------------------------------
' Main procedure
'------------------------------------------------------------------
gMessage = ""
Call freeAdmininput
Call freeListRole

strRoleID = Request.Form("txtroleid")
strRoleName = Request.Form("txtrolename")
strComment = Request.Form("txtcomment")

strAct = Request.QueryString("act")

strMode = Request.Form("txtmode")
fgRightDelete = Request.Form("txtrightdelete")

if Request.QueryString("outside") <> "" then
	if strRoleID <> "" then ' view or edit
		strAct = "PREPARE"
	else 'adding
		strAct = "ADD"
	end if
end if

select case strAct
case "SAVE"
	Call writearray(Session("CurPagefunc"), pageSize)
	
	if strRoleID <> "" then
		ret = task_update(strRoleID)
	else
		ret = task_save()
		if ret<>"" then 'successful
			strRoleID = ret
		end if
	end if
	if ret then 'successful
		strMode = "VIEW"
		fgRightDelete = 1
	end if
case "DELETE"
	if task_delete(strRoleID) then
		Response.Redirect "listofroles.asp?act=REFRESH"
	end if
case "EDIT"
	strMode = "EDIT"
	fgRightDelete = 1
case "ADD"
	fgRefresh = 1
	strMode = "EDIT"
	fgRightDelete = 0 'because this is a new record that is not saved
	strRoleID = ""
	strRoleName = ""
	strComment = ""
case "PREPARE"
	fgRefresh = 1
	ret = GetInfo(strRoleID)
	strMode = "VIEW"
	fgRightDelete = 1
end select

if fgRefresh = 1 then
	ret = GetDetail(strRoleID)
	if ret then
	  Session("CurPagefunc") = 1
	  session("NumPagefunc") = pageCount(session("arrFuncCache"), PageSize)
	else
	  Session("CurPagefunc") = 0
	  session("NumPagefunc") = 0
	end if
end if

oldCurpage = Session("CurPagefunc")
varNavi = Request.QueryString("navi")
if varNavi<>"" then
	tmpi = Session("CurPagefunc")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPagefunc") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPagefunc")
			end if
	End select
	Session("CurPagefunc") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPagefunc") = CInt(varGo)

if oldCurpage > 0 and oldCurpage <> Session("CurPagefunc") then
	Call writearray(oldCurpage, pageSize)
end if
strLast=""
if not isEmpty(session("arrFuncCache")) then
	arrFunc = session("arrFuncCache")
	strLast = OutBody(arrFunc, pageSize, Session("CurPagefunc"))
	set arrFunc = nothing
end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_admin.htm")
curpage = 7
If arrPageTemplate(1)<>"" then
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(0) = Replace(arrTmp(0),"@@function", strfunction)
	for i = 1 to NumOfAdminMenu
		if i <> curpage then
			arrTmp(0) = Replace(arrTmp(0),"@@markin"&cstr(i)&"@@", "")
			arrTmp(0) = Replace(arrTmp(0),"@@markout"&cstr(i)&"@@", "")
		else
			arrTmp(0) = Replace(arrTmp(0),"@@markin"&cstr(i)&"@@", "<font color='#CD0000'>")
			arrTmp(0) = Replace(arrTmp(0),"@@markout"&cstr(i)&"@@", "</font>")
		end if
	next
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css">

<script language="javascript" src="../library/library.js"></script>
<script>
function next() {
var curpage = <%=session("CurPagefunc")%>
var numpage = <%=session("NumPagefunc")%>
	if (curpage < numpage) {
		document.frmrole.action = "roles.asp?navi=NEXT";
		document.frmrole.target = "_self";
		document.frmrole.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagefunc")%>
var numpage = <%=session("NumPagefunc")%>
	if (curpage > 1) {
		document.frmrole.action = "roles.asp?navi=PREV";
		document.frmrole.target = "_self";
		document.frmrole.submit();
	}
}

function go() {
	var numpage = <%=session("NumPagefunc")%>
	var curpage = <%=session("CurPagefunc")%>
	var intpage = document.frmrole.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmrole.action = "roles.asp?Go=" + intpage;
		document.frmrole.target = "_self";
		document.frmrole.submit();		
	}
}

function setchecked(val) {
  with (document.frmrole) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkfunc") {
			elements[ii].checked = val;
		}
	}
  }
}

function chkselected() {
  fg = false;
  with (document.frmrole) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chkfunc") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 return(fg)
}

function checkdata() {
	var tmp = document.frmrole.txtrolename.value;
	document.frmrole.txtrolename.value = alltrim(tmp);
	if(document.frmrole.txtrolename.value==""){
		document.frmrole.txtrolename.focus();
		alert("Please enter a value.")
		return false;
	}
	if(chkselected()==false){
		alert("No Function selected.")
		return false;
	}
	return true;
}

function _save() {
  if (checkdata()==true) {
  	document.frmrole.action = "roles.asp?act=SAVE";
	document.frmrole.target = "_self";
	document.frmrole.submit();
  }
}

function _edit() {
  	document.frmrole.action = "roles.asp?act=EDIT";
	document.frmrole.target = "_self";
	document.frmrole.submit();
}

function _add() {
	document.frmrole.action = "roles.asp?act=ADD";
	document.frmrole.target = "_self";
	document.frmrole.submit();
}

function _delete() {
  	document.frmrole.action = "roles.asp?act=DELETE";
	document.frmrole.target = "_self";
	document.frmrole.submit();
}

function CheckMode(field){
var varMode="<%=strMode%>";
    if (varMode!="EDIT"){
        field.blur();
	}
}

function checkright(idx) {
	if(document.frmrole.chkfunc[idx-1].checked==false) {
		document.frmrole.chkupdate[idx-1].checked = false;
	}
}

function _assign() {
var varname = "<%=strrolename%>";
	window.document.frmrole.txthiddenrolename.value = varname;
  	document.frmrole.action = "roleassignment.asp";
	document.frmrole.target = "_self";
	document.frmrole.submit();
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmrole" method="post">
			<%
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
        <tr bgcolor=<%if gMessage="" then%>"#FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		 <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
        <tr align="center"> 
         <td class="blue" align="left" width="23%"> &nbsp;&nbsp;
		 <a href="listofroles.asp" onMouseOver="self.status='Return the previous page'; return true;" onMouseOut="self.status=''">Role 
           List</a></td>
        </tr>		
        <tr valign="middle"> 
          <td class="title" height="50" align="center">Role Informations</td>
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
                <td class="blue-normal" width="31%">&nbsp;</td>
                <td class="blue-normal" width="14%">Role Name</td>
                <td colspan="2"> 
                  <input type="text" name="txtrolename" class="blue-normal" size="15" style=" width:160" maxlength="30" 
					value="<%=Showvalue(strRoleName)%>" onFocus="CheckMode(this)">
                </td>
              </tr>
              <tr> 
                <td class="blue-normal" width="31%">&nbsp;</td>
                <td class="blue-normal" width="14%">Comment</td>
                <td colspan="2"> 
                  <input type="text" name="txtcomment" class="blue-normal" size="15" style=" width:160" maxlength="100" 
					value="<%=Showvalue(strComment)%>" onFocus="CheckMode(this)">
                </td>
              </tr>
              <tr> 
                <td colspan="4" height="40"> 
                  <table width="300" border="0" cellspacing="5" cellpadding="0" align="center" height="20">
                    <tr> 
                      <td align="center" class="blue" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20"> 
                          <a href="javascript:_add();" class="b" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add</a></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center"> 
<%if strMode = "EDIT" then%>Edit<%else%><a href="javascript:_edit();" class="b" onMouseOver="self.status='Edit'; return true;" onMouseOut="self.status=''">Edit</a><%end if%></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if strMode = "EDIT" then%><a href="javascript:_save();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Save</a><%else%>Save<%end if%>
						</td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if fgrightDelete = 1 then%><a href="javascript:_delete();" class="b" onMouseOver="self.status='Delete'; return true;" onMouseOut="self.status=''">Delete</a><%else%>Delete<%end if%></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if fgrightDelete = 1 then%><a href="javascript:_assign();" class="b" onMouseOver="self.status='Assignment'; return true;" onMouseOut="self.status=''">Assign</a><%else%>Assign<%end if%></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
                    
<%if strLast>"" then%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#617DC0"> 
                  <table width="100%" border="0" cellspacing="1" cellpadding="5">
                    <tr bgcolor="8CA0D1"> 
                      <td class="blue" align="center" width="23%">Function group</td>
                      <td class="blue" align="center" width="54%">Function</td>
                      <td class="blue" align="center" width="11%">Updateable</td>
                      <td class="blue" align="center" width="11%">View</td>
                    </tr>
	<%
	'--------------------------------------------------
	' Write the body of HTML page (menu)
	'--------------------------------------------------	
	Response.Write(strLast)
	%>		  </table>
           </td>
         </tr>
        </table>
<!--		<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
			<tr> 
			  <td class="blue-normal" align="left" height="20" width="69%">&nbsp;&nbsp;</td>
			  <td class="blue" align="right" height="20" width="31%">&nbsp;<a href="javascript:setchecked(1);" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
			    All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''">Clear 
			    All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:remove();" onMouseOver="self.status='Remove'; return true;" onMouseOut="self.status=''"> Remove</a> 
			    &nbsp;</td>
			</tr>
		</table>
-->
<%end if%>
        </td>
      </tr>
    </table>
    </td>
  </tr>
<%if strLast<>"" then%>
  <tr> 
    <td> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
		  <tr> 
		    <td align="right" bgcolor="#E7EBF5"> 
		      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
		        <tr class="black-normal"> 
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPagefunc")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPagefunc")%>/<%=session("NumPagefunc")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Go to previous page'; return true;" onMouseOut="self.status=''">Previous</a> /
		          <a href="javascript:next();" onMouseOver="self.status='Go to next page'; return true;" onMouseOut="self.status=''"> Next</a>&nbsp;&nbsp;&nbsp;</td>
		        </tr>
		      </table>
		    </td>
		  </tr>
		</table>
    </td>
  </tr>
<%end if%>
</table>
<%'end of @@content
  Response.Write(arrTmp(1))
%>
<input type="hidden" value="<%=strMode%>" name="txtmode">
<input type="hidden" name="txtrightdelete" value="<%=fgRightDelete%>">
<input type="hidden" name="txtroleid" value="<%=strRoleID%>">
<input type="hidden" name="txthiddenrolename">
</form>
</body>
</html>