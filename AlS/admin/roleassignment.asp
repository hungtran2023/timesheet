<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../inc/library.asp"-->
<%
'****************************************
' function: outBody
' Description: table of list staff
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize, ByVal whichpage)
	strOut = ""
	rsSrc.MoveFirst
	rsSrc.Move (whichpage-1)*psize
	if not rsSrc.EOF then
		For i = 1 to psize
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' width='29%' class='blue'>" & Showlabel(rsSrc("Fullname")) & "</td>" &_
			         "<td valign='top' width='30%' class='blue-normal'>" & Showlabel(rsSrc("Jobtitle")) & "</td>" &_
			         "<td valign='top' width='30%' class='blue-normal'>" & Showlabel(rsSrc("Department")) & "</td>" &_
			         "<td valign='top' width='11%' class='blue-normal' align='center'>" &_
			         "<input type='checkbox' name='chkstaff' value='" & rsSrc("UserID") & "'>" &_
			         "</td></tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
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
	objDb.recConnect(strConnect)
	strQuery = "SELECT a.UserID, e.Firstname+' '+isnull(e.lastname, '')+' '+isnull(e.Middlename, '') as Fullname, c.JobTitle, d.Department "&_
				"FROM ATC_UserGroup a "&_
				"LEFT JOIN ATC_Employees b ON a.UserID = b.StaffID " &_
				"LEFT JOIN ATC_JobTitle c ON b.JobtitleID = c.JobtitleID " &_
				"LEFT JOIN ATC_Department d ON b.DepartmentID = d.DepartmentID " &_
				"LEFT JOIN ATC_PersonalInfo e ON a.UserID = e.PersonID " &_
				"WHERE a.GroupID=" & DetailofID & " ORDER BY Fullname"
	If objDb.openRec(strQuery) Then	  
	  objDb.recDisConnect
	  if not objDb.noRecord then
		set rsParticipant = objDb.rsElement.Clone
		set session("rsEmpCache") = rsParticipant
		rsParticipant.MoveFirst
		fgret = true
	  else
		fgRet=false
	  end if
	  objDb.CloseRec
	Else
	  gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
	GetDetail = fgRet
End function
'****************************************
' function: task_remove
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function task_remove
Dim fgret
	fgret = false
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		objDb.cnDatabase.BeginTrans	  
		cntU = Request.Form("chkstaff").Count
		for i = 1 to cntU
			uID = int(Request.Form("chkstaff")(i))
			strQuery = "DELETE ATC_UserGroup WHERE UserID = " & uID & " AND GroupID =" & strRoleID
			if not objDb.runActionQuery(strQuery) then
				gMessage = objDb.strMessage
				exit for
			end if
		next
		if gMessage<>"" then 
			objDb.cnDatabase.RollbackTrans
		else
			objDb.cnDatabase.CommitTrans
			gMessage = "Removed successfully."
			fgret = true
		end if
		objDb.dbDisConnect
	else 'error in connection
	  gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	task_remove = fgret
End function
'----------------------------------------------------------------------------------------
	Dim strFunction
	Dim objDb, gMessage, strRolename, strRoleID, PageSize
	session("USERID") = session("Inhouse") 'session("USERID") must have value because selectemployee.asp need it to check session timeout
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
	end if
'------------------------------------	
' Get Full Name
'------------------------------------
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
	strtmp1 = Replace(logoff, "XX", session("strHTTP") & "admin/")
	strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"
'----------------------------------
' Main procedure
'----------------------------------
Call freeAdmininput
Call freeRole
Call freeListRole
gMessage = ""
fgReresh = 0

strRoleID = Request.Form("txtroleid")
strRolename = Request.Form("txthiddenrolename")

strAct = Request.QueryString("act")
if Request.QueryString("addass")<>"" then Call freeRoleAss
if Request.QueryString("outside")<>"" then Call freeRoleAss

select case strAct
case "REMOVE"
	if task_remove() then
		if not isEmpty(session("rsEmpcache")) then
			set rsPar = session("rsEmpcache")
			session("rsEmpcache") = empty
			rsPar.Close
			set rsPar = nothing
		end if
	end if
end select

if session("READYROLEASS")<> True and strRoleID<>"" Then
	if Getdetail(strRoleID) then
		set rsRole = session("rsEmpCache")
		rsRole.MoveFirst
		session("NumPageroleAss") = pageCount(rsRole, PageSize)
		if isEmpty(Session("CurPageroleAss")) then 
			Session("CurPageroleAss") = 1
		else
			if Session("CurPageroleAss")>session("NumPageroleAss") then
				Session("CurPageroleAss") = session("NumPageroleAss")
			elseif Session("CurPageroleAss") = 0 then
				Session("CurPageroleAss") = 1
			end if
		end if
	Else
		Session("CurPageroleAss") = 0
		Session("NumPageroleAss") = 0
	End if
End if

varNavi = Request.QueryString("navi")
if varNavi<>"" then
	tmpi = Session("CurPageRoleAss")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPageRoleAss") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPageRoleAss")
			end if
	End select
	Session("CurPageRoleAss") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPageRoleAss") = CInt(varGo)

strLast=""
if not isEmpty(session("rsEmpCache")) then
	set rsParticipant = session("rsEmpCache")
	strLast = OutBody(rsParticipant, pageSize, Session("CurPageRoleAss"))
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
var objSEWindow;
function fetch() { //v2.0
var taskid = "<%=strRoleID%>";
	window.status = "";
	strFeatures = "top="+(screen.height/2-200)+",left="+(screen.width/2-225)+",width=450,height=405,toolbar=no," 
	            + "menubar=no,location=no,directories=no,resizable=no";
	if((objSEWindow) && (!objSEWindow.closed))
		objSEWindow.focus();	
	else 
		objSEWindow = window.open("../management/project/selectemployee.asp?outside=1&kind=4&taskid=" + taskid, "MyNewWindow", strFeatures);
	window.status = "Opened a new browser window.";
}

function window_onunload() {
	if((objSEWindow)&&(!objSEWindow.closed))
		objSEWindow.close();
}

//function window_onload() {
//var tmp = "<%=gMessage%>";
//	if (tmp != "") {
//		alert(tmp)
//	}
//}
// onLoad="return window_onload();"

function next() {
var curpage = <%=session("CurPageRoleAss")%>
var numpage = <%=session("NumPageRoleAss")%>
	if (curpage < numpage) {
		document.frmrole.action = "roleassignment.asp?navi=NEXT";
		document.frmrole.target = "_self";
		document.frmrole.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageRoleAss")%>
var numpage = <%=session("NumPageRoleAss")%>
	if (curpage > 1) {
		document.frmrole.action = "roleassignment.asp?navi=PREV";
		document.frmrole.target = "_self";
		document.frmrole.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageRoleAss")%>
	var curpage = <%=session("CurPageRoleAss")%>
	var intpage = document.frmrole.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmrole.action = "roleassignment.asp?Go=" + intpage;
		document.frmrole.target = "_self";
		document.frmrole.submit();		
	}
}

function setchecked(val) {
  with (document.frmrole) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkstaff") {
			elements[ii].checked = val;
		}
	}
  }
}

function chkremove() {
  fg = false;
  with (document.frmrole) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chkstaff") && (elements[ii].checked)) {
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
  	document.frmrole.action = "roleassignment.asp?act=REMOVE";
	document.frmrole.target = "_self";
	document.frmrole.submit();
  }
}


function delete_() {
  	document.frmrole.action = "roleassignment.asp?act=DELETE";
	document.frmrole.target = "_self";
	document.frmrole.submit();
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  LANGUAGE="javascript" onUnload="return window_onunload();">
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
        <tr>
          <td class="blue" height="20" align="left">&nbsp;&nbsp;<a href="../admin/listofroles.asp" onMouseOver="self.status='Show the list of roles'; return true;" onMouseOut="self.status=''"> 
            Role List</a> <span class="blue-normal">/<%=strRolename%></span></td>
        </tr>
        <tr valign="middle"> 
          <td class="title" height="50" align="center"> Role Assignment</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
		<tr align="left">
		  <td class="blue" height="20">&nbsp;&nbsp;<a href="javascript:fetch();" onMouseOver="self.status='Add the employees'; return true;" onMouseOut="self.status=''">Add 
		    New</a>&nbsp;</td>
		</tr>
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
<%if strLast>"" then%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
                              <td class="blue" align="center" width="29%">Fullname</td>
                              <td class="blue" align="center" width="30%">Job Title</td>
                              <td class="blue" align="center" width="30%">Department</td>
                              <td class="blue" align="center" width="11%">&nbsp;</td>
                            </tr>
	<%
	'--------------------------------------------------
	' Write the body of HTML page (menu)
	'--------------------------------------------------	
	Response.Write(strLast)
	%>					  </table>
                        </td>
                      </tr>
                    </table>
					<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
						<tr> 
						  <td class="blue-normal" align="left" height="20" width="69%">&nbsp;&nbsp;</td>
						  <td class="blue" align="right" height="20" width="31%">&nbsp;<a href="javascript:setchecked(1);" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
						    All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''">Clear 
						    All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:remove();" onMouseOver="self.status='Remove'; return true;" onMouseOut="self.status=''"> Remove</a> 
						    &nbsp;</td>
						</tr>
					</table>
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
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageRoleAss")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPageRoleAss")%>/<%=session("NumPageRoleAss")%>&nbsp;&nbsp;</td>
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
<input type="hidden" name="txtroleid" value="<%=strRoleID%>">
<input type="hidden" name="txthiddenrolename" value="<%=strRolename%>">
</form>
</body>
</html>