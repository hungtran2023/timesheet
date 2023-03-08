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
			         "<input type='checkbox' name='chkstaff' value='" & rsSrc("StaffID") & "'>" &_
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
' Parameters: DetailofID is ShortlistID need to get details
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************
Function GetDetail(ByVal DetailofID)
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
	strQuery = "SELECT a.staffID, e.Firstname+' '+isnull(e.Middlename, '')+' '+isnull(e.lastname, '') as Fullname, c.JobTitle, d.Department "&_
				"FROM ATC_ShortlistDetails a "&_
				"LEFT JOIN ATC_Employees b ON a.StaffID = b.StaffID " &_
				"LEFT JOIN ATC_JobTitle c ON b.JobtitleID = c.JobtitleID " &_
				"LEFT JOIN ATC_Department d ON b.DepartmentID = d.DepartmentID " &_
				"LEFT JOIN ATC_PersonalInfo e ON a.StaffID = e.PersonID " &_
				"WHERE a.ShortlistID=" & DetailofID & " ORDER BY Fullname"
	If objDb.openRec(strQuery) Then	  
	  objDb.recDisConnect
	  if not objDb.noRecord then
		set rsParticipant = objDb.rsElement.Clone
		set session("rsShortCache") = rsParticipant
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
'**************************************************
' Function: Getlist
' Description: array or empty string 
' Parameters: DetailofID is ShortlistID need to get details
' Return value: 
' Author: 
' Date: 
' Note:
'**************************************************		
Function Getlist()
Dim arrlistShort
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	ret = objDb.dbConnect(strConnect)
	if ret then
		strQuery = "SELECT * FROM ATC_Shortlists WHERE OwnerID = " & session("USERID")
		If objDb.runQuery(strQuery) Then
			if not objDb.noRecord then
				arrlistShort = objDb.rsElement.GetRows
				objDb.CloseRec
			else
				arrlistShort = ""
			end if
		else
		  gMessage = objDb.strMessage
		End if
		objDb.dbDisConnect
	else 'error in connection
		gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	Getlist = arrlistShort
End function
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
function addshort()
Dim fgret
	fgret = false
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	ret = objDb.dbConnect(strConnect)
	if ret then
		strshortlistnamevalue = Request.Form("txtshortlistname")
		strshortlistnamevalue = replace(strshortlistnamevalue, "'", "''")
		strQuery = "INSERT INTO ATC_Shortlists(Shortlist, OwnerID) VALUES('" & strshortlistnamevalue & "', " & session("USERID") & ")"
		If objDb.runActionQuery(strQuery) Then
			strQuery = "Select @@IDENTITY as myid"
			if objDb.runQuery(strQuery)  then
				strShortID = objDb.rsElement("myid")
				gMessage = "Added successfully."
				objDb.CloseRec
				fgret = true
			else
				gMessage = objDb.strMessage
			end if
		else
		  gMessage = objDb.strMessage
		End if
		objDb.dbDisConnect
	else 'error in connection
		gMessage = objDb.strMessage
	end if
	Set objDb = Nothing
	if fgret then addshort = strShortID else addshort = ""
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
Function task_delete
Dim fgret
	fgret = false
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
	  objDb.cnDatabase.BeginTrans
	  strQuery = "DELETE ATC_ShortlistDetails WHERE ShortListID = " & strShortID
	  if objDb.runActionQuery(strQuery) then
		strQuery = "DELETE ATC_Shortlists WHERE ShortListID = " & strShortID
		if not objDb.runActionQuery(strQuery) then gMessage = objDb.strMessage
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
			strQuery = "DELETE ATC_ShortlistDetails WHERE StaffID = " & uID & " AND ShortlistID = " & strShortID
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
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, strshortlistname, strShortID, PageSize
	Dim rsParticipant
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If
	
'-----------------------------------
'Check MANAGER right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = "manager" then
				fgRight=true
				exit for
			end if
		next
	end if
	if fgRight = false then
		Response.Redirect("../welcome.asp")
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
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID")) 'strUserName)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu = "" then strChoseMenu = "C"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../")

'----------------------------------
' Main procedure
'----------------------------------
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
	if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
		Call freeShort
	end if
end if

gMessage = ""
fgReresh = 0

strShortID = Request.Form("lstshort")
strShortlistname = Request.Form("txtshortname")
if fgExecute then
	strAct = Request.QueryString("act")
	if Request.QueryString("addass")<>"" then fgReresh = 1
	if strShortID = "" then fgReresh = 1

	select case strAct
	case "ADDSHORT"
		ret = addshort()
		if ret<>"" then
			strShortID = ret
			set session("lstShort") = nothing
			session("lstShort") = Empty
			fgReresh = 1
		else
			strshortlistnamevalue = Request.Form("txtshortlistname")
			strShortID = ""
		end if
	case "GOPAGE"
		fgReresh = 1
	case "DELETE"
		if task_delete() then
			strShortID = ""
			fgReresh = 1
			set session("lstShort") = nothing
			session("lstShort") = Empty
		end if
	case "REMOVE"
		if task_remove() then
			fgReresh = 1
		end if
	end select

	'reset Detail if user click on another Shortlist
	if fgReresh = 1 then
	  if not isEmpty(session("rsShortCache")) then
		set rsParticipant = session("rsShortCache")
		rsParticipant.Close
		set rsParticipant = nothing
		session("rsShortCache") = empty
	  end if
	end if
end if 'fgExecute

'get list if not have yet
if isEmpty(session("lstShort")) then
	ret = Getlist()
	if isArray(ret) then session("lstShort") = ret
	Session("CurPageshort") = 0
	session("NumPageshort") = 0
end if

'get detail if no have yet and listshort<>empty
strlistShort = ""
if not isEmpty(session("lstShort")) then
	'generate list of shortlist
	arrShort = session("lstShort")
	strlistShort = "<select name='lstshort' size='1' class='blue-normal'>"
	if strShortID="" then 
		strShortID = arrShort(0, 0)
		strShortlistname = arrShort(1, 0)
	End if
	For ii = 0 to ubound(arrShort, 2)
		if Cint(strShortID) = arrShort(0, ii) then strSel = " selected " else strSel = ""
		strlistShort = strlistShort & "<option value='" &  arrShort(0, ii) & "'" & strSel & ">" &  Showlabel(arrShort(1, ii)) & "</option>"
	Next
	strlistShort = strlistShort & "</select>"
	'------------------------------get user
	if isEmpty(session("rsShortCache")) then
		ret = GetDetail(strShortID)
		if ret then
		  session("NumPageshort") = pageCount(rsParticipant, PageSize)
		  Session("CurPageshort") = 1
		else
		  Session("CurPageshort") = 0
		  session("NumPageshort") = 0
		end if
	end if
End if

if fgExecute then
	varNavi = Request.QueryString("navi")
	if varNavi<>"" then
		tmpi = Session("CurPageshort")
		select case varNavi
			case "PREV"
				if tmpi > 1 then
					tmpi = tmpi - 1
				else
					tmpi = 1
				end if
			case "NEXT"
				if tmpi < Session("NumPageshort") then
					tmpi = tmpi + 1
				else
					tmpi = Session("NumPageshort")
				end if
		End select
		Session("CurPageshort") = tmpi
	end if

	varGo = Request.QueryString("Go")
	if varGo <> "" then Session("CurPageshort") = CInt(varGo)
end if

strLast=""
if not isEmpty(session("rsShortCache")) then
	set rsParticipant = session("rsShortCache")
	strLast = OutBody(rsParticipant, pageSize, Session("CurPageshort"))
end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_menu.htm")


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

<link rel="stylesheet" href="../timesheet.css">

<script language="javascript" src="../library/library.js"></script>
<script>
var objSEWindow;
function fetch() { //v2.0
var taskid = "<%=strshortID%>";
	window.status = "";
	strFeatures = "top="+(screen.height/2-200)+",left="+(screen.width/2-225)+",width=450,height=405,toolbar=no," 
	            + "menubar=no,location=no,directories=no,resizable=no";
	if((objSEWindow) && (!objSEWindow.closed))
		objSEWindow.focus();	
	else 
		objSEWindow = window.open("../management/project/selectemployee.asp?outside=1&kind=3&taskid=" + taskid, "MyNewWindow", strFeatures);
	window.status = "Opened a new browser window.";
}

function window_onunload() {
	if((objSEWindow) && (!objSEWindow.closed))
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
var curpage = <%=session("CurPageshort")%>
var numpage = <%=session("NumPageshort")%>
	if (curpage < numpage) {
		document.frmshort.action = "shortlist.asp?navi=NEXT";
		document.frmshort.target = "_self";
		document.frmshort.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageshort")%>
var numpage = <%=session("NumPageshort")%>
	if (curpage > 1) {
		document.frmshort.action = "shortlist.asp?navi=PREV";
		document.frmshort.target = "_self";
		document.frmshort.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageshort")%>
	var curpage = <%=session("CurPageshort")%>
	var intpage = document.frmshort.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmshort.action = "shortlist.asp?Go=" + intpage;
		document.frmshort.target = "_self";
		document.frmshort.submit();		
	}
}

function setchecked(val) {
  with (document.frmshort) {
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
  with (document.frmshort) {
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
  	document.frmshort.action = "shortlist.asp?act=REMOVE";
	document.frmshort.target = "_self";
	document.frmshort.submit();
  }
}

function gopage() {
	var tmp = document.frmshort.lstshort.options[document.frmshort.lstshort.selectedIndex].text;
	document.frmshort.txtshortname.value = tmp;
  	document.frmshort.action = "shortlist.asp?act=GOPAGE";
	document.frmshort.target = "_self";
	document.frmshort.submit();
}

function addshort() {
	var tmp  = document.frmshort.txtshortlistname.value;
	document.frmshort.txtshortlistname.value = alltrim(tmp);
	if (document.frmshort.txtshortlistname.value=="") { alert("Please enter a shortlist name.") }
	else {
		document.frmshort.action = "shortlist.asp?act=ADDSHORT";
		document.frmshort.target = "_self";
		document.frmshort.submit();
	}
}

function delete_() {
  	document.frmshort.action = "shortlist.asp?act=DELETE";
	document.frmshort.target = "_self";
	document.frmshort.submit();
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  LANGUAGE="javascript" onUnload="return window_onunload();">
<form name="frmshort" method="post">
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
        <tr bgcolor=<%if gMessage="" then%>"#FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		 <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
        <tr valign="middle"> 
          <td class="title" height="50" align="center"> Shortlist</td>
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
                      <td>
					  <table width="100%" border="0" cellpadding="0" cellspacing="0">
					    <tr align="center"> 
					      <td class="blue-normal" height="30" align="right" width="108"> 
					        Shortlist Name&nbsp;&nbsp; </td>
					      <td class="blue" height="30" align="left" width="160"> 
					        <input type="text" name="txtshortlistname" maxlength="50" class="blue-normal" size="15" style=" width:160" value="<%=strshortlistnamevalue%>">
					      </td>
					      <td class="blue-normal" height="30" align="left" width="348"> 
					        <table width="60" border="0" cellspacing="5" cellpadding="0" height="20">
					          <tr> 
					            <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
					            <a href="javascript:addshort();" class="b" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add</a></td>
					          </tr>
					        </table>
					      </td>
					    </tr>
<%if strlistShort<>"" then
%>
					    <tr align="center"> 
					      <td class="blue-normal" align="right" width="108">Available 
					        Shortlist&nbsp;&nbsp; </td>
					      <td class="blue-normal" align="left" width="160"> 
		<%Response.Write strlistshort%> <a href="javascript:gopage();" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a>
					      </td>
					      <td align="left" width="348"> 
					        <table width="60" border="0" cellspacing="5" cellpadding="0" height="20">
					          <tr> 
					            <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
					            <a href="javascript:delete_();" class="b" onMouseOver="self.status='Delete'; return true;" onMouseOut="self.status=''">Delete</a></td>
					          </tr>
					        </table>
					      </td>
					    </tr>
					    <tr align="left"> 
					      <td class="blue" colspan="3" height="30">&nbsp;&nbsp;Employees 
					        in Shortlist: <%=strShortlistname%></td>
					    </tr>
					    <tr align="left">
					      <td class="blue" colspan="3" height="20">&nbsp;&nbsp;<a href="javascript:fetch();" onMouseOver="self.status='Add the employees'; return true;" onMouseOut="self.status=''">Add 
					        New</a>&nbsp;</td>
					    </tr>
<%end if%>
					</table>
<%if strLast>"" then%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
                              <td class="blue" align="center" width="29%">Full</td>
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
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageShort")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPageShort")%>/<%=session("NumPageshort")%>&nbsp;&nbsp;</td>
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
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			%>
<input type="hidden" name="txtshortid" value="<%=strShortID%>">
<input type="hidden" name="txtshortname" value="<%=strShortlistname%>">
</form>
</body>
</html>