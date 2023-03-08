<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' function: OutBody
' Description:
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize, Byval whichpage)
Dim strOut
	strOut = ""
	arrTmp = session("arrASS")

	topofpage = (whichpage-1)*psize
	if not rsSrc.EOF then

		cnt = 0
		For i = 1 to psize
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
'Response.Write "works"	& UBound(arrTmp,1)
			strCHK = ""
			if arrTmp(0, topofpage + i - 1) = 1 then
				strCHK = "checked"
			end if
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
					"<td valign='top' width='55%' class='blue'>&nbsp;" & Showlabel(rsSrc("Fullname")) & "</td>" & chr(13) &_
                    "<td valign='top' width='35%' class='blue-normal'>&nbsp;" & Showlabel(rsSrc("JobTitle")) & "</td>" & chr(13) &_
                    "<td valign='top' width='10%' class='blue-normal' align='center'>"

			strOut = strOut & "<input type='checkbox' name='chkass" & CStr(i)& "' value='" & rsSrc.BookMark & "'" & " " & strCHK & "></td>" & chr(13)
			strOut = strOut & "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	

	Outbody = strOut
end function

'--------------------------------------
Dim gMessage, PageSize
Dim arrlstFrom(2),arrlongmon
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check VIEWALL right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = "view all" then
				fgRight=true
				exit for
			end if
		next
		set getRight = nothing
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
	
'----------------------------------	
' Starting
'----------------------------------	

'---------------------------------
' Get ruleID
'---------------------------------
intRuleID=Request.Form("txtRuleID")
strURLBack="DetailOfExpiredRule.asp"
'-------------------------------
' get list of Employees
'-------------------------------

fgOutside = Request.QueryString("outside")

'In case go this page from Detail of Rule
if fgOutside = "1" then
	session("READYUSER") = false
	if not isEmpty(session("rsUser")) and not isNull(session("rsUser")) then
		set rsUser = session("rsUser")
		'rsUser.Close
		set rsUser = nothing
		session("rsUser") = Empty
	end if
end if

'If list employee is not available
if session("READYUSER") <> true then
'----------
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	ret = objDb.dbConnect(strConnect)
	if ret then
		set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = adUseClient     ' Set the Cursor Location to Client
'---------------------------------
'Call store procedure
'---------------------------------
		set myCmd = Server.CreateObject("ADODB.Command")
		set myCmd.ActiveConnection = objDb.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "sp_getListEmp"

		set myParama = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParama
		set myParamb = myCmd.CreateParameter("level",adTinyInt,adParamInput)
		myCmd.Parameters.Append myParamb
		set myParamc = myCmd.CreateParameter("strquery", adVarChar,adParamInput, 5000)
		myCmd.Parameters.Append myParamc
		set myParamd = myCmd.CreateParameter("fgCheck", adTinyInt,adParamInput)
		myCmd.Parameters.Append myParamd
				
		myCmd("StaffID") = session("USERID")
		myCmd("level") = 0
		
		strQuery = "select a.StaffID, b.FirstName + ' ' + isnull(b.MiddleName, '') + ' ' + b.LastName as FullName, a.DirectLeaderID, " &_
				"b.Gender, b.NationalityID, b.CountryID, a.DepartmentID, a.fgIndirect, isNull(c.jobTitle, '') jobTitle, c.JobTitleID, isNull(d.Department, '') Department, 0 as fgSel " &_
				"from ATC_Employees a LEFT JOIN ATC_PersonalInfo b " &_
				"ON a.StaffID = b.PersonID LEFT JOIN ATC_JobTitle c ON a.JobTitleID = c.JobTitleID " &_
				"LEFT JOIN ATC_Department d ON a.DepartmentID = d.DepartmentID WHERE b.fgDelete=0 " & _
				"AND a.StaffID not in (Select y.StaffID From ATC_EmployeeExpiredRule y INNER JOIN " & _
										"(SELECT StaffID, MAX(ApplyYear) as lastdate FROM ATC_EmployeeExpiredRule GROUP BY StaffID) w " & _
										"ON y.StaffID=w.StaffID AND ApplyYear=lastdate " & _
										"WHERE RuleYearlyID =" & intRuleID & " )"
		if fgRight then 'View all
		  myCmd("fgCheck") = 0
		else
		  strQuery = strQuery & " AND a.StaffID "
		  myCmd("fgCheck") = 1
		End if		
		
		myCmd("strquery") = strQuery	
				
		On Error Resume Next	
		rsUser.Open myCmd,,adOpenStatic,adLockBatchOptimistic
		If Err.number > 0 then
			gMessage = Err.Description
		End If
		Err.Clear
'----------
		if not rsUser.EOF then
			session("READYUSER") = true
			session("NumPagesee") = pageCount(rsUser, pageSize)
			Session("CurPagesee") = 1
			rsUser.MoveFirst
			set session("rsUser") = rsUser
			arrTmp = rsUser.GetRows (,,"fgSel")
			session("arrASS") = arrTmp
			rsUser.MoveFirst
			session("arrSortsee") = array(true, true)
		else
			Session("CurPagesee") = 0
			Session("NumPagesee") = 0
		end if
		session("fgShowsee") =  0 ' show all
		set myCmd = nothing
	else
		gMessage = objDb.strMessage
	end if
	set objDb = nothing
else
	set rsUser = session("rsUser")
end if



if session("READYUSER") then
	If fgOutside="" then
		'update array assign
	  If session("NumPagesee") > 0 then
		'seek to current page
		topofpage = (session("CurPagesee")-1)*PageSize
		arrTmp = session("arrASS")
		numitem = Ubound(arrTmp, 2)
		cnt = 0
		For i = topofpage to topofpage + PageSize - 1
			if i <= numitem then
				cnt = cnt + 1
				if Request.Form("chkass" & CStr(cnt)) <> "" then
					tmp = CInt(Request.Form("chkass" & CStr(cnt)))
					arrTmp(0, i) = 1
					
				else
					arrTmp(0, i) = 0
				end if
			else
				exit for
			end if
		Next
		session("arrASS") = arrTmp
		set arrTmp = nothing
	  end if
	End if

	strAct = Request.QueryString("act")
	if strURLBack="" then strURLBack = Request.Form("txtURLBack")


	
	if strAct="SAVE" then
		
		dateApply= Request.Form("lstYearF")
	
	  	Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		gMessage = ""
		if ret then
			objDb.cnDatabase.BeginTrans
			rsUser.MoveFirst
			arrTmp = session("arrASS")
			i = 0  
			Do Until rsUser.EOF
				if arrTmp(0, i) <> 0 then
'						  
					strQuery = "INSERT INTO ATC_EmployeeExpiredRule (StaffID,ApplyYear,RuleYearlyID) VALUES (" & _
								 rsUser("StaffID") & "," & dateApply  & "," & intRuleID & ")"						 
			 
				  ret = objDb.runActionQuery(strQuery)	  	 
				  if not ret then 
					gMessage = objDb.strMessage
					Exit Do
				  end if
				end if
				i = i + 1
				rsUser.MoveNext
			Loop
		 
			set arrTmp = nothing
			if gMessage<>"" then 
				objDb.cnDatabase.RollbackTrans
			else
				objDb.cnDatabase.CommitTrans
				'gMessage = "Assigned successfully."	  	
			end if
			objDb.dbdisConnect
		else
			gMessage = objDb.strMessage
		end if
	
		set objDb = nothing
		
		if gMessage = "" then
			'--free self
			if not isEmpty(session("READYUSER")) then session("READYUSER") = Empty
			if not isEmpty(session("rsUser")) then
				set rsUser = session("rsUser")
				rsUser.Close
				set rsUser = nothing
				session("rsUser") = empty
			end if
			if not isEmpty(session("arrSortsee")) then session("arrSortsee") = empty
			if not isEmpty(session("arrASS")) then session("arrASS") = empty
			session("CurPagesee") = 0
			session("NumPagesee") = 0
			session("fgShowsee") = 2	'clear
		end if
		
	else 'not SAVE
		varSort = Request.QueryString("sorttype")

		if varSort <> "" then
			strtypesort = ""
			if not isEmpty(session("arrSortsee")) then
				arrTmp = session("arrSortsee")

				if arrTmp(varSort-1) then
					strtypesort = "ASC"
				else
					strtypesort = "DESC"
				end if
				arrTmp(varSort-1) = not arrTmp(varSort-1)
				session("arrSortsee") = arrTmp
				set arrTmp = nothing
			end if	
			select case varSort
				case 1 
					rsUser.Sort = "FullName " & strtypesort
					session("arrASS") = rsUser.GetRows(,,"fgSel")
				case 2 
					rsUser.Sort = "JobTitle " & strtypesort
					session("arrASS") = rsUser.GetRows(,,"fgSel")
			end select
		end if

		varNavi = Request.QueryString("navi")

		if varNavi <> "" then
			tmpi = session("CurPagesee")
			select case varNavi
				case "PREV"
					if tmpi > 1 then
						tmpi = tmpi - 1
					else
						tmpi = 1
					end if
				case "NEXT"
					if tmpi < Session("NumPagesee") then
						tmpi = tmpi + 1
					else
						tmpi = Session("NumPagesee")
					end if
			End select
			session("CurPagesee") = tmpi
		end if

		varGo = Request.QueryString("Go")
		if varGo <> "" then Session("CurPagesee") = CInt(varGo)

		varShowAll = Request.QueryString("showall")
		if varShowAll <> "" then
			rsUser.Filter = ""
			'reset fgsel
			rsUser.MoveFirst
			session("arrASS") = rsUser.GetRows(,,"fgSel")
			session("fgShowsee") = 0
			rsUser.MoveFirst
			session("NumPagesee") = pageCount(rsUser, pageSize)
			Session("CurPagesee") = 1
		end if

		varSearch = trim(Request.QueryString("search"))
		if varSearch<>""  then
			varSearch = replace(varSearch, "%", "")
			varSearch = replace(varSearch, "#", "")
			criteria = trim(varSearch)
			if criteria <> "" then
				if Instr(criteria, "'")>0 then
					criteria = "#" & criteria & "#"
				else
					criteria = "'%" & Replace(criteria, "'", "''") & "%'"
				end if
				rsUser.Filter = "Fullname Like " & criteria
			else
				rsUser.MoveLast
				rsUser.MoveNext
			end if
			If rsUser.EOF then ' no result
				gMessage = "No results found."
				session("fgShowsee") = 0
				rsUser.Filter = ""
				rsUser.MoveFirst
				session("arrASS") = rsUser.GetRows(,,"fgSel")
				rsUser.MoveFirst
				Session("CurPagesee") = 1
				session("NumPagesee") = pageCount(rsUser, pageSize)
			else
				'reset fgsel
				rsUser.MoveFirst
				session("arrASS") = rsUser.GetRows(,,"fgSel")
				session("fgShowsee") = 1 ' show the result of filter
				rsUser.MoveFirst
				session("NumPagesee") = pageCount(rsUser, pageSize)
				Session("CurPagesee") = 1
				gMessage = ""
			end if
		end if
		
		varFilter = Request.QueryString("filter") '0 or 1
		if varFilter<>"" then
			if isEmpty(session("filteremp")) and not isNull(session("filteremp")) then		
				session("fgShowsee") = 0
				rsUser.MoveFirst
				Session("CurPagesee") = 1
				session("NumPagesee") = pageCount(rsUser, pageSize)
			else
				tmp = session("Filteremp")
				rsUser.Filter = tmp
				If rsUser.EOF then ' no result
					gMessage = "No results found."
					rsUser.Filter = ""
					session("fgShowsee") = 0
					rsUser.MoveFirst
					session("arrASS") = rsUser.GetRows(,,"fgSel")
					rsUser.MoveFirst
					Session("CurPagesee") = 1
					session("NumPagesee") = pageCount(rsUser, pageSize)
				else
					'reset fgsel
					rsUser.MoveFirst
					session("arrASS") = rsUser.GetRows(,,"fgSel")
					session("fgShowsee") = 1 ' show the result of filter
					rsUser.MoveFirst
					session("NumPagesee") = pageCount(rsUser, pageSize)
					Session("CurPagesee") = 1
					gMessage = ""
				end if
			end if
		end if
	End if

	strLast = ""
	curpage=1
	If session("NumPagesee") > 0 then
		rsUser.MoveFirst
		rsUser.Move (session("CurPagesee")-1)*pageSize

		curpage = session("CurPagesee")
		strLast = Outbody(rsUser, pageSize, curpage)
	End if
End if
	
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
	if strChoseMenu = "" then strChoseMenu = "AE"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

	arrlstFrom(0) = selectmonth("lstmonthF",month(Date()) , -1)
	arrlstFrom(1) = selectday("lstdayF", day(date()), -1)
	arrlstFrom(2) = selectyear("lstyearf", year(date()), 1999, year(date())+2, 0)
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
<!--
var objEMFIWindow;

function filter() { //v2.0
  window.status = "";
  strFeatures = "top="+(screen.height/2-170)+",left="+(screen.width/2-126)+",width=252,height=240,toolbar=no," 
              + "menubar=no,location=no,directories=no";
  if ((objEMFIWindow) && (!objEMFIWindow.closed)) {
	objEMFIWindow.focus();
	
  } else {
	objEMFIWindow = window.open("empfilter.asp", "MyNewWindow1", strFeatures);
  }
  window.status = "Opened a new browser window.";
  
}
function window_onunload() {
	if((objEMFIWindow)&&(!objEMFIWindow.closed))
		objEMFIWindow.close();
}


function next() {
var curpage = <%=session("CurPagesee")%>
var numpage = <%=session("NumPagesee")%>
	if (curpage < numpage) {
		document.selectemployee.action = "selectemployee_assExpiredRule.asp?navi=NEXT";
		document.selectemployee.target = "_self";
		document.selectemployee.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagesee")%>
var numpage = <%=session("NumPagesee")%>
	if (curpage > 1) {
		document.selectemployee.action = "selectemployee_assExpiredRule.asp?navi=PREV";
		document.selectemployee.target = "_self";
		document.selectemployee.submit();
	}
}

function go() {
	var numpage = <%=session("NumPagesee")%>
	var curpage = <%=session("CurPagesee")%>
	var intpage = document.selectemployee.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.selectemployee.action = "selectemployee_assExpiredRule.asp?Go=" + intpage;
		document.selectemployee.target = "_self";
		document.selectemployee.submit();	
	}
	else
		alert("Enter another number please.")
}

function sort(type) {
	document.selectemployee.action = "selectemployee_assExpiredRule.asp?sorttype=" + type; //1: fullname, 2: jobtitle
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
}

function search() {
	var tmp = document.selectemployee.txtsearch.value
	if (tmp != "") {
		document.selectemployee.action = "selectemployee_assExpiredRule.asp?search=" + tmp;
		document.selectemployee.target = "_self";
		document.selectemployee.submit();
	}
}

function showall() {
var tmp = <%=session("fgShowsee")%>
  if (tmp==1) {
	document.selectemployee.action = "selectemployee_assExpiredRule.asp?showall=" + "1";
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
  }
}

function setchecked(val) {
  with (document.selectemployee) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].type == "checkbox") {
			elements[ii].checked = val;
		}
	}
  }
}

function checkass () {
  selection = false;
  with (document.selectemployee) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].type == "checkbox") && (elements[ii].checked==true)) {
			selection = true;
			break;
		}
	}
  }
  return(selection)
}

function assignment() {



		document.selectemployee.action = "selectemployee_assExpiredRule.asp?act=SAVE";
		document.selectemployee.target = "_self";
		document.selectemployee.submit();

}

function BackPrevious(strURL) {

	document.selectemployee.action = strURL;
	document.selectemployee.target = "_self";
	document.selectemployee.submit();
	
}

-->
</script>
</head>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onunload="return window_onunload();">
<form name="selectemployee" method="post">
<% If gMessage<>"" OR strAct<>"SAVE" then %>
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
      <td height="90"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		    <td class="red" colspan="4" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		  </tr>
			<tr> 
			  <td class="blue" width="11%" valign="middle" >&nbsp;&nbsp;&nbsp;
					<a href="javascript:BackPrevious('<%=strURLBack%>');" onMouseOver="self.status='Go to Previous page'; return true;" onMouseOut="self.status=''">Back</a></td>
			  <td class="blue-normal" align="right" width="22%" valign="middle"> Search for&nbsp;&nbsp;&nbsp;&nbsp; </td>
			  <td width="27%" valign="middle"> 
			    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:99%" value="">
			  </td>
			  <td class="blue"  width="40%" valign="middle"> 
			    <table width="240" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
			      <tr> 
			        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
			            <a href="javascript:search();" class="b" onMouseOver="self.status='Search for Fullname'; return true;" onMouseOut="self.status=''">Search</a></td>
			        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
						<a href="javascript:filter();" class="b" onMouseOver="self.status='Filter'; return true;" onMouseOut="self.status=''">Filter</a></td>
			        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
						<a href="javascript:showall();" class="b" onMouseOver="self.status='Show all of employees'; return true;" onMouseOut="self.status=''">Show All</a></td>
			       </tr>
			    </table>
			  </td>
			</tr>        
          
          <tr> 
            <td class="title" height="50" align="center" colspan="4"> Select Employees For Expired Rule </td>
          </tr>
          <tr> 
            <td class="blue" height="30" align="center" colspan="4">Apply from: <%Response.Write arrlstFrom(2)%>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr> 
           <td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="200">
			 <tr>
               <td bgcolor="#FFFFFF" valign="top">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
	             <tr>
	              <td bgcolor="#617DC0"> 
	              <table width="100%" border="0" cellspacing="1" cellpadding="5">
	                <tr bgcolor="8CA0D1"> 
	                  <td class="blue" bgcolor="8CA0D1" width="194">&nbsp;
						<a href="javascript:sort(1);" class="c" onMouseOver="self.status='Order by Fullname'; return true;" onMouseOut="self.status=''">Full 
	                    Name</a> </td>
	                  <td class="blue" width="189">&nbsp;<a href="javascript:sort(2);" class="c" onMouseOver="self.status='Order by Job Title'; return true;" onMouseOut="self.status=''">Job 
	                    Title</a> </td>
	                  <td class="blue" align="center" width="8%">&nbsp;</td>
	                </tr>
	<%
				Response.Write strLast
	%>
	              </table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td bgcolor="#FFFFFF" height="20" class="blue" align="right"><a href="javascript:setchecked(1);" class="c" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
			      All</a>&nbsp;&nbsp;&nbsp;<a href="javascript:setchecked(0);" class="c" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''"> Clear All</a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
			  </tr>
			  <tr> 
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal" align="center"> 
			      <table width="120" border="0" cellspacing="5" cellpadding="0" height="20">
			        <tr> 
			          <td align="center" class="blue" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
			            <a href="javascript: assignment();" class="b" onMouseOver="self.status='Assign'; return true;" onMouseOut="self.status=''">Assign</a>
			          </td>
			          <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
			          <a href="javascript:BackPrevious('<%=strURLBack%>');" class="b" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
			        </tr>
			      </table>
			    </td>
			  </tr>
			  <tr>
			    <td bgcolor="#FFFFFF" height="20" class="blue-normal">&nbsp;&nbsp;*Click 
			      on each column header to sort the list by alphabetical order.
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
      </td>
    </tr>
    <tr> 
      <td align="right" valign="bottom" bgcolor="#E7EBF5">
		      <table width="100%" border="0" cellspacing="1" cellpadding="0" height="20">
		        <tr class="black-normal"> 
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPagesee")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPagesee")%>/<%=session("NumPagesee")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Go to previous page'; return true;" onMouseOut="self.status=''">Previous</a> /
		          <a href="javascript:next();" onMouseOver="self.status='Go to next page'; return true;" onMouseOut="self.status=''"> Next</a>&nbsp;&nbsp;&nbsp;</td>
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
			
end if%>			
	<input type="hidden" name="txttaskid" value="<%=taskID%>">
	<input type="hidden" name="txtkind" value="<%=kindact%>">
	<input type="hidden" name="txtURLBack" value="<%=strURLBack%>">
	
	<input type="hidden" name="txtRuleID" value="<%=intRuleID%>">
	
  <%if kindact<=2 then%>
	
	
	<input type="hidden" name="txthiddenstrproName" value="<%=Request.Form("txthiddenstrproName")%>">
	<input type="hidden" name="txthiddenstrproID" value="<%=Request.Form("txthiddenstrproID")%>">
	<input type="hidden" name="txtpreviouspage" value="<%=Request.Form("txtpreviouspage")%>">
	
  <%end if%>
  

		      </table>
</form>
</body>
<% 
	If gMessage="" and strAct="SAVE" then %>
				<SCRIPT language="javascript">
				<!--				
				document.selectemployee.action = '<%=strURLBack%>';
				document.selectemployee.target = "_self";
				document.selectemployee.submit();
					
					-->
				</SCRIPT>
<%end if%>
</html>

 