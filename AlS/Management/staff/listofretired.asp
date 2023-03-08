<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%

'****************************************
' Function: Outbody
' Description: 
' Parameters: source recordset, number of lines on one page
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize)
	strOut = ""
	if not rsSrc.EOF then
		For i = 1 to psize
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' width='29%' class='blue'><a href='javascript:getdetail(" & rsSrc("StaffID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;View Timesheet&quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>"& Showlabel(rsSrc("Fullname")) & "</a></td>" &_
			         "<td valign='top' width='31%' class='blue-normal'>" & Showlabel(rsSrc("JobTitle")) & "</td>" &_
			         "<td valign='top' width='32%' class='blue-normal'>" & Showlabel(rsSrc("Department")) & "</td>" &_
					 "<td valign='top' width='8%' class='blue-normal' align='center'>" &_
					 "<input type='checkbox' name='chkrecover' value='" & rsSrc.Bookmark & "'></td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function
'------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, PageSize, fgUpdate, fgRight

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 
	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
				exit for
			end if
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
	
'-------------------------------
' Get Fullname and Job Title
'-------------------------------
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
'-----------------------------
' Make list of menu
'-----------------------------
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
	if strChoseMenu = "" then strChoseMenu = "AB"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

gMessage = ""
'----------------------------
' free session variables
'----------------------------
Call freeListpro
Call freeProInfo
Call freeAssignment
Call freeAssignRight
Call freeShort
Call freeSinglepro
Call freeSumpro

if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
	if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
		Call freelistEmp
		Session("CurPagele") = 1
	end if
end if

strAct = Request.QueryString("act")

if strAct="SAVE" and session("READYPER") then
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	gMessage = ""
	if ret then
		objDb.cnDatabase.BeginTrans
		set rsPerson = session("rsPerson")
		strQuery = ""
		For ii = 1 To Request.Form("chkrecover").Count
			varTmp = Request.Form("chkrecover")(ii)
			rsPerson.Bookmark = int(varTmp)
			strQuery = strQuery & "UPDATE ATC_PersonalInfo SET fgDelete = 0 WHERE PersonID = " & rsPerson("StaffID") & chr(13)
		Next
		if not objDb.runActionQuery(strQuery) then
			gMessage = objDb.strMessage
		end if
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
		rsPerson.Close
		set rsPerson = nothing
		session("rsPerson") = empty
		session("READYPER") = false
	end if
End if
		
If session("READYPER")<> True Then
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
	strQuery = "select a.StaffID, b.FirstName + ' ' + isnull(b.MiddleName, '') + ' ' + b.LastName as FullName, a.DirectLeaderID, " &_
			"b.Gender, b.NationalityID, b.CountryID, a.DepartmentID, a.fgIndirect, isNull(c.jobTitle, '') jobTitle, c.JobTitleID, isNull(d.Department, '') Department " &_
			"from ATC_Employees a LEFT JOIN ATC_PersonalInfo b " &_
			"ON a.StaffID = b.PersonID LEFT JOIN HR_CurrentJobtitle c ON a.StaffID = c.StaffID " &_
			"LEFT JOIN ATC_Department d ON a.DepartmentID = d.DepartmentID WHERE b.fgDelete=1 and b.UserType=1 Order by b.FirstName, b.MiddleName, b.LastName"
	If objDb.openRec(strQuery) Then
		objDb.recDisConnect
		IF not objDb.noRecord then
			set rsPerson = objDb.rsElement.Clone
			session("READYPER") = true
			rsPerson.MoveFirst
			session("NumPagele") = pageCount(rsPerson, PageSize)
			if isEmpty(Session("CurPagele")) then 
				Session("CurPagele") = 1
			else
				if Session("CurPagele") > Session("NumPagele") then
					Session("CurPagele") = Session("NumPagele")
				elseif Session("CurPagele") = 0 then
					Session("CurPagele") = 1
				end if
			end if
			set session("rsPerson") = rsPerson
			session("fgShowle") =  0 ' show all
			session("arrSort") = array(true, true, true)
		else
			Session("CurPagele") = 0
			Session("NumPagele") = 0
			session("fgShowle") = -1
		end if
	Else
		gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
Else
	set rsPerson = session("rsPerson")
End if

if fgExecute and session("READYPER") then
	varSort = Request.QueryString("sorttype")
	if varSort <> "" then
		strtypesort = ""
		if not isEmpty(session("arrSort")) then
			arrTmp = session("arrSort")
			if arrTmp(varSort-1) then
				strtypesort = "ASC"
			else
				strtypesort = "DESC"
			end if
			arrTmp(varSort-1) = not arrTmp(varSort-1)
			session("arrSort") = arrTmp
			set arrTmp = nothing
		end if
		select case varSort
			case 1 rsPerson.Sort = "Fullname " & strtypesort
			case 2 rsPerson.Sort = "Jobtitle " & strtypesort
			case 3 rsPerson.Sort = "Department " & strtypesort
		end select
	end if

	varNavi = Request.QueryString("navi")
	if varNavi <> "" then
		tmpi = session("CurPagele")
		select case varNavi
			case "PREV"
				if tmpi > 1 then
					tmpi = tmpi - 1
				else
					tmpi = 1
				end if
			case "NEXT"
				if tmpi < Session("NumPagele") then
					tmpi = tmpi + 1
				else
					tmpi = Session("NumPagele")
				end if
		End select
		session("CurPagele") = tmpi
	end if

	varGo = Request.QueryString("Go")
	if varGo <> "" then Session("CurPagele") = CInt(varGo)

	varShowAll = Request.QueryString("showall")
	if varShowAll <> "" then 
		session("fgShowle") = 0
		rsPerson.Filter = ""
		rsPerson.MoveFirst
		session("NumPagele") = pageCount(rsPerson, PageSize)
		Session("CurPagele") = 1
	end if

	varSearch = Request.QueryString("search")
	if varSearch<>"" then
		'making custom recordser
		varSearch = replace(varSearch, "%", "")
		varSearch = replace(varSearch, "#", "")
		criteria = trim(varSearch)
		if criteria <> "" then
			if Instr(criteria, "'")>0 then
				criteria = "#" & criteria & "#"
			else
				criteria = "'%" & Replace(criteria, "'", "''") & "%'"
			end if
			rsPerson.Filter = "Fullname Like " & criteria
		else
			rsPerson.MoveLast
			rsPerson.MoveNext
		end if
		If rsPerson.EOF then ' no result
			gMessage = "No results found."
			session("fgShowle") = 0
			rsPerson.Filter = ""
			rsPerson.MoveFirst
			Session("CurPagele") = 1
			session("NumPagele") = pageCount(rsPerson, PageSize)
		else
			session("fgShowle") = 1 ' show the result
			rsPerson.MoveFirst
			session("NumPagele") = pageCount(rsPerson, PageSize)
			Session("CurPagele") = 1
			gMessage = ""
		end if
	end if

	varFilter = Request.QueryString("filter") '0 or 1
	if varFilter<>"" then
		if isEmpty(session("filteremp")) and not isNull(session("filteremp")) then		
			session("fgShowle") = 0
			rsPerson.MoveFirst
			Session("CurPagele") = 1
			session("NumPagele") = pageCount(rsPerson, PageSize)
		else				
			tmp = session("Filteremp")
			rsPerson.Filter = tmp
			If not rsPerson.EOF then
				session("fgShowle") = 1 ' show the result
				rsPerson.MoveFirst
				session("NumPagele") = pageCount(rsPerson, PageSize)
				Session("CurPagele") = 1
				gMessage = ""
			else
				gMessage = "No results found."
				rsPerson.Filter = ""
				session("fgShowle") = 0
				rsPerson.MoveFirst
				Session("CurPagele") = 1
				session("NumPagele") = pageCount(rsPerson, PageSize)
			End if
		end if
	end if
end if 'fgExecute

If session("READYPER") then
	rsPerson.MoveFirst
	rsPerson.Move (session("CurPagele")-1)*PageSize
	strLast = Outbody(rsPerson, PageSize)
else
	strLast=""
end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_pro.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", session("CurPagele"))
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", session("NumPagele"))	
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/menu.js"></script>
<script language="javascript" src="../../library/library.js"></script>
<script>
<!--
var objEMFIWindow;

function filter() { //v2.0
  window.status = "";
  strFeatures = "top="+(screen.height/2-170)+",left="+(screen.width/2-126)+",width=252,height=340,toolbar=no," 
              + "menubar=no,location=no,directories=no";
  if ((objEMFIWindow) && (!objEMFIWindow.closed)) {
	objEMFIWindow.focus();
	
  } else {
	objEMFIWindow = window.open("../../management/project/empfilter.asp", "MyNewWindow1", strFeatures);
  }
  window.status = "Opened a new browser window.";
  
}
function window_onunload() {
	if((objEMFIWindow) && (!objEMFIWindow.closed))
		objEMFIWindow.close();
}
//-->

function next() {
var curpage = <%=session("CurPagele")%>
var numpage = <%=session("NumPagele")%>
	if (curpage < numpage) {
		document.navi.action = "listofretired.asp?navi=NEXT";
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagele")%>
var numpage = <%=session("NumPagele")%>
	if (curpage > 1) {
		document.navi.action = "listofretired.asp?navi=PREV";
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function go() {
	var numpage = <%=session("NumPagele")%>;
	var curpage = <%=session("CurPagele")%>;
	var intpage = document.navi.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.navi.action = "listofretired.asp?Go=" + intpage;
		document.navi.target = "_self";
		document.navi.submit();		
	}
}

function sort(type) {
	document.navi.action = "listofretired.asp?sorttype=" + type; //1: fullname, 2: jobtitle, 3: department
	document.navi.target = "_self";
	document.navi.submit();
}

function search() {
	var tmp = document.navi.txtsearch.value;
		tmp = escape(tmp);
	if (tmp != "") {
		document.navi.action = "listofretired.asp?search=" + tmp;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function showall() {
var tmp = <%=session("fgShowle")%>;
  if (tmp==1) {
	document.navi.action = "listofretired.asp?showall=1";
	document.navi.target = "_self";
	document.navi.submit();
  }
}

function checkass () {
  selection = false;
  with (document.navi) {
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

function recover() {
	if (checkass()==true) {
		document.navi.action = "listofretired.asp?act=SAVE";
		document.navi.target = "_self";
		document.navi.submit();
	}
	else
		alert("Please select at least one employee.")
}

function getdetail(varid){
	document.navi.txtuserid.value = varid;
	
	document.navi.action = "viewemployeeInfor.asp";
	document.navi.target = "_self";
	document.navi.submit();
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onunload="return window_onunload();">
<form name="navi" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			Response.Write(arrTmp(0))
			%>
          <tr> 
            <td>
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
					<td class="red" colspan="3" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
				</tr>
                <tr> 
                  <td class="blue-normal" align="right" width="30%" valign="middle">Search for&nbsp; </td>
                  <td class="blue" align="right" width="30%" valign="middle"> 
                    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:100%" value="<%=Showvalue(varSearch)%>">
                  </td>
                  <td class="blue" align="left" width="40%" valign="middle"> 
                    <table width="180" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
                      <tr> 
                        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
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
                  <td class="title" height="50" align="center" colspan="4"> List 
                    of retired employees</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
                              <td class="blue" width="29%">&nbsp;<a href="javascript:sort(1);" class="c" onMouseOver="self.status='Order by Fullname'; return true;" onMouseOut="self.status=''">Full Name</a></td>
                              <td class="blue" width="31%">&nbsp;<a href="javascript:sort(2);" class="c" onMouseOver="self.status='Order by Job Title'; return true;" onMouseOut="self.status=''">Job Title</a></td>
                              <td class="blue" width="32%">&nbsp;<a href="javascript:sort(3);" class="c" onMouseOver="self.status='Order by Department'; return true;" onMouseOut="self.status=''">Department</a></td>
                              <td class="blue" width="8%">&nbsp;</td>
                            </tr>
<%
	Response.Write(strLast)
%>                            
                          </table>
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue" align="right"> 
                                <%if fgUpdate then%><a href="javascript:recover();" class="c" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''"> 
                                Recover</a><%else%>Recover<%end if%>&nbsp;&nbsp;</td>
                            </tr>
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal">
                                &nbsp;&nbsp;*Click on each column header to sort 
                                the list by alphabetical order.</td>
                            </tr>
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal">&nbsp;&nbsp; 
                                *Choose the checkbox, then click Recover 
                                  to recover the employee.
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
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(1))

			%>		

			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
<input type="hidden" name="txtuserid" value="">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>