<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' function: outBody
' Description: table of list project
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize)
dim strpname
	strOut = ""
	
	if not rsSrc.EOF then
		For i = 1 to psize
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
			strAct = ""
			if rsSrc("fgActivate") = true then	strAct = "checked"
			
			strpname = replace(rsSrc("ProjectName"), "'", "%27")
			strpname = replace(strpname, chr(34), "%22")
			
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' width='15%' class='blue'>" & Showlabel(rsSrc("ProjectID")) & "</td>" &_
			         "<td valign='top' width='45%' class='blue-normal'>" & Showlabel(rsSrc("ProjectName")) & "</td>" &_
			         "<td valign='top' width='13%' class='blue-normal' align='center'><a href='javascript:goRelativepage(" & chr(34) &_
			         rsSrc("ProjectID") & chr(34) & ", " & chr(34) & strpname & chr(34) & ", 2);' OnMouseOver = 'self.status=&quot;Projects Assignment&quot; ; return true;' " &_
			         "OnMouseOut = 'self.status = &quot;&quot;'>...</a></td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function
'--------------------------------------------------------------------------------
'Built header of sheet
'--------------------------------------------------------------------------------
function OutHeader(Byval sortType, ByVal Col)

	dim strOut,strTypeForName,strImageforName
	dim strTypeForID,strImageforID
	
	strTypeForName="&quot;ASC&quot;"	
	strTypeForID="&quot;ASC&quot;"
	strImageforName=""
	strImageforID=""
	
	if Col="ProjectName" then 
		if sortType="ASC" then 
			strTypeForName="&quot;DESC&quot;"
		else
			strTypeForName="&quot;ASC&quot;"
		end if
	elseif  Col="ProjectID" then
		if sortType="ASC" then 
			strTypeForID="&quot;DESC&quot;"
		else
			strTypeForID="&quot;ASC&quot;"
		end if
	end if
    strOut="<tr bgcolor='8CA0D1'> "  & chr(13)
    strOut=strOut & "<td class='blue' bgcolor='8CA0D1' width='20%'>&nbsp;<a href='javascript:sort(&quot;ProjectID&quot;," & strTypeForID & ");' class='c' onMouseOver='self.status=&quot;Order by ProjectID&quot; ; return true;' onMouseOut='self.status=&quot;&quot;'>APK</a></td>" & chr(13)
    strOut=strOut & "<td class='blue' width='60%'>&nbsp;<a href='javascript:sort(&quot;ProjectName&quot;," & strTypeForName & ");' class='c' onMouseOver='self.status=&quot;Order by ProjectID&quot; ; return true;' onMouseOut='self.status=&quot;&quot;'>Project Name</a></td>" & chr(13)
    strOut=strOut & "<td class='blue' align='center' width='20%'>Assignment</td></tr>" & chr(13)
			         
	OutHeader = strOut
end function
'--------------------------------------------------------------------------------
'Built one row for activate or deactivate link
'--------------------------------------------------------------------------------
Function GetActivateLink(byval varActive)
	dim strOut,strName,typeActived,strCheck
	if varActive=0 then
		strName="Activate"
		strCheck="Choose"
		typeActived=1
	else
		typeActived=0
		strCheck="Uncheck"
		strName="Deactivate"
	end if	
	
	strOut="<tr ><td>"
	strOut = strOut & "<table width='100%' height='25' border='0' cellspacing='0' cellpadding='0'><tr>"
	strOut = strOut & "<td valign='middle' width='100%' class='blue' align='right'>"
	if fgRegistor then
		strOut = strOut & "<a href='javascript:activeproject(" & typeActived & ");' class='c'>"  
		strOut = strOut & strName & " Project&nbsp;&nbsp;"
		strOut = strOut & "</a>"
	end if
	strOut = strOut & "</td></tr>"
	strOut = strOut & "<tr><td bgcolor='#FFFFFF' height='20' class='blue-normal'>"
	strOut = strOut & "&nbsp;&#8226;&nbsp;Click on ProjectID and Project Name column header to sort the list by alphabetical order.</td></tr>"
	if fgRegistor then
		strOut = strOut & "<tr><td bgcolor='#FFFFFF' height='20' class='blue-normal'>"
		strOut = strOut & "&nbsp;&#8226;&nbsp;" & strCheck & " the checkbox, then click " & strName & " Project to update project.</td></tr>"
	end if
	strOut = strOut & "</table>"
	strOut = strOut & "</td></tr>"
	
	GetActivateLink	= strOut                        
End Function
'--------------------------------------------------------------------------------
'Retrive data from project
'--------------------------------------------------------------------------------
Sub GetProjectData(byval intSearchType, byval intActive,byval strSearch,byref rsReturn)
	dim strConnect,objDb,strQuery
	
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
	
	if strSearch<>"" then
		strSearch = replace(strSearch, "%", "")
		strSearch = Replace(strSearch, "'", "''")
	end if
		
	strQuery = "SELECT ProjectID, ProjectName, fgActivate FROM ATC_Projects a " & _
					"LEFT JOIN ATC_Companies b ON a.CompanyID=b.CompanyID " & _
				"WHERE fgDelete = 0 AND CHARINDEX('_',ProjectID)=0"
	
	strTemp=getWherePhase("a",session("USERID"))

	if not fgRight then strQuery = strQuery & " AND " & strTemp
	'Response.Write strQuery	
	strQuery= strQuery & " AND fgActivate=" & intActive
	if cint(intSearchType)=1 then
		strQuery= strQuery & "AND ProjectName Like '%" & strSearch & "%'"
	elseif cint(intSearchType)=2 then
		strQuery= strQuery & "AND ProjectID Like '%" & strSearch & "%'"
	elseif cint(intSearchType)=3 then
		strQuery= strQuery & " AND (ProjectID Like '%" & strSearch & "%' OR ProjectName Like '%" & strSearch & "%')"
	else
		strQuery= strQuery & " AND CompanyName Like '%" & strSearch & "%'"
	end if
	

	If objDb.openRec(strQuery) Then
		objDb.recDisConnect
		set rsReturn = objDb.rsElement.Clone
		if not objDb.noRecord then 
			session("NumPage")=pageCount(rsReturn.Clone, PageSize)
		else
			gMessage = "No results found."
		end if
		objDb.CloseRec
	Else
		gMessage = objDb.strMessage	  
	End if
	
	Set objDb = Nothing
	
end sub

'--------------------------------------------------------------------------------
'Update activate field data from project
'--------------------------------------------------------------------------------
Sub UpdateActivate(byval strProjectID, byval intActive)
	dim strConnect,objDb,strQuery
	
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	
	if objDb.dbConnect(strConnect) then	
		strQuery = "UPDATE ATC_Projects SET fgActivate=" & intActive & ",DeActivatedUser=" & session("USERID") & " WHERE ProjectID IN (" & strProjectID & ")"
		if not objDb.runActionquery(strQuery) then
		  gMessage = objDb.strMessage
		else
		  gMessage = "Saved successfully."
		end if
		
	else
		gMessage = objDb.strMessage
	end if
	
	
	
	Set objDb = Nothing
	
end sub

'--------------------------------------------------------------------------------
'********************************************************************************
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, numpage
	Dim rsProject, gMessage, PageSize
	dim varSearch,varSearchType,varActive,varPage,varSortType,varCol,varUpdate
	Dim strProjectID,fgRegistor

	varSearch = trim(Request.QueryString("search"))

	varSearchType = Request.Form("lbSeachType")
	if varSearchType="" then varSearchType=3
	
	varActive=Request.Form("lbActived")
	if varActive="" then varActive=1
	
	varPage=Request.QueryString("Page")
	
	varSortType=Request.QueryString("Sorttype")
	varCol=Request.QueryString("Col")
	
	varUpdate=Request.QueryString("active")
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check ACCESS right
'-----------------------------------
'tmp= /timesheet/management/project/listofproject.asp 
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
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if
	
'--------------------------------------------------
' Check VIEWALL project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "View all projects" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------
' Check Approving Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRegistor = False
	Else
		varGetRight = session("RightOn")
		fgRegistor = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "registration" Then
				fgRegistor = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
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

'--------------------------------------------
' Get Full Name and menu
'--------------------------------------------
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

'----------------------------------	
' Make list of menu
'----------------------------------
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
	if strChoseMenu = "" then strChoseMenu = "AG"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

Message = ""
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
	if Request.ServerVariables("QUERY_STRING")="" or Request.QueryString("outside")<>"" then
		Call freelistPro
	end if
end if
'--------------------------------------------------
' Get project item
'--------------------------------------------------

if varUpdate<>"" then
	strProjectID=Request.Form("txthiddenstrActive")
	if strProjectID<>"" then
		Call UpdateActivate(left(strProjectID,len(strProjectID)-1),varUpdate)
	end if
end if

'varPage="" --> didn't retrive data from database
if varPage="" then	
	varPage=1
	call GetProjectData(varSearchType,varActive,varSearch,rsProject)	
	if rsProject.RecordCount>0 then
		if not isEmpty(session("rsProject")) then session("rsProject") = empty
		set session("rsProject")=rsProject.Clone
	else
		if not IsEmpty(Session("rsProject")) then set rsProject = session("rsProject")
		if rsProject.RecordCount>o then rsProject.MoveFirst
	end if
else
	set rsProject = session("rsProject")
	rsProject.MoveFirst
	if varSortType<>"" then
		rsProject.Sort = varCol & " " & varSortType		
	end if
	rsProject.Move (cint(varPage)-1)* PageSize
	
end if

strLast = Outbody(rsProject, PageSize)

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_pro.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", varPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", session("NumPage"))	
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>
<script>
<!--
var objNewWindow;
var strURL="TPStaffAssignment.asp";

function window_onunload() {
	if((objNewWindow)&&(!objNewWindow.closed))
		objNewWindow.close();
}
//-->

function next() {
var curpage = <%=varPage%>;
var numpage = <%=session("NumPage")%>;
	if (curpage < numpage) {
		//document.navi.action = "listofproject.asp?navi=NEXT"
		document.navi.action = strURL + "?page=" + (curpage + 1)
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=varPage%>;
var numpage = <%=session("NumPage")%>;
	if (curpage > 1) {
		document.navi.action = strURL + "?page=" + (curpage - 1)
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function go() {
	var numpage = <%=session("NumPage")%>;
	var curpage = <%=varPage%>;
	var intpage = document.navi.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.navi.action = strURL + "?page=" + intpage
		document.navi.target = "_self";
		document.navi.submit();		
	}
}

function sort(column,type) {
	document.navi.action = strURL + "?page=1&sorttype=" + type + "&Col=" + column; //1: id, 2: name, 3: activate
	document.navi.target = "_self";
	document.navi.submit();
}

function checkass (value) {
  strID="";
  with (document.navi) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].type == "checkbox") && (elements[ii].checked==value)) {
			strID = strID + "'" + elements[ii].value + "',";
		}
	}
  }
  return(strID)
}

function activeproject(type) {
	var strID
	strID=checkass((type==1));
	
	if (strID!="") {
		document.navi.txthiddenstrActive.value = strID;
		document.navi.action = strURL + "?active=" + type;
		document.navi.target = "_self";
		document.navi.submit();
	}
	else
		alert("Please select at least one project.")
}

function search() {
	var tmp = document.navi.txtsearch.value;
	tmp = alltrim(tmp);
	document.navi.txtsearch.value = tmp;
	tmp = escape(tmp);
	
	document.navi.action = strURL + "?search=" + tmp;
	document.navi.target = "_self";
	document.navi.submit();	
}

function goRelativepage(varid, varname, varkind) {
	window.document.navi.txthiddenstrproID.value = varid;
	window.document.navi.txthiddenstrproName.value = varname;
	if(varkind==1)
		varfilename = "TPsubtask.asp";
	else 
		varfilename = "TPassignment.asp";


	window.document.navi.action = varfilename + "?outside=1";
	window.document.navi.target = "_self";
	window.document.navi.submit();
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
              <table width="100%" border="0" cellpadding="0" cellspacing="0" >
                <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
					<td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
				</tr>
                <tr> 
                  <td class="blue-normal" width="10%" valign="middle" align="right"> Search for&nbsp; </td>
                  <td align="right" width="20%" valign="middle"> <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:120" value="<%=Showvalue(varSearch)%>">
                   </td>
                  <td align="center" width="25%" valign="middle" > 
                    <select name="lbSeachType" class='blue-normal'>
                      <option value="1" <%if cint(varSearchType)=1 then %>selected<%end if%>>Project Name</option>
                      <option value="2"<%if cint(varSearchType)=2 then %>selected<%end if%>>APK</option>
                      <option value="3"<%if cint(varSearchType)=3 then %>selected<%end if%>>Project Name &amp; APK</option>
                      <option value="4"<%if cint(varSearchType)=4 then %>selected<%end if%>>Company Name</option>                                            
                    </select>
                  </td>
                       
                  <td class="blue" align="left" width="40%" valign="middle"> 
                    <table width="100" border="0" cellspacing="5" cellpadding="0" height="20">
                      <tr> 
                        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
                            <a href="javascript:search();" class="b" onMouseOver="self.status='Search for ProjectID'; return true;" onMouseOut="self.status=''">Search</a></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td class="title" height="50" align="center" colspan="5">List of Projects for Assigment</td>
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
<%
	Response.Write(OutHeader(varSortType,varCol))
	Response.Write(strLast)
%>                            
                          </table>
                        </td>
                      </tr>
<%if strLast<>"" then Response.Write(GetActivateLink(varActive))%>                      
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
<input type="hidden" name="txthiddenstrproID" value="">
<input type="hidden" name="txthiddenstrproName" value="">
<input type="hidden" name="txthiddenstrActive" value="">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>