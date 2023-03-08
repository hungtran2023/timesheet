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
function Outbody(ByRef rsSrc,byval intPage, ByVal psize)
	dim intStart,intFinish
	strOut = ""

	if not rsSrc.EOF then
		
		rsSrc.AbsolutePage = intPage
		intStart = rsSrc.AbsolutePosition
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsData.PageSize - 1)
		End if
	
		For i = intStart to intFinish
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' class='blue'><a href='javascript:getdetail(" & rsSrc("SoftwareID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;Computer Details&quot; ; return true' OnMouseOut =" &_
			         "'self.status = &quot;&quot;'>" & Showlabel(rsSrc("SoftwareName")) & "</a></td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("Vendor")) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("Category")) & "</td>" &_
			         "<td valign='top' class='blue-normal' align='center'><input type='checkbox' name='chkRemove' value='" & rsSrc("SoftwareID") & "'></td>" &_	
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function

'***************************************************************
'
'***************************************************************
function ExecuteSQL(strSql)

	dim objDatabase
	dim strCnn
	dim blnReturn
	
	blnReturn = false	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	If objDatabase.dbConnect(strCnn) then		
		blnReturn= (objDatabase.runActionQuery(strSql))	
		strError="Update successfull."
		if not blnReturn then strError=objDatabase.strMessage		
	else
		strError=objDatabase.strMessage
	end if
	
	Set objDatabase = nothing
	ExecuteSQL=strError
	
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
	if strChoseMenu = "" then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
'Get list of data
'--------------------------------------------------
	strAct=Request.QueryString("act")

	if strAct="remove" then
		arrSoftware=Request.Form("chkRemove")
		if trim(arrSoftware)<>"" then
		
			strSql="DELETE FROM ATC_Softwares WHERE SoftwareID IN (" & arrSoftware & ")"
			strError= ExecuteSQL(strSql)
		
		end if
	end if

	strSearch=Request.Form("txtSearch")

	strSQL="SELECT SoftwareID, SoftwareName, b.Description as category,Vendor,NumberOfLicence " & _
			"FROM ATC_Softwares a LEFT JOIN ATC_SoftwareType b ON a.SoftTypeID=b.SoftTypeID "

	if trim(strSearch<>"") then
		intSearchType=Request.Form("lstType")
		if CInt(intSearchType)=1 then 
			strSearch=" SoftwareName like '%" & trim(strSearch) & "%'"
		elseif CInt(intSearchType)=2 then
			strSearch="  b.Description like '%" & trim(strSearch) & "%'"
		end if
		strSQL=strSQL & " WHERE " & strSearch 
	end if
	
	strSQL=strSQL & " ORDER BY SoftwareName"

	Call GetRecordset(strSQL,rsData)
	
'--------------------------------------------------
'Start Paging
'--------------------------------------------------

' Set the PageSize, CacheSize and populate the intPageCount

	rsData.PageSize=PageSize
' The Cachesize property sets the number of records that will be cached locally in memory	
	rsData.CacheSize=rsData.PageSize	
	intPageCount=rsData.PageCount
	intRecordCount=rsData.RecordCount
	
' Checking to make sure that we are not before the start or beyond end of the recordset
' If we are beyond the end, set the current page equal the last page of the recordset.
' If we are before the start, set the current page equal the start of the recordset
	
	intPage=Request.QueryString("Navi")

	if intPage="" then intPage=1
	
	if cint(intPage)>Cint(intPageCount) then intPage=intPageCount
	if cint(intPage)<=0 then intPage=1

'--------------------------------------------------
'End Paging	
'--------------------------------------------------
	strLast=Outbody(rsData,intPage,PageSize)
	'strLast=Outbody(rsData,PageSize)
	
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_pro.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intPageCount)	
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
function search() {
	var tmp = document.navi.txtsearch.value;
		tmp = escape(tmp);
	document.navi.action = "SoftwareList.asp?search=" + tmp;
	document.navi.target = "_self";
	document.navi.submit();
	
}

function getdetail(varid){
	document.navi.txtID.value = varid;	
	document.navi.action = "SoftwareDetail.asp?act=EDIT";
	document.navi.target = "_self";
	document.navi.submit();
}

function addnew(){

	document.navi.txtID.value=-1
	document.navi.action = "SoftwareDetail.asp?act=ADD";
	document.navi.target = "_self";
	document.navi.submit();
}


function checkedAll (own) {

	var aa= document.getElementById('navi');
	var chkName
	
	chkName="chkRemove"
		
	for (var i =0; i < aa.elements.length; i++) 
	{
		strName=String(aa.elements[i].name)
		
		if (aa.elements[i].type == "checkbox" && strName.indexOf(chkName)>-1)
			aa.elements[i].checked = own.checked;
	}
}


function RemoveSoftware()
{
	var agree=confirm("Are you sure you want to delete selected software(s)?");
	if (agree)
	{
		window.document.navi.action = "SoftwareList.asp?act=remove"	;		
		window.document.navi.submit();
		return true ;
		}
	else
		return false ;
}

function next() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage < numpage) {
	
		curpage=<%=intPage+1%>
		document.navi.action = "SoftwareList.asp?navi=" + curpage;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage > 1) {
		curpage=<%=intPage-1%>
		document.navi.action = "SoftwareList.asp?navi=" + curpage;
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function go() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	var intpage = document.navi.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.navi.action = "SoftwareList.asp?navi=" + intpage;
		document.navi.target = "_self";
		document.navi.submit();		
	}
}

-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
					<td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
				</tr>
                <tr> 
                  <td class="blue" width="10%" valign="middle">&nbsp;
<%if fgUpdate then%>
                  <a href="javascript:addnew();" onMouseOver="self.status='Add a new employee'; return true;" onMouseOut="self.status=''">Add New</a>
<%end if%>
					</td>
                  <td class="blue-normal" align="right" width="15%" valign="middle"> 
                    Search for&nbsp; </td>
                  <td align="right" width="25%" valign="middle"> 
                    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:150" value="<%=Showvalue(varSearch)%>">
                  </td>
                  <td align="right" width="20%" valign="middle"> 
                    <select name='lstType' size='1' height='26px' width='70px' style='width:95%;height=24px;' class='blue-normal'>
						<option value='1'>Software Name</option>
						<option value='2'>Category</option>
					</select>
                    
                  </td>
                  <td class="blue" width="30%" valign="middle"> 
                    <table width="100" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
                      <tr> 
                        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
                            <a href="javascript:search();" class="b" onMouseOver="self.status='Search for Fullname'; return true;" onMouseOut="self.status=''">Search</a></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td class="title" height="50" align="center" colspan="5">Software list</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
                              <td class="blue" align="center"  width="35%">Software name</td>
                              <td class="blue" align="center" width="27%">Vendor</td>
                              <td class="blue" align="center" width="28%">Category</td>
                              <td class="blue" align="center" width="10%"><input type='checkbox' name='chkAll' value='1' onclick='checkedAll(this);' ></td> 
                            </tr>
<%
	Response.Write(strLast)
%>
                          </table>
<%if fgUpdate then%>                          
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue" align="right"><a href="javascript:RemoveSoftware()">Remove</a>&nbsp;&nbsp;</td>
                            </tr>
                          </table>
<%End if%>                          
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
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="txtID" value="">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>

<SCRIPT language=JavaScript1.2>
var hotkey=13
if (document.layers)
document.captureEvents(Event.KEYPRESS)
function backhome(e){
	if (document.layers){
		if (e.which==hotkey)
			search();}
	else if (document.all){
		if (event.keyCode==hotkey){
			event.keyCode = 0;
			search();}
	}
}
document.onkeypress=backhome
</SCRIPT>

</body>
</html>