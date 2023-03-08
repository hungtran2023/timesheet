<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
'StatusList =Array("In used", "Broken", "Loss", "Liquidate/Charity", "Stock")

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
			strComputerName=Showlabel(rsSrc("ComputerName"))
			if strComputerName <>"" then strComputerName="<a href='javascript:inatlasnetwork(" & rsSrc("AtlasPCID") & ");' " &_
													"class='c' OnMouseOver = 'self.status=&quot;In Atlas Network&quot; ; return true' OnMouseOut =" &_
													"'self.status = &quot;&quot;'>" & Showlabel(rsSrc("ComputerName")) & "</a>"
			
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' class='blue'><a href='javascript:getdetail(" & rsSrc("PCID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;Computer Details&quot; ; return true' OnMouseOut =" &_
			         "'self.status = &quot;&quot;'>" & Showlabel(rsSrc("PC_Code")) & "</a></td>" &_
			         "<td valign='top' class='blue-normal' class='blue'>" & strComputerName & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("Username")) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("PublicName")) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("ComputerNote")) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("StatusDescription")) & "</td>" &_
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
	if strChoseMenu = "" then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

	strSearch=Request.Form("txtSearch")

	strSQL="SELECT a.PCID,a.PC_Code,b.ComputerName,c.Username,PublicName,ComputerNote,Outdated, AtlasPCID,StatusDescription FROM ATC_Computers a " & _
				"LEFT JOIN ATC_ComputerStatus d ON a.Outdated=d.StatusID " & _
				"LEFT JOIN ATC_AtlasPC b ON a.PCID=b.PCID " & _
				"LEFT JOIN ATC_Users c ON b.UserID=c.UserID "

	If Request.QueryString("act")<>"out" then
		if trim(strSearch<>"") then
			intSearchType=Request.Form("lstType")
			if CInt(intSearchType)=1 then 
				strSearch=" PC_Code like '%" & trim(strSearch) & "%'"
			elseif CInt(intSearchType)=2 then
				strSearch=" ComputerName like '%" & trim(strSearch) & "%'"
			elseif CInt(intSearchType)=3 then
				strSearch=" (Username like '%" & trim(strSearch) & "%' OR PublicName like '%" & trim(strSearch) & "%') "
			else
			    strSearch="  SeriNum like '%" & trim(strSearch) & "%'"
			end if
			
			strSQL=strSQL & " WHERE " & strSearch 
		end if

		intFilter=Request.Form("lstFilter")
		if intFilter="" then intFilter= -1
	
		if cint(intFilter)<>-1 then
			if strSearch<>"" THEN
				strSQL =strSQL & " AND Outdated=" & intFilter
			else
				strSQL =strSQL & " WHERE Outdated=" & intFilter
			end if
		end if
	else
		
		strYear_=Request.Form("year_")
		intType_=Request.Form("type_")
		
		strSearch=" '20' + SUBSTRING(PC_code, 3, 2) ='" & strYear_ & "' "
		if cint(intType_)<>0 then
			strSearch = strSearch & " AND "
			if Cint(intType_)=1 then
				strSearch = strSearch & " (a.Outdated=1 OR a.Outdated=7) AND b.TypeOfPC= 1 AND b.UserID IN (SELECT StaffID FROM ATC_Employees WHERE fgIndirect = 0)"
			elseif Cint(intType_)=2 then
				strSearch = strSearch & " (a.Outdated=1 OR a.Outdated=7) AND b.TypeOfPC= 1 AND b.UserID IN (SELECT StaffID FROM ATC_Employees WHERE fgIndirect = 1)"
			elseif Cint(intType_)=3 then
				strSearch = strSearch & " a.Outdated=1 AND (b.TypeOfPC= 2 OR b.TypeOfPC=3)"
			elseif Cint(intType_)=4 then
				strSearch = strSearch & " a.Outdated=1 AND b.TypeOfPC= 4 "
			elseif Cint(intType_)=5 then
				strSearch = strSearch & "  a.Outdated=1 AND b.TypeOfPC= 5 "
			elseif Cint(intType_)=6 then
				strSearch = strSearch & "  a.Outdated=1 AND b.TypeOfPC= 6 "
            elseif Cint(intType_)=7 then
				strSearch = strSearch & "  (a.Outdated=1 OR a.Outdated=7) AND b.TypeOfPC= 7 "				
			elseif Cint(intType_)=8 then
				strSearch = strSearch & "  a.Outdated=1 AND b.TypeOfPC= 8 "
			elseif Cint(intType_)=9 then
				strSearch = strSearch & " a.Outdated= 5 "
			elseif Cint(intType_)=11 then
				strSearch = strSearch & "  a.Outdated=1 AND b.TypeOfPC= 10 "	
		    elseif Cint(intType_)=12 then
				strSearch = strSearch & "  a.Outdated=1 AND b.TypeOfPC= 11 "	
			elseif Cint(intType_)=13 then		
				strSearch = strSearch & "  (a.Outdated=1 OR a.Outdated=7) AND b.TypeOfPC= 1 AND b.UserID IN (SELECT TPUserID FROM ATC_TPUsers )"
			elseif Cint(intType_)=14 then		
				strSearch = strSearch & "  a.Outdated=7"
			elseif Cint(intType_)=15 then		
				strSearch = strSearch & "  a.Outdated=8"
			else
				strSearch = strSearch & "  a.Outdated=1 AND (b.TypeOfPC= 1 OR b.TypeOfPC IS NULL)  AND b.UserID IS NULL"
			end if
		end if
		
		strSQL=strSQL & " WHERE " & strSearch 
	end if	
	
	strSQL=strSQL & " ORDER BY RIGHT(PC_code,7) DESC ,ComputerName"

'Response.Write 	strSQL

	Call GetRecordset(strSQL,rsData)
	
	
	strSql="SELECT * FROM ATC_ComputerStatus WHERE fgActivate=1 ORDER BY StatusDescription"
	Call GetRecordset(strSql,rsStatus)

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

function next() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage < numpage) {
	
		curpage=<%=intPage+1%>
		<%if Request.QueryString("act")="out" then%>
		document.navi.action = "ComputerList.asp?act=out&navi=" + curpage;
		<%else%>
		document.navi.action = "ComputerList.asp?navi=" + curpage;
		<%end if%>
		document.navi.target = "_self";
		document.navi.submit();
	}
}

function prev() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage > 1) {
		curpage=<%=intPage-1%>
		<%if Request.QueryString("act")="out" then%>
		document.navi.action = "ComputerList.asp?act=out&navi=" + curpage;
		<%else%>
		document.navi.action = "ComputerList.asp?navi=" + curpage;
		<%end if%>
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
		<%if Request.QueryString("act")="out" then%>
		document.navi.action = "ComputerList.asp?act=out&navi=" + intpage;
		<%else%>
		document.navi.action = "ComputerList.asp?navi=" + intpage;
		<%end if%>
		document.navi.target = "_self";
		document.navi.submit();		
	}
}

function sort(type) {
	document.navi.action = "ComputerList.asp?sorttype=" + type; //1: fullname, 2: jobtitle, 3: department
	document.navi.target = "_self";
	document.navi.submit();
}

function search() {
	var tmp = document.navi.txtsearch.value;
		tmp = escape(tmp);
	document.navi.action = "ComputerList.asp?search=" + tmp;
	document.navi.target = "_self";
	document.navi.submit();
	
}

function getdetail(varid){
	document.navi.txtID.value = varid;	
	document.navi.action = "computerdetail.asp?act=EDIT";
	document.navi.target = "_self";
	document.navi.submit();
}

function addnew(){

	document.navi.txtID.value=-1
	document.navi.action = "computerdetail.asp?act=ADD";
	document.navi.target = "_self";
	document.navi.submit();
}


function inatlasnetwork(varid){

	document.navi.txtAtlasPCID.value=varid
	document.navi.action = "AtlasComputer.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

function GetFilterByStatus(own)
{
	document.navi.action = "ComputerList.asp";
	document.navi.target = "_self";
	document.navi.submit();
}
-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
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
                  <a href="javascript:addnew();" onMouseOver="self.status='Add a new harware'; return true;" onMouseOut="self.status=''">Add New</a>
<%end if%>
					</td>
                  <td class="blue-normal" align="right" width="15%" valign="middle"> 
                    Search for&nbsp; </td>
                  <td align="right" width="25%" valign="middle"> 
                    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:150" value="<%=Showvalue(varSearch)%>">
                  </td>
                  <td align="right" width="20%" valign="middle"> 
                    <select name='lstType' size='1' height='26px' width='70px' style='width:95%;height=24px;' class='blue-normal'>
						<option value='1'>PC code</option>
						<option value='2'>Computer Name</option>
						<option value='3' selected>User name</option>
						<option value='4'>Seri Num</option>
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
                  <td class="title" height="50" align="center" colspan="5">List of Hardware</td>
                </tr>     
                
                 <tr> 
                  <td class="blue" height="30" valign="top"  align="center" colspan="5">Status: &nbsp;&nbsp;
                  <select name='lstFilter' size='1' height='26px' width='150px' class='blue-normal' onchange="javascript:GetFilterByStatus(this)">
					<option value='-1' <%if cint(intFilter)=-1 then%>selected<%end if%>>&nbsp; </option>					
					<%if rsStatus.RecordCount>0 then
						rsStatus.MoveFirst
						Do while not rsStatus.EOF%>
						
							<option value='<%=rsStatus("StatusID")%>' <%if cint(intFilter)=rsStatus("StatusID") then%>selected<%end if%>><%=rsStatus("StatusDescription")%></option>
							
					<%		rsStatus.MoveNext
						loop
						
					  end if%>
					</select>
					</td>
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
                              <td class="blue" align="center"  width="13%">PC Code</td>
                              <td class="blue" align="center" width="15%">Computer Name</td>
                              <td class="blue" align="center" width="13%">Atlas User</td>
                              <td class="blue" align="center" width="14%">Other</td>
                              <td class="blue" align="center" width="35%">Note</td>
                              <td class="blue" align="center" width="10%">Status</td>                              
                            </tr>
<%
	Response.Write(strLast)
%>                            
                          </table>
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal"> 
                                &nbsp;&nbsp;</td>
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
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="txtID" value="">
<input type="hidden" name="txtAtlasPCID" value="">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
<input type="hidden" name="year_" value="<%=strYear_%>">
<input type="hidden" name="type_" value="<%=intType_%>">

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