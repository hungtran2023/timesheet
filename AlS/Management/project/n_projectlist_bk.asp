<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'*********************************************************
'Generate report
'*********************************************************
Function ATSsql(byval strF,byval strT)
	dim strATS
	If year(strF)<> year(strT) then
		strATS="(SELECT * FROM "
		For ii=year(strF) To year(strT)
			strATS=strATS & selectTable(ii)
			If ii<>	year(strT) then
				strATS=strATS & " UNION ALL SELECT * FROM "
			else
				strATS=strATS & ")"
			end if
		Next
	else
		strATS=selectTable(year(strT))
	end if
	ATSsql=strATS
End function
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

Function Outbody(ByRef rsSrc, ByVal psize)
	strOut = ""
	dim strLinkPro,strLinkInvoice,strColor,strDate,i
	dim strFileLink,strServerPath
	
	strServerPath="..\..\data\CSO\"
	
	If Not rsSrc.EOF Then
		For i = 1 To psize
			strColor = "#FFF2F2"
			If i Mod 2 = 0 Then	strColor = "#E7EBF5"
			strDate=ConvertDate(cdate(rsSrc("DateTransfer")))
			strLinkPro="<a href='javascript:viewpro(&quot;" & rsSrc("ProjectKey") & "&quot;,&quot;" & rsSrc("fgStatus") & "&quot;,&quot;" & rsSrc("DateTransfer") & "&quot;," & rsSrc("ManagerID") & ");' " &_
					         "class='c' OnMouseOver = 'self.status=&quot;Project Detail&quot; ; return true' OnMouseOut =" &_
					         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("ProjectKey")) & "</a>"
			If rsSrc("fgStatus") = "New" and not fgApproval then strLinkPro=Showlabel(rsSrc("ProjectKey"))
			
			strSignContract=""	
			if rsSrc("SignContract")=1 then 
				strSignContract="<img src='../../images/notyet.gif'>"
			elseif rsSrc("SignContract")=2 then
				strSignContract="<img src='../../images/icon_doc_download.gif' border=0>"
			end if
			
			if rsSrc("CSOFileName")<>"" then strSignContract="<a href='" & strServerPath & rsSrc("CSOFileName") & "'>" & strSignContract & "</a>"
			
			strUtilised=IIF(cint(rsSrc("ProjectKey2"))=1,"<img src='../../images/yes.gif'>","")
			strBillable=IIF(rsSrc("Billable"),"<img src='../../images/yes.gif'>","")
			
			strLinkInvoice="<a href='javascript:viewinvoice(&quot;" & rsSrc("ProjectKey") & "&quot;);' " &_
					         "class='c' OnMouseOver = 'self.status=&quot;Project Invoice&quot; ; return true' OnMouseOut =" &_
					         " 'self.status = &quot;&quot;'>--</a>"
			
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' class='blue'>" & strLinkPro & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("ProjectName")) & "</td>" &_
			         "<td valign='top' class='blue' align='center'>" & strLinkInvoice & "</td>" &_
			         "<td valign='top' class='blue' align='center'>" & strSignContract & "</td>" &_
			         "<td valign='top' class='blue' align='center'>" & Showlabel(rsSrc("fgStatus")) & "</td>" &_
			         "</tr>" & chr(13)                                                                                                                                                                                                                                                                                                                                                                                                    
					         
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	End If
	Outbody = strOut
End Function
'--------------------------------------------------------------------------------
'Built header of sheet
'--------------------------------------------------------------------------------
function OutHeader(Byval sortType, ByVal Col)

	dim strOut,strTypeForName
	dim strTypeForID,strTypeForTranfer
	
	strTypeForName="&quot;ASC&quot;"	
	strTypeForID="&quot;ASC&quot;"
	strTypeForTranfer="&quot;ASC&quot;"
	
	if Col="ProjectName" then 
		if sortType="ASC" then 
			strTypeForName="&quot;DESC&quot;"
		else
			strTypeForName="&quot;ASC&quot;"
		end if
	elseif  Col="ProjectKey" then
		if sortType="ASC" then 
			strTypeForID="&quot;DESC&quot;"
		else
			strTypeForID="&quot;ASC&quot;"
		end if
	elseif Col="DateTransfer" then
		if sortType="ASC" then 
			strTypeForTranfer="&quot;DESC&quot;"
		else
			strTypeForTranfer="&quot;ASC&quot;"
		end if
	end if
	
    strOut="<tr bgcolor='#8CA0D1'>"
	strOut=strOut & "<td class='blue' bgcolor='#8CA0D1' align='center' width='18%'>"
	strOut=strOut & "<a href='javascript:sort(&quot;ProjectKey&quot;," & strTypeForID & ");' onMouseOver='self.status=&quot;Sort by Project Key&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true' class='c'>APK</a></td>"
	strOut=strOut & "<td class='blue' align='center' width='58%'>"
	strOut=strOut & "<a href='javascript:sort(&quot;ProjectName&quot;," & strTypeForName & ");' onMouseOver='self.status=&quot;Sort by Project Name&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true' class='c'>Project Name</a></td>"
	strOut=strOut & "<td class='blue' align='center' width='7%'>Invoice</td>"
	'strOut=strOut & "<td class='blue' align='center' width='7%'>APPR</td>"
	strOut=strOut & "<td class='blue' align='center' width='10%'>CSO</td>"
	'strOut=strOut & "<td class='blue' align='center' width='7%'>Contract</td>"
	'strOut=strOut & "<td class='blue' align='center' width='7%'>Utilised</td>"		
	'strOut=strOut & "<td class='blue' align='center' width='7%'>Billable</td>"	
'strOut=strOut & "<a href='javascript:sort(&quot;DateTransfer&quot;," & strTypeForTranfer & ");' onMouseOver='self.status=&quot;Sort by Transfered Date&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true' class='c'>Transfered Date</a></td>"
	strOut=strOut & "<td class='blue' align='center' width='7%'>Status</td></tr>"
								         
	OutHeader = strOut
end function
'--------------------------------------------------------------------------------
'Retrive data from project
'--------------------------------------------------------------------------------
Sub GetProjectData(byval intSearchType, byval intBillable,byval strSearch,byval intBooked,byref rsReturn)
	dim strConnect,objDb,strQuery
	
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)	
	if strSearch<>"" then
		strSearch = replace(strSearch, "%", "")
		strSearch = Replace(strSearch, "'", "''")
	end if
	strQuery = "SELECT b.ProjectID AS ProjectKey,Projectkey2, " & _
				" ProjectName, ISNULL(CSOFilename,'') as CSOFilename, DateTransfer, (CASE WHEN CHARINDEX('___',a.ProjectID,7) > 1 THEN 'New' ELSE 'Issued' END) AS fgStatus, CSOApproval,SignContract,CSOCompleted,ManagerID,billable " & _  
				" FROM ATC_ProjectStage a INNER JOIN ATC_Projects b ON a.ProjectID=b.ProjectID LEFT JOIN ATC_Companies d ON b.CompanyID=d.CompanyID " & _
				" WHERE b.fgDelete = 0"
	strQuerySearch=""
	
	if not fgRight then strQuery = strQuery & " AND " & getWherePhase("b",session("USERID"))
	
	if intBillable = 1 then 
		strQuerySearch= strQuerySearch & " AND ProjectKey2=1"
	elseif intBillable = 0 then 
		strQuerySearch= strQuerySearch & " AND (ProjectKey2=5 OR ProjectKey2=7)"
	end if
		
	if cint(intSearchType)=1 then
		strQuerySearch= strQuerySearch & " AND ProjectName Like '%" & strSearch & "%'"
	elseif cint(intSearchType)=2 then
		strQuerySearch= strQuerySearch & " AND b.ProjectID Like '%" & strSearch & "%'"
	elseif cint(intSearchType)=3 then
		strQuerySearch= strQuerySearch & " AND (b.ProjectID Like '%" & strSearch & "%' OR ProjectName Like '%" & strSearch & "%')"
	else
		strQuerySearch= strQuerySearch & " AND d.CompanyName Like '%" & strSearch & "%'"
	end if

	if cint(intBooked)>0 then 
		strQuerySearch= strQuerySearch & " AND b.ProjectID IN (" & _
							"SELECT DISTINCT ProjectID FROM ATC_Tasks a INNER JOIN ATC_Assignments b ON a.SubtaskID =b.SubtaskID " & _
							"INNER JOIN " & ATSsql(date(),date()-cint(intBooked)) & " c ON c.AssignmentID=b.AssignmentID WHERE Tdate<=Getdate() AND Tdate>=GetDate()-" & intBooked & ")"
	end if
	strQuery=strQuery & strQuerySearch & " ORDER BY SortDate DESC"
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

'===============================================================================
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	Dim rsProject, gMessage, PageSize
	dim varSearch,varSearchType,varBillable,varPage,varSortType,varCol,varUpdate,varBooked
	Dim strProjectID

	varSearch = trim(Request.QueryString("search"))

	varSearchType = Request.Form("lbSeachType")
	if varSearchType="" then varSearchType=3
	
	varBillable=Request.Form("lbBill")
	if varBillable="" then varBillable=2
	
	varBooked=Request.Form("lbBooked")
	if varBooked="" then varBooked=0
	
	varPage=Request.QueryString("Page")
	
	varSortType=Request.QueryString("Sorttype")
	varCol=Request.QueryString("Col")
	
	varUpdate=Request.QueryString("active")
	
	strName = Request.Form("name")

		
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
	
'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then PageSize = arrPre(1, 0) Else PageSize = PageSizeDefault
		Set arrPre = Nothing
	Else
		PageSize = PageSizeDefault
	End If

'--------------------------------------------------
' Check ACCESS right
'--------------------------------------------------

	strTemp = Request.ServerVariables("URL") 
	While Instr(strTemp, "/")<>0
		strTemp = Mid(strTemp, Instr(strTemp, "/") + 1, Len(strTemp))
	Wend
	
	strFilename = strTemp
	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = strTemp Then
				fgRight=True
				Exit For
			End If
		Next
		
		Set varGetRight = Nothing		
	End If	
	If fgRight = False Then		
		Response.Redirect("../../welcome.asp")
	End If

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
		fgApproval = False
	Else
		varGetRight = session("RightOn")
		fgApproval = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "approving project" Then
				fgApproval = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Check Registration Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRegister = False
	Else
		varGetRight = session("RightOn")
		fgRegister = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "registration" Then
				fgRegister = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Analyse query and prepare project list
'--------------------------------------------------
'varPage="" --> didn't retrive data from database
if varPage="" then	
	varPage=1
	call GetProjectData(varSearchType,varBillable,varSearch,varBooked,rsProject)	
	if rsProject.RecordCount>0 then
		if not isEmpty(session("rsProject")) then session("rsProject") = empty
		set session("rsProject")=rsProject.Clone
	else
		if not IsEmpty(Session("rsProject")) then set rsProject = session("rsProject")
		if rsProject.Recordcount>0 then rsProject.MoveFirst
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
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
	
'--------------------------------------------------
' Make list of menu
'--------------------------------------------------

	If isEmpty(session("Menu")) Then 
		getRes = getarrMenu(intUserID)
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End If	
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

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
<title>Atlas Industries Timesheet System</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--
var objNewWindow;

function next() {
var curpage = <%=varPage%>;
var numpage = <%=session("NumPage")%>;
	if (curpage < numpage) {
		document.frmreport.action = "n_projectlist.asp?page=" + (curpage + 1)
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function prev() {
var curpage = <%=varPage%>;
var numpage = <%=session("NumPage")%>;
	if (curpage > 1) {
		document.frmreport.action = "n_projectlist.asp?page=" + (curpage - 1)
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function go() {
	var numpage = <%=session("NumPage")%>;
	var curpage = <%=varPage%>;
	var intpage = document.frmreport.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmreport.action = "n_projectlist.asp?page=" + intpage
		document.frmreport.target = "_self";
		document.frmreport.submit();		
	}
}

function sort(column,type) {
	document.frmreport.action = "n_projectlist.asp?page=1&sorttype=" + type + "&Col=" + column; //1: id, 2: name, 3: Tranfer Date
	document.frmreport.target = "_self";
	document.frmreport.submit();
}

function window_onunload() 
{
	if ((objNewWindow)&&(!objNewWindow.closed))
		objNewWindow.close();
}

function search() {
	var tmp = document.frmreport.txtsearch.value;
	tmp = alltrim(tmp);
	document.frmreport.txtsearch.value = tmp;
	tmp = escape(tmp);
	
	document.frmreport.action = "n_projectlist.asp?search=" + tmp;
	document.frmreport.target = "_self";
	document.frmreport.submit();	
}

function viewpro(varid,status,tdate,managerID)
{
	window.document.frmreport.txthidden.value = varid + ";" + status + ";" + tdate+ ";" + managerID;
	window.document.frmreport.action = "project_register.asp";
	window.document.frmreport.submit();
}

function addproject()
{
	window.document.frmreport.txthidden.value="";
	window.document.frmreport.action = "project_register.asp";
	window.document.frmreport.submit();
}
function viewinvoice(varid)
{
	window.document.frmreport.txthidden.value=varid;
	window.document.frmreport.action = "pro_invoice.asp";
	window.document.frmreport.submit();
}

function window_onunload() {
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.close();
}

//-->
</script>

</head>
<body>
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
	
%>
<tr> 
<td>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
	<tr> 
		<td> 
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr bgcolor=<%if strError="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
					<td class="red" colspan="6" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td></tr>
				<tr> 
					<td class="blue-normal" width="12%" valign="middle" align="right"> &nbsp;&nbsp;Search for&nbsp; </td>
					<td align="right" width="20%" valign="middle"> 
						<input type="text" name="txtsearch" class="blue-normal" size="15" style="width:100%" value="<%=Showvalue(varSearch)%>"></td>
					<td align="center" width="20%" valign="middle" > 
						<select name="lbSeachType" class='blue-normal' style="width:95%">
							<option value="1"<%if cint(varSearchType)=1 then %>selected<%end if%>>Project Name</option>
							<option value="2"<%if cint(varSearchType)=2 then %>selected<%end if%>>APK</option>
							<option value="3"<%if cint(varSearchType)=3 then %>selected<%end if%>>Project Name &amp; APK</option>
						</select></td>
					<td align="left" width="14%" valign="middle" > 
						<select name="lbBill" class='blue-normal' style="width:100%">
							<option value="2"<%if cint(varBillable)=2 then %>selected<%end if%>>&nbsp;</option>
							<option value="1"<%if cint(varBillable)=1 then %>selected<%end if%>>Utilised</option>
							<option value="0"<%if cint(varBillable)=0 then %>selected<%end if%>>Non-Utilised</option>
						</select></td>  
					<td align="left" width="14%" valign="middle" > 
						<select name="lbBooked" class='blue-normal' style="width:100%">
							<option value="0"<%if cint(varBooked)=0 then %>selected<%end if%>>&nbsp;</option>
							<option value="7"<%if cint(varBooked)=7 then %>selected<%end if%>>Booked - 7 days</option>
							<option value="30"<%if cint(varBooked)=30 then %>selected<%end if%>>Booked - 30 days</option>
						</select></td>                 
					<td class="blue" align="left" valign="middle"> 
						<table width="100%" border="0" cellspacing="3" cellpadding="0" height="20" align="left">
							<tr> 
								<td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
									<a href="javascript:search();" class="b" onMouseOver="self.status='Search'; return true;" onMouseOut="self.status=''">Search</a></td>
							</tr>
						</table></td>
				</tr>
				<tr> 
					<td class="title" height="50" align="center" colspan="6"> List of Projects</td>
				</tr>
				<tr align="left"> 
					<td colspan="6" class="blue" height="20">&nbsp;<%If fgRegister Then%><a href="javascript:addproject()">Add New</a><%End If%></td>
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
									<table width="100%" border="0" cellspacing="1" cellpadding="5">
<%
Response.Write(OutHeader(varSortType,varCol))
Response.Write(strLast)
%>                            
									</table>
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr> 
											<td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;&nbsp;* Click on the exact project to approve or update it.</td>
											<td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">&nbsp;</td>
										</tr>
										<tr> 
											<td bgcolor="#FFFFFF" height="20" class="blue-normal" colspan="2">&nbsp;&nbsp;* Click on each column header to sort the list by alphabetical order </td>
										</tr>
										<tr> 
											<td bgcolor="#FFFFFF" class="blue-normal" colspan="2">&nbsp;</td>
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
</table></td></tr>      
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
<input type="hidden" name="txtstatus" value="<%=strStatus%>">
<input type="hidden" name="P" value="<%=intCurPage%>">
<input type="hidden" name="S" value="<%=fgSort%>">
<input type="hidden" name="name" value="<%=strSName%>">
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