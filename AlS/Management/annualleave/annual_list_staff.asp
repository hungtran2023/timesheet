<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError, intPageSize, fgRight 'view all or Not

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

	dim dblExRate
	strOut = ""
	
	If Not rsSrc.EOF Then

		For i = 1 To psize
			dblNumberOfMonths=0
			If i Mod 2 = 0 Then
				strColor = "#E7EBF5"
			Else
				strColor = "#FFF2F2"
			End If
			strExpire=""
			'if cint(rsSrc("ExpiredDay"))<>0 then
				'strExpire="Expired Date: <b>" & rsSrc("ExpiredDay") & "-" & MonthName(rsSrc("ExpiredMonth"),true) & "</b><br>Keep Pass year:<b>" & rsSrc("KeepPassYear") & " days</b>"
			'end if
			if not ISNULL(rsSrc("Lastdate")) then strApplyDate=day(rsSrc("Lastdate")) & "-" & MonthName(month(rsSrc("Lastdate")),true) & "-" & year(rsSrc("Lastdate"))
			
			dblExRate=GetExtraLeave(rsSrc("JoinDate"))
			if rsSrc("StaffID")=251 then  dblExRate=GetExtraLeave(Cdate("1-Dec-2000"))
			

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' class='blue-normal'>" & rsSrc("StaffID") & "</td>" &_
			         "<td valign='top' class='blue'><a href='javascript:viewAnnual(" & rsSrc("StaffID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;View or Annual Leave Information &quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("Fullname")) & "</a></td>" &_
			         "<td valign='top' class='blue-normal'>"& FormatNumber(dblExRate,2) &"</td>" &_
					 "<td valign='top' class='blue-normal'>"& FormatNumber(rsSrc("RatePerYear"),2) &"</td>" &_			         
			         "<td valign='top' class='blue-normal'>"& FormatNumber((dblExRate +cdbl(rsSrc("RatePerYear")))/12,2) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & strExpire & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & strApplyDate & "</td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			
			If rsSrc.EOF Then Exit For
		Next
	End If
	Outbody = strOut
End Function

'***************************************************************
'
'***************************************************************
Function GetExtraLeave(dateStart)
	
	dim dblNum,dblNumOfYears
		
	dblNum=0
	dblNumOfYears=Year(date())-Year(dateStart)

	if (dblNumOfYears mod 2=1) OR (Month(dateStart)<=month(Date())-1) then
		dblNum=(dblNumOfYears\2)
	else
		 dblNum=((dblNumOfYears-1)\2)
	end if
	
	if dblNum>=5 then dblNum=5
				
	GetExtraLeave = dblNum
end function

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	strDepartment = Request.Form("lbdepartment")
	fgSort = Request.Form("S")
	
	intCurPage = trim(Request.Form("P"))
	If intCurPage = "" Then
		intCurPage = 1
	End If		
	strName = Request.Form("name")
	intDepart = Request.Form("depart")
	
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
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = PageSizeDefault
		Set arrPre = Nothing
	Else
		intPageSize = PageSizeDefault
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
' Check VIEWALL right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "view all" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Initialize department array
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	If isEmpty(session("varDepartment")) = False Then
		varDepartment = session("varDepartment")
		intNum = Ubound(varDepartment,2)
	Else
		If objDatabase.dbConnect(strConnect) Then			
			strSQL = "SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					varDepartment = objDatabase.rsElement.GetRows
					intNum = Ubound(varDepartment,2)					
					session("varDepartment") = varDepartment
					objDatabase.closeRec
				End If
			Else
				Response.Write objDatabase.strMessage
			End If
		Else
			Response.Write objDatabase.strMessage		
		End If
	End If	

'--------------------------------------------------
' End Of initializing department array
'--------------------------------------------------

'--------------------------------------------------
' Analyse query and prepare staff list
'--------------------------------------------------

	strAct = Request.QueryString("act")
	If strAct = "" Then
		strAct = Request.Form("txtstatus")
	End If

	If strAct = "" Then					' Call this page the first
		fgSort = "N"
		
		strConnect = Application("g_strConnect")
		Set objDatabase = New clsDatabase
	
		If objDatabase.dbConnect(strConnect) Then

'--------------------------------------------------
' End of checking right on page
'--------------------------------------------------
		
			Set rsStaff = Server.CreateObject("ADODB.Recordset")
			rsStaff.CursorLocation = adUseClient			' Set the Cursor Location to Client

			Set myCmd = Server.CreateObject("ADODB.Command")
			Set myCmd.ActiveConnection = objDatabase.cnDatabase
			myCmd.CommandType = adCmdTable
			myCmd.CommandText = "HR_StaffCurrentAnnualLeave"

			On Error Resume Next	
			rsStaff.Open myCmd,,adOpenStatic,adLockBatchOptimistic
			If Err.number > 0 then
				strError = Err.Description
			End If
			
			Err.Clear
			
			If Not rsStaff.EOF Or rsStaff.RecordCount > 0 Then
				intTotalPage = pageCount(rsStaff, intPageSize)
				rsStaff.MoveFirst
				rsStaff.Move (intCurPage-1)*intPageSize
				strLast = Outbody(rsStaff, intPageSize)

				Set session("rsStaff") = rsStaff
			End if
			Set myCmd = Nothing
		Else
			strError = objDatabase.strMessage
		End If
		Set objDatabase = Nothing
		
		
		
	Else															' Submit this page
	
		Set rsStaff = session("rsStaff")
		rsStaff.MoveFirst
		If recCount(rsStaff) >= 0 Then
			intTotalPage = pageCount(rsStaff, intPageSize)
		
			Select Case strAct
				Case "vpsn"											' Sort by fullname
					
					strStatus = strAct

'--------------------------------------------------
' This If..Then..End If to check status
' of the form when it go back					
'--------------------------------------------------

					If Request.QueryString("b") <> "" Then
						If fgSort = "A" Then
							fgSort = "D"
						ElseIf fgsort = "D" Then
							fgSort = "A"
						End If
					End If
					
'--------------------------------------------------
' End of checking		
'--------------------------------------------------								
								
					If fgSort = "N" Or fgSort = "D" Then
						rsStaff.Sort = "FullName ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "FullName DESC"
						fgSort = "D"				
					End If

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vpst"											' Sort by job title

					strStatus = strAct

'--------------------------------------------------
' This If..Then..End If to check status
' of the form when it go back					
'--------------------------------------------------

					If Request.QueryString("b") <> "" Then
						If fgSort = "A" Then
							fgSort = "D"
						ElseIf fgsort = "D" Then
							fgSort = "A"
						End If
					End If
					
'--------------------------------------------------
' End of checking		
'--------------------------------------------------								

					If fgSort = "N" Or fgSort = "D" Then
						rsStaff.Sort = "JobTitle ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "JobTitle DESC"
						fgSort = "D"				
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				
				Case "vpa1"											' When user click button "Go"
					If CInt(Request.Form("txtpage")) <= CInt(intTotalPage) Then
						intCurPage = Request.Form("txtpage")
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
				Case "vpa2"											' When user click Previous link	
					If CInt(intCurPage) > 1 Then
						intCurPage = CInt(intCurPage) - 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
				Case "vpa3"											' When user click Next link		
					If CInt(intCurPage) < CInt(intTotalPage) Then
						intCurPage = CInt(intCurPage) + 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
				Case "vra1"											' When user click button "Search"
					strSName = Request.Form("txtname")
					'strDepartment=""
					strDepartment = Request.Form("lbdepartment")
					
		
					If strSName <> "" And strDepartment <> "" Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%' AND Department like '" & strDepartment & "%'"
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "# AND AND Department like #" & strDepartment & "#"
						End If		
					ElseIf strSName = "" And strDepartment <> "" Then
						rsStaff.Filter = "Department like '" & strDepartment & "%'"
					ElseIf strSName <> "" And strDepartment = "" Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%'"
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "#"
						End If	
					End If

					If Not rsStaff.EOF Or rsStaff.RecordCount > 0 Then
						intCurPage = 1
						intTotalPage = pageCount(rsStaff, intPageSize)

						rsStaff.MoveFirst
						rsStaff.Move CInt((intCurPage-1)*intPageSize)
					Else
						strError = "No data for your request."
						rsStaff.Filter = ""
					End If	

					strStatus = Request.Form("txtstatus")
				Case "vra2"											' When user click button "Show all"
					rsStaff.Filter = ""
					intTotalPage = pageCount(rsStaff, intPageSize)

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					strDepartment = ""
					strName = ""
					
					strStatus = Request.Form("txtstatus")
					
			End Select	 

		strLast = Outbody(rsStaff, intPageSize)
			
		End If		
	End If
	
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
	If strChoseMenu = "" Then strChoseMenu = "AE"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

function sort(kind)
{
	document.frmreport.action = "annual_list_staff.asp?act=vps" + kind;
	document.frmreport.submit();
}

function viewpage(kind)
{
	var intpage = parseInt(window.document.frmreport.txtpage.value,10);
	var curpage = "<%=CInt(intCurPage)%>";
	var pagetotal = "<%=CInt(intTotalPage)%>";
	
	if (kind == 1)
	{
		window.document.frmreport.txtpage.value = intpage
		if ((intpage > 0) & (intpage <= pagetotal) & (intpage != curpage)) 
		{
			document.frmreport.action = "annual_list_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "annual_list_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function search(kind)
{
	if (kind == "1")
	{
		if ((document.frmreport.txtname.value != "") || (document.frmreport.lbdepartment.options[document.frmreport.lbdepartment.selectedIndex].value != "0"))
		{
			document.frmreport.action = "annual_list_staff.asp?act=vra" + kind;
			document.frmreport.submit();	
		}
	}	
	else
	{
		document.frmreport.action = "annual_list_staff.asp?act=vra" + kind;
		document.frmreport.submit();
	}	
}

function viewAnnual(varid)
{
	window.document.frmreport.txthidden.value = varid;
	window.document.frmreport.action = "annual_detail_staff.asp";
	window.document.frmreport.submit();

}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<%	If strError1 = "" Then%>        
		  <tr> 
		    <td> 
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%If strError <> "" Then%>		      
					<tr bgcolor="#E7EBF5">
						<td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
					</tr>
<%End If%>		        
					<tr align="center"> 
						<td class="blue-normal" height="30" align="right" width="10%">&nbsp;&nbsp;Name&nbsp;&nbsp;&nbsp;</td>
						<td class="blue" height="30" align="left" width="25%"> 
							<input type="text" name="txtname" value="<%=showvalue(strSName)%>" class="blue-normal" size="15" style=" width:100%">
						</td>
						<td class="blue-normal" height="30" align="right" width="5%">Department&nbsp;&nbsp;&nbsp;</td>
						<td class="blue" height="30" align="left" width="25%"> 
							<select id="lbdepartment" size="1" name="lbdepartment" style="width:100%" class="blue-normal">
								  <option value="" selected>&nbsp;</option>
<%
If intNum >= 0 Then
	For ii = 0 To intNum
%>                    
								  <option <%If strDepartment=varDepartment(1,ii)Then%> selected <%End If%> value="<%=showlabel(trim(varDepartment(1,ii)))%>"><%=showlabel(varDepartment(1,ii))%></option>
<%
	Next
End If	
%>
							</select>
						</td>
						<td class="blue-normal" height="30" align="left" width="25%"> 
							<table width="120" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
								<tr> 
									<td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
										<div align="center" class="blue"><a href="javascript:search('1');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Search</a></div>
									</td>
									<td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
										<a href="javascript:search('2');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Show All</a>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<tr align="center"> 
				  <td class="title" height="50" align="center" colspan="5">Staff Annual Leave Information</td>
				</tr>
				</table>
		    </td>
		  </tr>
		  <tr> 
		    <td height="100%"> 
		      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:'79%'" height="365">
		        <tr> 
		          <td bgcolor="#FFFFFF" valign="top"> 
		            <table width="100%" border="0" cellspacing="0" cellpadding="0">
		              <tr> 
		                <td bgcolor="#617DC0"> 
		                  <table width="100%" border="0" cellspacing="1" cellpadding="5">
		                    <tr bgcolor="#8CA0D1"> 
		                      <td class="blue" align="center" width="5%"></td>
		                      <td class="blue" align="center" width="25%"><a href="javascript:sort('n')" onMouseOver="self.status='Sort by Full Name';return true" onMouseOut="self.status='';return true" class="c">Full Name</a></td>
		                      <td class="blue" align="center" width="11%">Extra <br> leave</td>
		                      <td class="blue" align="center" width="11%">Rate <br>for level</td>
		                      <td class="blue" align="center" width="12%">Rate <br>per month</td>
		                      <td class="blue" align="center" width="22%">Expired</td>
		                      <td class="blue" align="center" width="14%">Apply Date</td>
		                    </tr>
<%
	Response.Write(strLast)
%>                            
		                  </table>
		                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		                    <tr> 
		                      <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;&nbsp;* Click 
		                        on the exact name to view or update Annual Leave Information.</td>
		                      <td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">&nbsp;</td>
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
          <tr> 
            <td> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
                <tr> 
                  <td align="right" bgcolor="#E7EBF5"> 
                    <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
                      <tr class="black-normal"> 
                        <td align="right" valign="middle" width="27%" class="blue-normal">Page 
                        </td>
                        <td align="center" valign="middle" width="13%" class="blue-normal"> 
                          <input type="text" name="txtpage" class="blue-normal" value="<%=intCurPage%>" size="2" style="width:50">
                        </td>
                        <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:viewpage(1);" onMouseOver="self.status='';return true"><font color="#990000">Go</font></a></td>
						<td align="right" valign="middle" width="25%" class="blue-normal"><%If CInt(intTotalPage) <> 0 Or intTotalPage <> "" Then%>Pages <%=intCurPage%>/<%=intTotalPage%><%End If%>&nbsp;&nbsp;</td>
						<td valign="middle" align="right" width="28%" class="blue-normal"><%If CInt(intCurPage) <> 1 Then%><a href="javascript:viewpage(2);" onMouseOver="self.status='Move Previous';return true" onMouseOut="self.status='';return true">Previous</a><%End If%><%If CInt(intCurPage) <> 1 And  CInt(intCurPage) <> CInt(intTotalPage) Then%>/<%End If%><%If CInt(intCurPage) <> CInt(intTotalPage) And (CInt(intTotalPage) <> 0 Or intTotalPage <> "") Then%><a href="javascript:viewpage(3);" onMouseOver="self.status='Move Next';return true" onMouseOut="self.status='';return true"> Next</a><%End If%>&nbsp;&nbsp;&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
<%	Else          
		If strError <> "" Then
%>               
				<tr>
				  <td class="red">&nbsp;<%=strError%></td>
				</tr>
<%		End If%>				

		  <tr>
         	<td class="red" align="center" valign="middle"><b><%=strError1%></b></td>
		  </tr>	          

<%	End If
	Set objDatabase = Nothing
%>
          
        </table>
      
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
<input type="hidden" name="depart" value="<%=intDepart%>">

</form>
</body>
</html>