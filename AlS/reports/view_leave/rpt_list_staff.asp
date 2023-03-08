<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%

	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strError, PageSize, fgRight 'view all or Not

	
	Dim intUserID, intMonth, intYear, dblCurLeave
	Dim strConnect, objDatabase,rsAnnual


'**********************************************************

'**********************************************************
Sub GetAnnualLeaveOverview()


	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase

	If objDatabase.dbConnect(strConnect) Then

			
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "AnnualLeaveOverview"

	
		set rsAnnual=myCmd.Execute
	
		'response.write rsAnnual.recordcount
		set Session("rsAnnual")=rsAnnual
	end if	
end sub

'***************************************************************
'
'***************************************************************
Function GetApplication(intStaffID,dateF,DateT)

	dim strSql,strTable
	dim rs, dblApplication

	dblApplication=0
'(year(Date)+1,intstaffID,cdate("1-Jan-" & year(date)+1),cdate("31-Dec-" & year(date)+1))

	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "StaffApplication"
		Set myParam = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("datefromIn",adDate,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("datetoIn",adDate,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("applicationDatesOut",adVarChar,adParamOutput,10)
		myCmd.Parameters.Append myParam

		myCmd("StaffID") = intStaffID
		myCmd("datefromIn")=dateF
		myCmd("datetoIn")=dateT

		myCmd.Execute
		dblApplication=myCmd("applicationDatesOut")
	end if

	GetApplication = cdbl(dblApplication)
end function

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

Function Outbody(Byval rsSrc)
	
	Dim dblRateperYear, dblBalance
	Dim dblBalanceToDate, dblANReserved

	strOut = ""
	i=0
	If Not rsSrc.EOF Then
		do while  Not rsSrc.EOF
			If i Mod 2 = 0 Then
				strColor = "#E7EBF5"
			Else
				strColor = "#FFF2F2"
			End If

			dblRateperYear=cdbl(rsSrc("ratePerYear"))+Cdbl(rsSrc("numberofyear"))/2
			dblBalance=formatnumber(cdbl(rsSrc("balance"))/cdbl(rsSrc("WorkingHours")),2)
			dblBalanceToDate=formatnumber(cdbl(rsSrc("balanceToDate"))/cdbl(rsSrc("WorkingHours")),2)

			dblANReserved=GetApplication(cdbl(rsSrc("StaffID")),Date(),cdate("31-Dec-" & year(date)+1))
			dblANReserved=formatnumber(cdbl(dblANReserved)/cdbl(rsSrc("WorkingHours")),2)

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
					"<td valign='top' class='blue'>" & Showlabel(rsSrc("HRCode")) & "</a></td>" &_
			         "<td valign='top' class='blue'>" & Showlabel(rsSrc("Fullname")) & "</a></td>" &_
			         "<td valign='top' class='blue-normal'>" & formatnumber(dblRateperYear/12,2) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & dblBalanceToDate & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & dblANReserved & "</td>" &_
			          "<td valign='top' class='blue-normal'>" & dblBalance & "</td>" &_
			         "</tr>" & chr(13)
			         
			rsSrc.MoveNext
			i=i+1
		Loop
	End If
	Outbody = strOut
End Function

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intDepartmentID = Request.Form("lbdepartment")
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
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = 12'PageSizeDefault
		Set arrPre = Nothing
	Else
		intPageSize = 12'PageSizeDefault
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
		Call GetAnnualLeaveOverview()
	Else															' Submit this page
	
		'Set rsAnnual = Session("rsAnnual")
		'rsAnnual.MoveFirst
		
	End If
	
	strLast = Outbody(rsAnnual)
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
				Help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
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
	If strChoseMenu = "" Then strChoseMenu = "B"
	
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
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--
var objPrintWindow;

function sort(kind)
{
	document.frmreport.action = "rpt_list_staff.asp?act=vps" + kind;
	document.frmreport.submit();
}

function viewpage(kind)
{
	window.document.frmreport.txtpage.value = intpage
	var intpage = parseInt(window.document.frmreport.txtpage.value,10);
	var curpage = "<%=CInt(intCurPage)%>";
	var pagetotal = "<%=CInt(intTotalPage)%>";
	
	if (kind == 1)
	{
		if ((intpage > 0) & (intpage <= pagetotal) & (intpage != curpage)) 
		{
			document.frmreport.action = "rpt_list_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "rpt_list_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function search(kind)
{
	if (kind == "1")
	{
		if ((document.frmreport.txtname.value != "") || (document.frmreport.lbdepartment.options[document.frmreport.lbdepartment.selectedIndex].value != "0"))
		{
			document.frmreport.action = "rpt_list_staff.asp?act=vra" + kind;
			document.frmreport.submit();	
		}
	}	
	else
	{
		document.frmreport.action = "rpt_list_staff.asp?act=vra" + kind;
		document.frmreport.submit();
	}	
}

function viewtms(varid)
{
	window.document.frmreport.txthidden.value = varid;
	window.document.frmreport.action = "rpt_view_leave.asp";
	window.document.frmreport.submit();

}

function printall()
{
	window.status = "";
	
	strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=450,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";

	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();	

	objPrintWindow = window.open('rpt_print_leave.asp', "MyNewWindow", strFeatures);
	objPrintWindow.focus();

	window.status = "Opened a new browser window.";  
}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" Topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" >
		  <tr> 
		    <td> 
		      <table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
		        <tr <%If strError="" Then%> bgcolor="#FFFFFF" <%Else%> bgcolor="#E7EBF5" <%End If%>>
				  <td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
		        </tr>
		        <tr align="center"> 
		          <td class="blue-normal" height="30" align="right" width="12%">&nbsp;&nbsp;Search for Name&nbsp;</td>
  				  <td class="blue" height="30" align="left" width="30%"> 
				    <input type="text" name="txtname" value="<%=showvalue(strSName)%>" class="blue-normal" size="15" style=" width:100%">
				  </td>
				  <td class="blue-normal" height="30" align="right" width="6%">Department&nbsp;</td>
				  <td class="blue" height="30" align="left" width="20%"> 
				    <select id="lbdepartment" size="1" name="lbdepartment" class="blue-normal" style="width:100%">
					  <option value="0" selected>&nbsp;</option>
<%
		If intNum >= 0 Then
			For ii = 0 To intNum
%>                    
					  <option <%If CInt(intDepartmentID)=CInt(varDepartment(0,ii)) Then%> selected <%End If%> value="<%=varDepartment(0,ii)%>"><%=showlabel(varDepartment(1,ii))%></option>
<%
			Next
		End If	
%>					

				    </select>
				  </td>
				  <td class="blue-normal" height="30" align="left" width="32%"> 
				    <table width="200" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
					  <tr> 
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
						  <div align="center" class="blue"><a href="javascript:search('1');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Search</a></div>
					    </td>
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
						  <a href="javascript:search('2');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Show All</a>
					    </td>
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
						  <a href="javascript:printall();" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Print All</a>
					    </td>
					  </tr>
				    </table>
				  </td>
			    </tr>
			    <tr align="center"> 
				  <td class="title" height="50" align="center" colspan="5">Annual Leave Overview</td>
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
		                      <td class="blue" align="center" width="15%">StaffID</td>
		                      <td class="blue" align="center" width="25%">Full Name</td>
		                      <td class="blue" align="center" width="15%">Accrual rate/month</td>
		                       <td class="blue" align="center" width="15%">AN available to date (days)</td>
		                      <td class="blue" align="center" width="15%">AN Reserved (days)</td>
		                      <td class="blue" align="center" width="15%">Balance</td>
		                    </tr>
<%
	Response.Write(strLast)
%>                            
		                  </table>
		                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		                    <tr> 
		                      <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;&nbsp;* Click 
		                        on the exact name to view timesheet.</td>
		                      <td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">&nbsp;</td>
		                    </tr>
		                    <tr> 
		                      <td bgcolor="#FFFFFF" height="20" class="blue-normal" colspan="2">&nbsp;&nbsp;* Click 
		                         on each column header to sort the list by alphabetical order </td>
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