<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	'****************************************
' Function: Outbody1
' Description: holiday
' Parameters: source recordset
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Outbody1(ByRef rsSrc)
	dim inYear
	strOut = ""
	i = 0
	inYear=0
	strColor = "#FFF2F2"
	do until rsSrc.EOF
		
		'if i mod 2 = 0 then	strColor = "#E7EBF5"
		
		if isnull(rsSrc("YearAN")) then Exit do
		if rsSrc("YearAN")<>inYear then
			if strColor = "#FFF2F2" then 
				strColor="#d2daec"
			else
				strColor = "#FFF2F2" 
			end if
			
		end if
	
				
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" &_
				"<td valign='top' class='blue' align='center'>" & rsSrc("YearAN") & "</td>" &_
				"<td valign='top' class='blue-normal' align='center'>" & day(rsSrc("DateFrom")) & "/" & month(rsSrc("DateFrom")) & "/" & Year(rsSrc("DateFrom"))   & "</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:center;'>"& day(rsSrc("DateTo")) & "/" & month(rsSrc("DateTo")) & "/" & Year(rsSrc("DateTo")) &  "</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:center;'>"& rsSrc("MonthstoCalLeavedue") &"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:center;'>"& rsSrc("WorkingHours") &"</td>" &_
				"<td valign='top' class='blue-normal'style='text-align:center;'>"&rsSrc("RateperYear") &"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:center;'>"&rsSrc("RateByYTD")&"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:right; padding-right:10px'>"& formatnumber((cdbl(rsSrc("RateperYear"))+Cdbl(rsSrc("RateByYTD")))/12,2) &"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:center;'> " & IIF(cdbl(rsSrc("KeepPassYear"))=0,"",rsSrc("KeepPassYear")) & " </td>" &_
				"<td valign='top' class='blue-normal' style='text-align:center;'>"&IIF(cdbl(rsSrc("MoreHours"))=0,"",rsSrc("MoreHours"))&"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:right; padding-right:10px'>"&formatnumber(rsSrc("Leavedue"),3)&"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:right; padding-right:10px'>"&formatnumber(rsSrc("ApplicationBy"),3)&"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:right; padding-right:10px'>"&formatnumber(rsSrc("BeforeExpired"),3)&"</td>" &_
				"<td valign='top' class='blue-normal' style='text-align:right; padding-right:10px'>"&formatnumber(rsSrc("AfterExpired"),3)&"</td>" &_
				"</tr>"
		if rsSrc("ANType")="Reserved" then	dblBalance=cdbl(rsSrc("AfterExpired"))/cdbl(rsSrc("WorkingHours"))
		inYear =rsSrc("YearAN")
		rsSrc.MoveNext
		'i = i + 1
	loop
	Outbody1 = strOut
end function
'==================================================================
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage,gErrMessage
	
	Dim strAct,strStatus, intStaffID
	dim dblBalance
	
	gMessage=""
	

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
%>	
<script LANGUAGE="javascript">
<!--
	opener.document.location = "../../message.htm";
	window.close();
//-->
</script>
<%
	End If	

	intUserID	= session("USERID")
	intStaffID  = Request.Form("txthidden")

	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "GetDurationAnnualLeave_2018"
		
		Set myParam = myCmd.CreateParameter("staffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("dateTo",adDate,adParamInput)
		myCmd.Parameters.Append myParam

		myCmd("staffID") = intStaffID
		myCmd("dateTo") = Date()
		set rsTemp=myCmd.Execute		
		
	end if
	if isnull(rsTemp) then
		Response.write ("Wrong")
	else
		strList=Outbody1(rsTemp)
	End if	
	
	

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	intDepartID = varFullName(2)

	Set objEmployee = Nothing


	
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------
	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:goback();' onMouseOver='self.status=&quot;Back&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Back</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view annual leave&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "<b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../../templates/template1/")

%>	
<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<script language="javascript" src="../../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objNewWindow;
ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	URL = "../../logout.asp";
	if (ns)
		document.location = URL;
	else
	{
		window.document.frmtms.action = URL;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}


function goback()
{
	if (ns)
		document.location =  "staff_view_leave.asp";
	else
	{
		window.document.frmtms.action = "staff_view_leave.asp";
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function selstaff()
{
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=490,height=325,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('tms_select_staff.asp?view=l', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function back_menu()
{
	window.document.frmtms.action = "rpt_list_staff.asp?act=vra1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}
function ANTRacking()
{
	window.document.frmtms.action = "CalBalanceByUser.asp?b=1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmtms" method="post">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6" background="../../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
    <td valign="top" height="100%" align="center">
		<table  width="80%" border="0" cellspacing="1" cellpadding="0" align="center" style="height:79%" height="365">
<%If strError <> "" Then%>
			<tr>
				<td height="80"> 
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr bgcolor="#E7EBF5"> <td class="blue" align="center"><%=strError%></td></tr>
					</table>
				</td>	
			</tr>
<%End If%>			    
			<tr> 
				<td height="80"> 
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr><td class="title" align="center">Annual Leave Tracking</td></tr>
						<tr><td class="blue-normal" align="center" bgcolor="#FFFFFF" height="20"><%=strTitle1%></td></tr>			  
					</table>
				</td>
			</tr>			
			<tr>
				<td valign="top"> 

					<table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
						<tr> 
							<td bgcolor="#FFFFFF" valign="top"> 
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr> 
										<td bgcolor="#c2cce7"> 
											<table width="100%" border="0" cellspacing="1" cellpadding="4">
												<tr bgcolor="#8CA0D1"> 
													<td class="blue"  align="center" colspan="1" rowSpan="2"  width="6%">Year</td>
													<td class="blue" align="center" colspan="3" >Duration</td>								
													<td class="blue" align="center" colspan="1" rowSpan="2" width="7%">Working Hrs<br>(2)</td>								
													<td class="blue" align="center" colspan="3"  >Total per year (days)</td>								
													<td class="blue" align="center" colspan="2">Expired Information</td>
													<td class="blue" align="center" colspan="1" rowSpan="2" width="8%">Total Leave (hrs)<br>(5)=(1)*(2)*(3)+(4)</td>
													<td class="blue" align="center" colspan="1" rowSpan="2" width="7%">Hours taken<br>(hrs)<br>(6)</td>
													<td class="blue" align="center" colspan="2">Balance (hrs)</td>													  
												</tr>
												<tr bgcolor="#8CA0D1"> 								
													<td class="blue" align="center" width="10%">From</td>
													<td class="blue" align="center" width="10%">To</td>
													<td class="blue" align="center" width="7%">Months To Cal.<br> (1)</td>
													<td class="blue" align="center" width="7%">Days/year<br> (3a)</td>
													<td class="blue" align="center" width="7%">Rate YTD<br> (3b)</td>
													<td class="blue" align="center" width="6%">Accrual rate/month <br> (3)=(3a+3b)/12</td>
													<td class="blue" align="center" width="6%">Bring to next year<br> (days)</td>
													<td class="blue" align="center" width="6%">More hours(4)</td>								
													<td class="blue" align="center" width="7%">Before Expired (7)=(8)+(4)+(5)-(6)</td>
													<td class="blue" align="center" width="6%">After Expired (8)</td>							  
												</tr>
												<%Response.Write strList%>
												<tr bgcolor="#8CA0D1"> 
													<td class="blue"  colspan="13" align="right">Balance</td>
													<td class="blue" align="center" colspan="2"><%=formatnumber(dblBalance,2)%>(days)</td>													  
												</tr>
											</table>
											<table width="100%" border="0" cellspacing="0" cellpadding="0">
												<tr> 
													<td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%" align="right"></td>
													<td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="left">&nbsp;</td>
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
		</table>
	
    
    </td>
    <td width="2" background="../../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
  </tr>
</table>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>

<input type="hidden" name="txthidden" value="<%=intStaffID%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">

</form>
</body>
</html>
