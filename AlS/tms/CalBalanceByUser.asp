<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/library.asp"-->
<!-- #include file = "../inc/constants.inc"-->
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
' Initialize variables
'--------------------------------------------------

    strTmp=Request.ServerVariables("URL")

	strTmp = Mid(strTmp , 1, Instr(2, strTmp, "/")-1)
	strHTTP = "http://" & Request.ServerVariables("SERVER_NAME") & strTmp & "/"
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
%>
<script language="javascript">
<!--
    opener.document.location = "../message.htm";
    window.close();
    //-->
</script>
<%
	End If

	intUserID	= session("USERID")
	intStaffID  = intUserID
	
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
	strFunction = "<a class='c' href='../welcome.asp' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:window.history.back();' onMouseOver='self.status=&quot;Back&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Back</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "<b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../templates/template1/")

'strError="The system is being upgraded. Please try to log-in again after 12:00 AM"

%>
<html>
<head>
    <meta http-equiv="PRAGMA" content="NO-CACHE">

    <title>Atlas Industries - Timesheet</title>

    <link rel="stylesheet" href="../timesheet.css">

    
	<script language="javascript" src="../library/library.js"></script>

<script language="JavaScript">
<!--
    var ns, ie, objNewWindow;
    ns = (document.layers) ? true : false;
    ie = (document.all) ? true : false;
	
	
    function logout() {
		var URL;
        URL = "../logout.asp";
		window.document.frmtms.action = URL;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
        
    }


    function goback() {
       
		window.document.frmtms.action = "timesheet.asp?act=vpa";
		//		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
        
    }

       function back_menu() {
        window.document.frmtms.action = "../welcome.asp";
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
                <td width="6" background="../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
                <td valign="top" height="100%" align="center">
                    <table width="80%" border="0" cellspacing="1" cellpadding="0" style="height: 79%" height="365">
                        <%If strError <> "" Then%>
                        <tr>
                            <td height="80">
                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                    <tr bgcolor="#E7EBF5">
                                        <td class="blue" align="center"><%=strError%></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <%else%>
                        <tr>
                            <td height="80">
                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td class="title" align="center">Annual Leave Tracking</td>
                                    </tr>
                                    <tr>
                                        <td class="blue-normal" align="center" bgcolor="#FFFFFF" height="20"><%=strTitle1%></td>
                                    </tr>
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
                        <%End If%>
                    </table>
                </td>
                <td width="2" background="../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
            </tr>
        </table>
        <%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
        %>
        <input type="hidden" name="M" value="<%=intMonth%>">
        <input type="hidden" name="Y" value="<%=intYear%>">
        <input type="hidden" name="P" value="<%=Request.Form("P")%>">
        <input type="hidden" name="S" value="<%=Request.Form("S")%>">
        <input type="hidden" name="txthidden" value="<%=intStaffID%>">
        <input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
    </form>
</body>
</html>
