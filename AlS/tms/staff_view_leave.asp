<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/library.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<%

	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	Dim intUserID, intMonth, intYear, dblCurLeave
	Dim strConnect, objDatabase, strError
	Dim rsDuration,rsIndividualRule
	Dim dateTo
	Dim rsANUser 
	
	Dim dblBalance,dblApplication,dblLeaveDue,dblRatePerMonth
	dim dblLeaveDueThisYear,dblMoreHoursThisYear,dblApplicationReserve
	Dim dblBalanceLastYear,dblBalanceByDays,dblWorkingHour
	

    Dim SessionSharing
    Set SessionSharing = server.CreateObject("SessionMgr.Session2")
	
'***************************************************************
'
'***************************************************************
Sub GetANUser(intStaffID)

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
		set rsANUser=myCmd.Execute		
		
	end if
	
end sub	

'***************************************************************
'
'***************************************************************
Function GetApplication(intYear,staffID,dateF,DateT)

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

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------	
	
	dateTo=Date()

	intMonth = Request.Form("M")
	intYear = Request.Form("Y")
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


	Set objEmployee = New clsEmployee

	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	intDepartID = varFullName(2)

	Set objEmployee = Nothing
	
	dblRatePerMonth=0
	dblBalanceByDays=0
	dblBalanceLastYear=0
	dblLeaveDueThisYear=0
	dblMoreHoursThisYear=0
	dblApplicationThisYear=0
	dblApplicationReserve=0
	dblBalance=0
	
	Call GetANUser(intStaffID)

	if isnull(rsANUser) then
		Response.write ("Wrong")
	else
		'strHistoryRate=DataAnalysis(rsANUser) 		
	
		dim strOut,intNumYears,intRateperYear
	
		intNumYears=0
		intRateperYear=0
		
		rsANUser.MoveFirst

		if Not rsANUser.EOF OR rsANUser.recordcount>0 then

			if rsANUser("ANType")<>"Reserved" then 
				totalperYear=(cdbl(rsANUser("RateByYTD"))+ cdbl(rsANUser("RateperYear")))/12
	
				strHistoryRate=strHistoryRate & "<tr bgcolor='#E7EBF5' height='25' > " & _
							"<td class='blue-normal'>&nbsp;&nbsp;" & day(rsANUser("DateFrom")) & "-" & MonthName(month(rsANUser("DateFrom")),true) & "-" & year(rsANUser("DateFrom")) & "</td>" & _
							"<td class='blue-normal' align='center'>" & rsANUser("NumberOfYear") & "</td>" & _
							"<td class='blue-normal' align='center'>" & rsANUser("RateByYTD") & "</td>" & _
							"<td class='blue-normal' align='center' >" & rsANUser("RateperYear") & "</td>" & _
							"<td class='blue-normal' align='center'><b>" & cdbl(rsANUser("RateByYTD"))+ cdbl(rsANUser("RateperYear")) & "<b></td>" & _
							"<td class='red' align='center'><b>" & FormatNumber(totalperYear,2) & "</b></td>" & _
							"<td class='blue-normal'></td></tr>"
				intNumYears=cdbl(rsANUser("NumberOfYear"))
				intRateperYear=cdbl(rsANUser("RateperYear"))								

				do while not rsANUser.EOF
					if cdbl(rsANUser("NumberOfYear"))<>intNumYears OR cdbl(rsANUser("RateperYear"))<>intRateperYear then
						intNumYears=cdbl(rsANUser("NumberOfYear"))
						intRateperYear=cdbl(rsANUser("RateperYear"))
						totalperYear=(cdbl(rsANUser("RateByYTD"))+ cdbl(rsANUser("RateperYear")))/12
						
						strHistoryRate=strHistoryRate & "<tr bgcolor='#E7EBF5' height='25' > " & _
						"<td class='blue-normal'>&nbsp;&nbsp;" & day(rsANUser("DateFrom")) & "-" & MonthName(month(rsANUser("DateFrom")),true) & "-" & year(rsANUser("DateFrom")) & "</td>" & _
						"<td class='blue-normal' align='center'>" & rsANUser("NumberOfYear") & "</td>" & _
						"<td class='blue-normal' align='center'>" & rsANUser("RateByYTD") & "</td>" & _
						"<td class='blue-normal' align='center' >" & rsANUser("RateperYear") & "</td>" & _
						"<td class='blue-normal' align='center'><b>" & cdbl(rsANUser("RateByYTD"))+ cdbl(rsANUser("RateperYear")) & "<b></td>" & _
						"<td class='red' align='center'><b>" & FormatNumber(totalperYear,2) & "</b></td>" & _
						"<td class='blue-normal'>" & intNumYears & " year(s)</td></tr>"
						
					end if	

					if rsANUser("YearAN")=Year(Date())-1 AND rsANUser("ANType")="End" then dblBalanceLastYear=cdbl(rsANUser("AfterExpired"))
	'response.write rsANUser("YearAN") & "<br>"				
					if rsANUser("YearAN")=Year(Date()) then
						dblLeaveDueThisYear=dblLeaveDueThisYear + cdbl(rsANUser("Leavedue"))
						dblMoreHoursThisYear=dblMoreHoursThisYear + cdbl(rsANUser("MoreHours"))
						if rsANUser("ANType")<>"Reserved" then 
							dblApplicationThisYear=dblApplicationThisYear+cdbl(rsANUser("ApplicationBy"))
						else
							dblApplicationReserve=cdbl(rsANUser("ApplicationBy"))
							dblWorkingHour=cdbl(rsANUser("WorkingHours"))
						end if
					end if
					
					dblBalance=cdbl(rsANUser("AfterExpired"))
					
					rsANUser.MoveNext
					
				loop		
				dblRatePerMonth=totalperYear
				if year(date())<>2022 AND year(date())<>2023 then
					dblApplicationThisYear= cdbl(dblApplicationThisYear)-cdbl(dblBalanceLastYear)
				end if
				if dblApplicationThisYear<0 then dblApplicationThisYear=0
		
				dblBalanceByDays=cdbl(dblBalance)/cdbl(dblWorkingHour)
			end if

		end if
	End if	
	


	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	
	strSql="SELECT * FROM ATC_AnnualLeaveIndividualRule WHERE StaffID="	& intStaffID & " ORDER BY ApplyYear"	
	Call GetRecordset(strSql,rsIndividualRule)	

	strSql="SELECT ExpiredDay, ExpiredMonth FROM ATC_EmployeeExpiredRule a INNER JOIN ATC_AnnualLeaveYearlyRule b " & _
					"ON a.RuleYearlyID=b.RuleYearlyID WHERE staffID=" & intStaffID & " AND applyYear IN " & _
						"(SELECT MAX(ApplyYear) FROM ATC_EmployeeExpiredRule WHERE staffID=" & intStaffID & " AND ApplyYear<=Year(getdate()))"
			
	Call GetRecordset(strSql,rsExpireday)

	expiredDate=null
	if not rsExpireday.EOF then
		expiredDate=cdate(rsExpireday("ExpiredMonth") & "/" & rsExpireday("ExpiredDay") & "/" &  Year(date))
	end if	
	expiredDateThisYear=expiredDate
		'if intStaffID=1315 then Response.end

	if not rsIndividualRule.EOF then				
		rsIndividualRule.MoveFirst
		rsIndividualRule.Filter="ApplyYear=" & year(Date())		
		if not rsIndividualRule.Eof then 
			expiredDateThisYear=cdate(rsIndividualRule("ExpiredMonth") & "/" & rsIndividualRule("ExpiredDay") & "/" &  rsIndividualRule("ApplyYear"))
		end if
		rsIndividualRule.Filter=""
	end if


	
    SessionSharing("CURRENT_RATE") = formatnumber(dblRatePerMonth,2)
    SessionSharing("BALANCE_DAY") = formatNumber(dblBalanceByDays,2)
    SessionSharing("BALANCE_LAST_YEAR") = formatNumber(dblBalanceLastYear, 2)
	
    SessionSharing("LEAVE_UNTIL_DAY") = formatnumber(dblLeaveDueThisYear-dblMoreHoursThisYear,2)
    SessionSharing("TOTAL_HOURS") = formatnumber(dblLeaveDueThisYear + dblBalanceLastYear ,2)
    SessionSharing("ANUAL_LEAVE_CURRENT_YEAR") = formatNumber(dblApplicationThisYear,2)
    SessionSharing("ANUAL_LEAVE_RESERVED") = formatNumber(dblApplicationReserve,2)
    SessionSharing("BALANCE_HOURS") = FormatNumber(dblBalance,2)
	
	
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

    <style type="text/css">
        .text-right{
            text-align: right;
        }
        .btn-custom{
            background-color: #88a5ea;
            color: #ffffff !important;
            display: inline-block;
            padding: 6px 12px;
            margin-bottom: 0;
            font-size: 14px;
            font-weight: normal;
            line-height: 1.428571429;
            text-align: center;
            white-space: nowrap;
            vertical-align: middle;
            cursor: pointer;
            border: 1px solid transparent;
            border-radius: 4px;
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            -o-user-select: none;
            user-select: none;
            text-decoration: none !important;
            border-color: #357ebd;
            margin-left: 10px;
        }
    </style>

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

    function selstaff() {
        window.status = "";

        strFeatures = "top=" + (screen.height / 2 - 225) + ",left=" + (screen.width / 2 - 230) + ",width=490,height=325,toolbar=no,"
                  + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";

        if ((objNewWindow) && (!objNewWindow.closed))
            objNewWindow.focus();
        else {
            objNewWindow = window.open('tms_select_staff.asp?view=l', "MyNewWindow", strFeatures);
        }
        window.status = "Opened a new browser window.";
    }

    function back_menu() {
        window.document.frmtms.action = "../welcome.asp";
        window.document.frmtms.target = "_self";
        window.document.frmtms.submit();
    }
	
	function ANTracking()
{
	window.document.frmtms.action = "CalBalanceByUser.asp";
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
                    <table width="772" border="0" cellspacing="1" cellpadding="0" style="height: 79%" height="365">
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
                                        <td class="title" align="center">View Annual Leave </td>
                                    </tr>
                                    <tr>
                                        <td class="blue-normal" align="center" bgcolor="#FFFFFF" height="20"><%=strTitle1%></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top">
                                <table width="50%" border="0" cellpadding="0" cellspacing="0" align="center">
                                    <tr>
                                        <td class="blue" align="Right">Current Rate: &nbsp </td>
                                        <td class="red" align="left"><b><%=formatnumber(dblRatePerMonth,2)%></b>&nbsp days/month &nbsp&nbsp &nbsp </td>
                                    </tr>

                                    <tr height="35">
                                        <td class="blue" align="Right">Your annual leave balance:  &nbsp </td>
                                        <td class="red" align="left"><b><%=FormatNumber(dblBalanceByDays,2)%></b> &nbspdays &nbsp&nbsp &nbsp </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top">
                                <table width="60%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399">

                                    <tr>
                                        <td bgcolor="#8FA4D3">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
                                                <%if not isnull(expiredDate) then
					 if date < expiredDateThisYear then%>
                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="blue" width="75%">&nbsp&nbsp &nbsp<b>Leave brought forward from last year to <%=year(date)%> (hours) </b></td>
                                                    <td bgcolor="#E7EBF5" class="blue-normal" align="center" width="25%"><b><%=formatnumber(dblBalanceLastYear,2)%> </b></td>

                                                </tr>
                                                <%end if
				else%>
                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="blue">&nbsp&nbsp &nbsp<b>Leave brought forward from last year to <%=year(date)%> (hours) </b></td>
                                                    <td bgcolor="#E7EBF5" class="blue-normal" align="center"><b><%=formatnumber(dblBalanceLastYear,2)%> </b></td>

                                                </tr>
                                                <%end if%>

                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="blue">&nbsp&nbsp &nbsp<b>Leave Due until 1/<%=month(date)%>/<%=year(date)%> (hours)</b></td>
                                                    <td bgcolor="#E7EBF5" class="blue-normal" align="center"><b><%=formatnumber(dblLeaveDueThisYear+dblMoreHoursThisYear,2)%></b></td>

                                                </tr>
                                                <%if dblMoreHoursThisYear>0 then %>
                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="blue">&nbsp&nbsp &nbsp<b>Exceptions for <%=year(date)%> (hours)</b></td>
                                                    <td bgcolor="#E7EBF5" class="blue-normal" align="center"><b><%=formatnumber(dblMoreHoursThisYear,2)%></b></td>

                                                </tr>
                                                <%end if %>
                                                <tr height="25">
                                                    <td bgcolor="#617DC0" class="white" align="right"><b>Total (hours) </b>&nbsp&nbsp &nbsp</td>
                                                    <td bgcolor="#FFF2F2" class="red" align="center"><b><%=formatnumber(dblLeaveDueThisYear + dblBalanceLastYear ,2)%> </b></td>

                                                </tr>
                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="blue" width="70%">&nbsp&nbsp &nbsp<b>Annual leave in <%=year(date)%> (hours) </b></td>
                                                    <td bgcolor="#E7EBF5" class="blue-normal" align="center" width="30%"><b><%=FormatNumber(dblApplicationThisYear,2)%> </b></td>

                                                </tr>
                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="blue" width="70%">&nbsp&nbsp &nbsp<b>Annual leave reserved  (hours) </b></td>
<%dblApplicationReserve=GetApplication(year(Date),intstaffID,Date(),cdate("31-Dec-" & year(date)+1))%>														
                                                    <td bgcolor="#E7EBF5" class="blue-normal" align="center" width="30%"><b><%=FormatNumber(dblApplicationReserve,2)%> </b></td>
                                                </tr>
                                                <tr height="25">
                                                    <td bgcolor="#617DC0" class="white" align="right"><b>Balance  (hours) </b>&nbsp&nbsp &nbsp</td>
                                                    <td bgcolor="#FFF2F2" class="red" align="center"><b><%=FormatNumber(dblBalance,2)%> </b></td>

                                                </tr>
                                                <tr height="25">
                                                    <td bgcolor="#C2CCE7" class="white" align="right"></td>
                                                    <td bgcolor="#FFF2F2" class="red" align="center"><b><%if dblWorkingHour<>0 then%><%=FormatNumber(dblBalance/dblWorkingHour,2)%><%else%>0<%end if%> (days)</b></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>&nbsp</td>
                                    </tr>
                                      <tr>
                                        <td valign="top" class="text-right">
                                            <a href="<%=strHTTP%>aisnet/OverviewOfRequestHistory/OverviewOfRequestHistory" class="btn-custom">Calendar</a>
                                            <a href="<%=strHTTP%>aisnet/HolidayBooking/HolidayBooking" class="btn-custom">Holidays Booking</a>
											 <a href="javascript:ANTracking()" class="btn-custom">Tracking</a>
                                        </td>
                                    </tr>
                                    
                                     <tr>
                                        <td>&nbsp</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top">
                                <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399">
                                    <tr>
                                        <td bgcolor="#8FA4D3">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">

                                                <tr bgcolor="#617DC0" height="25">
                                                    <td class="white" align="center" width="12%"><b>Date applied</b></td>
                                                    <td class="white" align="center" width="12%"><b>Number years at Atlas</b></td>
                                                    <td class="white" align="center" width="12%"><b>Extra leave for
                                                        <br>
                                                        long service </b></td>
                                                    <td class="white" align="center" width="12%"><b>Rate for level</b></td>
                                                    <td class="white" align="center" width="12%"><b>Total days/year</b></td>
                                                    <td class="white" align="center" width="12%"><b>Rate per month</b></td>
                                                    <td class="white" align="center" width="28%"><b>Note</b></td>
                                                </tr>
<%Response.Write strHistoryRate %>		
                                            </table>
                                        </td>
                                    </tr>

                                    <tr>
                                        <td bgcolor="#FFFFFF" class="blue">&nbsp;
                                        </td>
                                    </tr>

                <%if not isnull(expiredDateThisYear) AND Year(Date())<>2022 then 
					 if date < expiredDateThisYear then
						dblApplicationReserveBeforeExpired= GetApplication(year(Date),intstaffID,cdate("1-Jan-" & year(date)),expiredDateThisYear)%>
                                    <tr>
                                        <td bgcolor="#FFFFFF" class="blue-normal">&nbsp;* Annual leave balance for <%Response.Write(year(Date)-1)%>
                                            <%if dblBalanceLastYear>0 then%> -
           					<span class="red"><b><%=FormatNumber(cdbl(dblBalanceLastYear)/cdbl(dblWorkingHour),2)%> days</b></span> -
           				<%end if%>
           				 will expire on &nbsp;
           					<span class="red"><b><%=day(expiredDateThisYear) & "-" & MonthName(month(expiredDateThisYear),true) & "-" & year(expiredDateThisYear)%></b></span>
                                        </td>
                                    </tr>
                                    <%if dblBalanceLastYear>0 then%>
                                    <tr>
                                        <td bgcolor="#FFFFFF" class="blue-normal">&nbsp;*  Annual leave booked from <b>1-Jan-<%=year(expiredDateThisYear)%></b>
                                            to <b><%=day(expiredDateThisYear) & "-" & MonthName(month(expiredDateThisYear),true) & "-" & year(expiredDateThisYear)%></b>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           				<span class="red"><b><%if cdbl(dblApplicationReserveBeforeExpired)>0 then%>
                               <%=formatnumber(cdbl(dblApplicationReserveBeforeExpired)/cdbl(dblWorkingHour),2)%> <%else%>0.00<%end if%>&nbsp; (days)</b></span>
                                        </td>
                                    </tr>
                                    <%end if%>
                                    <tr>
                                        <td bgcolor="#FFFFFF" class="blue-normal">&nbsp;* Balance to use before <b><%=day(expiredDateThisYear) & "-" & MonthName(month(expiredDateThisYear),true) & "-" & year(expiredDateThisYear)%></b>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           				<span class="red"><b><%if cdbl(dblBalanceLastYear)>cdbl(dblApplicationReserveBeforeExpired) then%>
                               <%=formatnumber(cdbl(dblBalanceLastYear-dblApplicationReserveBeforeExpired)/cdbl(dblWorkingHour),2)%> <%else%>0.00<%end if%>&nbsp; (days)</b></span>
                                        </td>
                                    </tr>
                                    <%end if
			  end if%>
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
