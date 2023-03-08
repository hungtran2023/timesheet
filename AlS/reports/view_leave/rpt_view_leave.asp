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
	
	Dim intUserID, intMonth, intYear, dblCurLeave
	Dim strConnect, objDatabase, strError
	Dim rsDuration,rsIndividualRule,dblBalance,dblApplication,dblLeaveDue,dblMoreHoursThisYear
	Dim dateTo

Function GetApplication(intYear,staffID,dateF,DateT)

	dim strSql,strTable
	dim rs, dblApplication

	dblApplication=0

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
'***************************************************************
'
'***************************************************************
Function GetNumberOfMonthForAnnualLeave(dateStart,dateF,dateT)
	
	dim dblNum

'Response.Write dateStart & "- From: " & dateF & "--To:" & dateT 	
	'from the firstdate to the lastdate of month	
	if day(dateF)=1 AND Month(dateT)<>Month(dateT + 1) then
		dblNum=DateDiff("m",dateF,dateT)+1		
	else
		dblNum=0
		'If dayFrom from 6 to 15 then get a haft
		if dateF=dateStart and day(dateF)>5 and day(dateF)<=15 then
			dblNum=0.5
			'dateF=cdate(month(dateF) + 1 & "/1/" & year(dateF))
			dateF=DateAdd("m",1,dateF) - day(dateF)+1 
		'If dayFrom>15 then move to next month
		elseif day(dateF)>15 then
			'dateF=cdate(month(dateF) + 1 & "/1/" & year(dateF))	
			dateF=DateAdd("m",1,dateF) - day(dateF)+1 	
		'If dayFrom from 1-5 then get full month
		else
			dateF=dateF-day(dateF) + 1
		end if		
				
		'Response.Write dateStart & "- From: " & dateF & "--To:" & dateT & "-->" & dblNum & "<br>"	

		if Month(dateT)=Month(dateT + 1) then
			if day(dateT)>=15 then
'Response.Write " chuyen dateTo " & dateT & "-->" & cdate(month(dateT) + 1 & "/1/" & year(dateT))-1 & "<br>"

				'dateT=cdate(month(dateT) + 1 & "/1/" & year(dateT))-1
				dateT=DateAdd("m",1,dateT) - day(dateT)
				
			else
				dateT=dateT-day(dateT)
			end if
		end if
		
		if dateF<dateT then	dblNum=dblNum +	DateDiff("m",dateF,dateT) + 1
		
	end if

'Response.Write dateStart & " ===== From: " & dateF & "--To:" & dateT & "-->" & dblNum & "<br>"				
	GetNumberOfMonthForAnnualLeave = dblNum
end function
'**************************************************************
'
'**************************************************************
Function GetHistoryOfAnnualLeaveRate(staffID)
	
	dim rs, strOut
	dim numYear,numMonth,totalperYear,ratepermonth
	dim strworkingYear
	
	strOut=""
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then
	
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "GetAnnualLeaveRate"
		Set myParam = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
			
		myCmd("StaffID") = intStaffID
		
		SET rs=myCmd.Execute
		
		if Not rs.EOF then
			do while not rs.EOF	
				strworkingYear=""
				if cdbl(rs("numberofmonth"))>0 then
					numYear=cdbl(rs("numberofmonth")) \ 12
					numMonth=cdbl(rs("numberofmonth")) mod 12
					if numYear>0 then strworkingYear=numYear & " year(s)" 
					if (numMonth>0 and numYear>0) then strworkingYear=strworkingYear & " & "
					if numMonth>0 then strworkingYear=strworkingYear & numMonth & " month(s)"
				end if
				totalperYear=cdbl(rs("longservice"))+ cdbl(rs("rateperyear"))
				
				strOut=strOut & "<tr bgcolor='#E7EBF5' height='25' > " & _
								"<td class='blue-normal'>&nbsp;&nbsp;" & day(rs("applydate")) & "-" & MonthName(month(rs("applydate")),true) & "-" & year(rs("applydate")) & "</td>" & _
								"<td class='blue-normal' align='center'>" & numYear & "</td>" & _
								"<td class='blue-normal' align='center'>" & rs("longservice") & "</td>" & _
								"<td class='blue-normal' align='center' >" & rs("rateperyear") & "</td>" & _
								"<td class='blue-normal' align='center'><b>" & totalperYear & "<b></td>" & _
								"<td class='red' align='center'><b>" & FormatNumber(totalperYear/12,2) & "</b></td>" & _
								"<td class='blue-normal'>&nbsp;&nbsp;"& strworkingYear & "</td></tr>"
				rs.MoveNext
			loop
		end if
		
		
	end if
	
	GetHistoryOfAnnualLeaveRate=strOut

End Function

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	dateTo=Date()
	
	intMonth = Request.Form("M")
	intYear = Request.Form("Y")
	
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

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	intDepartID = varFullName(2)

	Set objEmployee = Nothing


	dblApplication=0
	dblBalance=0
	dblBalanceLastYear=0
	dblLeaveDueThisYear	=0
	dblApplicationThisYear=0
	dblLeaveDue=0
	dblMoreHours=0	
	dblMoreHoursThisYear=0

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
	if not rsIndividualRule.EOF then
		expiredDate=cdate(rsIndividualRule("ExpiredMonth") & "/" & rsIndividualRule("ExpiredDay") & "/" &  Year(date))
	end if	
	If objDatabase.dbConnect(strConnect) Then
	
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "GetDurationAnnualLeave"
		Set myParam = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("dateTo",adDate,adParamInput)
		myCmd.Parameters.Append myParam
	
		myCmd("StaffID") = intStaffID
		myCmd("dateTo")=dateTo
		
		set rsDuration= myCmd.Execute
		
		if not rsDuration.EOF then
						
			dateStart=rsDuration("DateFrom")
			intYear=0
		
			Do while not rsDuration.EOF	
				dblKeepPassYear=cdbl(rsDuration("KeepPassYear"))
				
				if intYear<>rsDuration("YearAN") then 
					dblBalanceLastYear=dblBalance
					dblLeaveDueThisYear=0
					dblApplicationThisYear=0
					
					dblMoreHours=0
					intYear=rsDuration("YearAN")
                    			
				    expiredDate=null
	                if not rsExpireday.EOF then
		                expiredDate=cdate(rsExpireday("ExpiredMonth") & "/" & rsExpireday("ExpiredDay") & "/" &  intYear)
	                end if		

					if not rsIndividualRule.EOF then				
						rsIndividualRule.MoveFirst
						rsIndividualRule.Filter="ApplyYear=" & intYear
						
						if not rsIndividualRule.Eof then 
						    dblMoreHours=rsIndividualRule("MoreHours")						
						    if intYear=year(Date()) then dblMoreHoursThisYear=cdbl(dblMoreHours)
						    if rsIndividualRule("ExpiredDay") <>0 AND rsIndividualRule("ExpiredMonth")<>0 then
		    				    dblKeepPassYear=cdbl(rsIndividualRule("KeepPassYear"))
		    				    expiredDate=cdate(rsIndividualRule("ExpiredMonth") & "/" & rsIndividualRule("ExpiredDay") & "/" &  intYear)
		    				end if

						end if
						rsIndividualRule.Filter=""
					end if
				end if		
				
				dblApplication=GetApplication(rsDuration("YearAN"),intstaffID,rsDuration("DateFrom"),rsDuration("DateTo"))
				dblRatePerMonth=(cdbl(rsDuration("RatePerYear"))+ cdbl(rsDuration("RateByYTD")))/12	
						
				dblNumOfMonths=GetNumberOfMonthForAnnualLeave(dateStart,rsDuration("DateFrom"),rsDuration("DateTo"))
				
				
				dblWorkingHour=cdbl(rsDuration("WorkingHours"))
				'dblKeepPassYear=cdbl(rsDuration("KeepPassYear"))
				
				dblLeaveDue=dblNumOfMonths * dblRatePerMonth * dblWorkingHour
				
				if cdbl(dblMoreHours)<>0 and (rsDuration("ANType")="End" OR rsDuration("ANType")="Inprogress") then					
						dblLeaveDue=dblLeaveDue+ cdbl(dblMoreHours)
						dblMoreHours=0					
				end if				
				


				dblLeaveDueThisYear=dblLeaveDueThisYear+dblLeaveDue
				dblApplicationThisYear=dblApplicationThisYear + dblApplication
								
				dblBalance=dblBalance + dblLeaveDue  - dblApplication		
 
'Response.Write "Expired date" & rsDuration("DateTo") & "#" & expiredDate-1		& "<br>"		
				if rsDuration("ANType")="Expired" or (rsDuration("ANType")="Inprogress" AND rsDuration("DateTo")=expiredDate-1) then
					if dblBalance- dblLeaveDueThisYear>0 then dblBalance=dblLeaveDueThisYear					
'Response.Write 	intYear & "#" & dblBalanceLastYear & "#" & 	dblApplicationThisYear	& "<br>"				

					dblApplicationThisYear=dblApplicationThisYear-dblBalanceLastYear
					if dblApplicationThisYear<0 then
						dblApplicationThisYear=0
					end if
					dblBalanceLastYear=0
				end if
				
				if rsDuration("ANType")="End" then
					if dblKeepPassYear>0 and dblBalance>(dblKeepPassYear * dblWorkingHour) then dblBalance =(dblKeepPassYear * dblWorkingHour)						
				end if 			
				
'if intStaffID=1014 then Response.Write intYear & "#" & dblNumOfMonths & "#" & dblApplication & "#" & dblLeaveDue & "#" & 	dblBalance & "#" & dblMoreHours	& "<br>"	
'Response.Write dblApplication & "#" & dblLeaveDue & "#" & dblBalance & "<br>"	

				rsDuration.MoveNext
			loop
			
			rsDuration.close
			set rsDuration=nothing
		end if
		
		dblApplicationThisMonth=GetApplication(year(Date()),intstaffID,date-Day(Date)+1,Date)
		dblApplicationReserve= GetApplication(year(Date),intstaffID,date + 1,cdate("31-Dec-" & year(date)))
		dblApplicationReserveNextYear=GetApplication(year(Date)+1,intstaffID,cdate("1-Jan-" & year(date)+1),cdate("31-Dec-" & year(date)+1))
		dblApplicationReserve=dblApplicationReserve+dblApplicationReserveNextYear
			
		'For expired at the first month of year
		if month(Date)=1 then
			dblBalanceLastYear=dblBalance
			dblLeaveDueThisYear=0
			dblApplicationThisYear=dblApplicationThisMonth
		else
			dblApplicationThisYear=dblApplicationThisYear + dblApplicationThisMonth
		end if
		
		dblBalance=dblBalance-dblApplicationThisMonth-dblApplicationReserve
		if dblWorkingHour<>0 then	dblBalanceByDays=dblBalance/dblWorkingHour
	Else
		strError = objDatabase.strMessage
	End If		
	

		
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'><a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view timesheet&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>" & _
					"&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" & _ 
					strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
					Help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
					"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Annual Leave of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
					
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
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<script language="javascript" src="../../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie,objNewWindow;
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

function selstaff()
{
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=490,height=325,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('rpt_select_staff.asp', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function goback()
{
	if (ns)
		document.location = "rpt_list_staff.asp?b=1";
	else
	{
		window.document.frmtms.action = "rpt_list_staff.asp?b=1";
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
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
	Response.Write(arrTmp(0))
%>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
		  <tr> 
		    <td class="title" height="20" align="center">&nbsp;&nbsp;</td>
		  </tr>
		  <tr> 
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
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
<%End If%>

				<tr>
				  <td class="blue" height="50"> 
					
						&nbsp&nbsp &nbsp&nbsp&nbsp &nbsp <a href="rpt_list_staff.asp?act=vpsn">Annual Leave Overview</a>
				  </td>	
				</tr>			    
			    <tr> 
				  <td class="title" align="center">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr> 
							<td class="title" align="center">View Annual Leave </td>
						</tr>
  						<tr> 
							<td class="blue-normal" align="center" bgcolor="#FFFFFF" height="20"><%=strTitle1%></td>
						</tr>		
						  <tr> 
							<td class="blue-normal" align="center" bgcolor="#FFFFFF" height="20">&nbsp&nbsp &nbsp</td>
						</tr>	  
			</table>
				  </td>
				</tr>  
              </table>
		    </td>
		  </tr>
		  <tr>
			<td valign="top">
				<table width="50%" border="0" cellpadding="0" cellspacing="0" align="center" >
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
			<td valign="top">&nbsp &nbsp</td>
		</tr>
		<tr>
			<td valign="top">
			<table width="80%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" >
				
			   <tr> 
           		<td bgcolor="#8FA4D3"> 
			      <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
 			  <%if not isnull(expiredDate) then 				
					 if date < expiredDate then%>		      
                    <tr height="25"> 
                      <td bgcolor="#C2CCE7" class="blue"  width="75%">&nbsp&nbsp &nbsp<b>Leave brought forward from last year to <%=year(date)%> (hours) </b></td>
                      <td bgcolor="#E7EBF5" class="blue-normal" align="center" width="25%"><b> <%=formatnumber(dblBalanceLastYear,2)%> </b></td>
                      
					</tr> 					
					<%end if
				else%>
					<tr height="25"> 
                      <td bgcolor="#C2CCE7" class="blue" >&nbsp&nbsp &nbsp<b>Leave brought forward from last year to <%=year(date)%> (hours) </b></td>
                      <td bgcolor="#E7EBF5" class="blue-normal" align="center"><b> <%=formatnumber(dblBalanceLastYear,2)%> </b> </td>
                      
					</tr> 					
				<%end if%>
				
					 <tr height="25"> 
                      <td bgcolor="#C2CCE7" class="blue" >&nbsp&nbsp &nbsp<b>Leave Due until 1/<%=month(date)%>/<%=year(date)%> (hours)</b></td>
                      <td bgcolor="#E7EBF5" class="blue-normal" align="center"><b> <%=formatnumber(dblLeaveDueThisYear-dblMoreHoursThisYear,2)%></b></td>
                      
					</tr>
                <%if dblMoreHoursThisYear>0 then %>
					 <tr height="25"> 
                      <td bgcolor="#C2CCE7" class="blue" >&nbsp&nbsp &nbsp<b>Exception for <%=year(date)%> (hours)</b></td>
                      <td bgcolor="#E7EBF5" class="blue-normal" align="center"><b> <%=formatnumber(dblMoreHoursThisYear,2)%></b></td>
					</tr>				    
				<%end if %>							
					<tr height="25"> 
                      <td bgcolor="#617DC0" class="white" align="right"><b>Total (hours) </b>&nbsp&nbsp &nbsp</td>
                      <td bgcolor="#FFF2F2" class="red" align="center"><b> <%=formatnumber(dblLeaveDueThisYear + dblBalanceLastYear,2)%> </b></td>
                      
					</tr>
					<tr height="25"> 
                      <td bgcolor="#C2CCE7" class="blue" width="70%">&nbsp&nbsp &nbsp<b>Annual leave in <%=year(date)%> (hours) </b></td>
                      <td bgcolor="#E7EBF5" class="blue-normal" align="center" width="30%"><b> <%=FormatNumber(dblApplicationThisYear,2)%> </b></td>
                      
					</tr> 
					<tr height="25"> 
                      <td bgcolor="#C2CCE7" class="blue" width="70%">&nbsp&nbsp &nbsp<b>Annual leave reserved  (hours) </b></td>
                      <td bgcolor="#E7EBF5" class="blue-normal" align="center" width="30%"><b> <%=FormatNumber(dblApplicationReserve,2)%> </b></td>
					</tr> 					
					<tr height="25"> 
                      <td bgcolor="#617DC0" class="white" align="right"><b>Balance  (hours) </b>&nbsp&nbsp &nbsp</td>
                      <td bgcolor="#FFF2F2" class="red" align="center"><b> <%=FormatNumber(dblBalance,2)%> </b></td>
                      
					</tr>		
					
										<tr height="25"> 
                      <td bgcolor="#C2CCE7" class="white" align="right"></td>
                      <td bgcolor="#FFF2F2" class="red" align="center"><b> <%if dblWorkingHour<>0 then%><%=FormatNumber(dblBalance/dblWorkingHour,2)%><%else%>0<%end if%> (days)</b></td>
                      
					</tr>
								
			      </table>
				</td>
				</tr>
				<tr><td> &nbsp</td></tr>
				</table>
			</td>
		</tr>
		  
		 <tr>

		  <td valign="top"> 
			<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" >			
			   <tr> 
           		<td bgcolor="#8FA4D3"> 
			      <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
			      
                    <tr bgcolor="#617DC0" height="25"> 
                      <td class="white" align="center"  width="15%"><b>Date applied</b></td> 
                      <td class="white" align="center" width="12%"><b>Number years at Atlas</b></td>
                      <td class="white" align="center" width="12%"><b>Extra leave for <br> long service </b></td>
                      <td class="white" align="center" width="12%"><b>Rate for level</b></td>
                      <td class="white" align="center" width="12%"><b>Total days/year</b></td>
                      <td class="white" align="center" width="12%"><b>Rate per month</b></td>
                      <td class="white" align="center" width="25%"><b>Note</b></td>
					</tr>	
<%=GetHistoryOfAnnualLeaveRate(intStaffID)%>							      

					
				  </table>
				</td>
			  </tr>	
			  
			   <tr> 
           			<td bgcolor="#FFFFFF" class="blue"> 
           				&nbsp;
					</td>
			  </tr>				  		
			
 			  <%if not isnull(expiredDate) then 				
					 if date < expiredDate then	
		 
						dblApplicationReserveBeforeExpired= GetApplication(year(Date),intstaffID,cdate("1-Jan-" & year(date)),expiredDate)%>
			   <tr> 
           			<td bgcolor="#FFFFFF" class="blue-normal"> 
           				&nbsp;* Annual leave balance for <%Response.Write(year(Date)-1)%>
           				<%if dblBalanceLastYear>0 then%> - 
           					<span class="red"><b><%=FormatNumber(cdbl(dblBalanceLastYear)/cdbl(dblWorkingHour),2)%> days</b></span> - 
           				<%end if%>
           				 will expire on &nbsp;
           					<span class="red"><b><%=day(expiredDate) & "-" & MonthName(month(expiredDate),true) & "-" & year(expiredDate)%></b></span>
					</td>
			   <%if dblBalanceLastYear>0 then%>
			   <tr> 
           			<td bgcolor="#FFFFFF" class="blue-normal"> 
           				&nbsp;*  Annual leave booked from <b>1-Jan-<%=year(expiredDate)%></b>            				
           				 to <b><%=day(expiredDate) & "-" & MonthName(month(expiredDate),true) & "-" & year(expiredDate)%></b>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           				<span class="red"><b><%if cdbl(dblApplicationReserveBeforeExpired)>0 then%> 
           					<%=formatnumber(cdbl(dblApplicationReserveBeforeExpired)/cdbl(dblWorkingHour),2)%> <%else%>0.00<%end if%>&nbsp; (days)</b></span>
					</td>
			  </tr>
			  <%end if%>
			  <tr> 
           			<td bgcolor="#FFFFFF" class="blue-normal">&nbsp;* Balance to use before <b><%=day(expiredDate) & "-" & MonthName(month(expiredDate),true) & "-" & year(expiredDate)%></b>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           				<span class="red"><b><%if cdbl(dblBalanceLastYear)>cdbl(dblApplicationReserveBeforeExpired) then%> 
           					<%=formatnumber(cdbl(dblBalanceLastYear-dblApplicationReserveBeforeExpired)/cdbl(dblWorkingHour),2)%> <%else%>0.00<%end if%>&nbsp; (days)</b></span>
           			</td>
			  </tr>	
					<%end if
			  end if%>			  
			  
			  
			</table>
		  </td>
		</tr>
		
	 <!-- <tr> 
		
		
		<td align="center" bgcolor="#FFFFFF" height="50">
			<img src='..\..\website_page_under_construction.jpg' width="300" height="200"><br>
			This page is under construction so the data is incorrect.<br>Please go to timesheet page for viewing Annual Leave</td>
	</tr>-->			
        </table>      
<%
	Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))
	
	Set myCmd = Nothing
	Set myCmd_ = Nothing
	Set objDatabase = Nothing
%>
<input type="hidden" name="txthidden" value="<%=intStaffID%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">

</form>
</body>
</html>
