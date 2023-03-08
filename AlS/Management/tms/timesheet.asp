<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim intUserID, intMonth, intYear, intWeekday, intDayNum, intDayCol, intDayCount, intRow, eRow, intTotalRow, ii, kk, intCurMonth 
	Dim dblHour, dblTotal, strHour
	Dim strFirstDay, strParm, strURLSetHour, strColorOpt, strError, varTimesheet, varEvent,dateLimit, fgEnterAnnualy,arrFingerprint
	dim strDateLock
'***************************************************************
'
'***************************************************************
sub GetFingerprintByStaff()
	
	dim ii, intColumns
	dim objConn
	dim rs

	strOut=""
	strConnect = Application("g_strConnect")
	set objConn= Server.CreateObject("ADODB.Connection")
	objConn.Open strConnect  
	
	If objConn.State=1 Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objConn
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "FingerprintByStaff"
		Set myParam = myCmd.CreateParameter("month",adInteger,adParamInput)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("year",adInteger,adParamInput)
		myCmd.Parameters.Append myParam

		Set myParam = myCmd.CreateParameter("staffID",adNumeric,adParamInput)
		
		myParam.Precision = 18
		myParam.NumericScale = 0


		myCmd.Parameters.Append myParam
		myCmd("month") = intMonth
		myCmd("year") = intyear
		myCmd("staffID") = intStaffID
	  
		SET rs=myCmd.Execute
		arrFingerprint=rs.GetRows()
		

	end if
	
End sub
'--------------------------------------------------
'
'--------------------------------------------------	
Function Format_hhmm(dblMinuteIn)
	dim strOut
	dim dblHour, dblMinute
	strOut=""
	if dblMinuteIn<>0 then
		dblHour=cint(abs(dblMinuteIn)\60)
		dblMinute=abs(dblMinuteIn) mod 60
		strOut=dblHour
		if dblHour<10 then strOut="0" & strOut
		strOut=strOut &":" 
		if dblMinute<10 then
			strOut=strOut & "0" & dblMinute
		else
			strOut=strOut & dblMinute
		end if
		if dblMinuteIn<0 then strOut="-" & strOut
	end if
	
	Format_hhmm=strOut
end function	
'--------------------------------------------------
'
'--------------------------------------------------
function RecordsetFilter(dateCompare)	
	dim blnReturn
	
	blnReturn=false	
	
	if rsApproval.RecordCount>0 then
		rsApproval.MoveFirst
		Do While NOT rsApproval.EOF and not blnReturn
			blnReturn = (rsApproval("DateFrom")<=dateCompare) and (rsApproval("Dateto")>=dateCompare)
			rsApproval.MoveNext
		loop
	end if
	
	'blnReturn=false
	RecordsetFilter=blnReturn
	
end function 

	dateLimit=cint(365)
	strDateLock=Cdate("30-Apr-2022")
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	If Request.Form("M") = "" Then
		intMonth = Month(Date)
	Else
		intMonth = Request.Form("M")
	End If
	If Request.Form("Y") = "" Then
		intYear = Year(Date)
	Else
		intYear	= Request.Form("Y")
	End If		

	'From Approval by team
	strFr=request.form("txtFrApp")
	if strFr<>"" then
		intStaffIDs  = Request.Form("txtstaffIDs")	
		strWeekStart= Request.Form("txtstartDate")
		if request.querystring("app")<>"" then
			intMonth=month(strWeekStart)
			intYear=year(strWeekStart)
		end if
	end if
	
	intCurMonth = Month(Date)
	strAction	= Request.QueryString("act")
	
	intRow		= -1
	eRow		= -1
	intDayNum	= GetDay(intMonth,intYear)				' Numbers of days in a month
	intDayCol	= intDayNum + 6

	fgEnterAnnualy=false
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If

	intUserID	= session("USERID")

	intStaffID  =request.querystring("id")
	if intStaffID="" then intStaffID  = Request.Form("txthidden")
	
	
	strFirstDay = FirstOfMonth(intMonth,intYear)		' Get the first day in a month				
	intDayCount	= curDayNum(intDayNum,strFirstDay)		' Numbers of days since the first day in month to now
	
	if intUserID=252 then strDateLock=Cdate("30-Sep-2017")
	'if intUserID=252 OR intUserID=1242 then strDateLock=Cdate("30-Sep-2016")
	'if intUserID=2134  then strDateLock=Cdate("31-Dec-2020")

'--------------------------------------------------
' Check Enter Annual Leave right
'--------------------------------------------------

	If not isEmpty(session("RightOn")) Then
		varGetRight = session("RightOn")
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "Write Timesheet as HR control" OR intUserID=252 Then
				fgEnterAnnualy = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If


'--------------------------------------------------
' The timesheet array initializing function is called 
' when session("varTimesheet")/session("varEvent") is not initialized
' or user changes month/year to view timesheet    
'--------------------------------------------------

	If Request.QueryString("act") = "" Then
		If Not IsEmpty(session("varTimesheet")) And Not IsEmpty(session("varEvent")) Then
			session("varTimesheet") = Empty
			session("varEvent") = Empty
		End If
	End If

	If (IsEmpty(session("varTimesheet")) And IsEmpty(session("varEvent"))) Or (Request.QueryString("act") = "vmya") Then
		strError	=  tmsInitial(intStaffID,intMonth,intYear)
		If strError = "" Then
			varTimesheet = session("varTimesheet")		' Array stores timesheet data
			varEvent	 = session("varEvent")			' Array stores event data
		Else
			varEvent	 = session("varEvent")	
		End If
	Else
		varTimesheet = session("varTimesheet")			' Array stores timesheet data
		varEvent	 = session("varEvent")				' Array stores event data
	End If
	
	If isarray(varTimesheet) Then
		intRow	= Ubound(varTimesheet,3)
	End If
	
	If isarray(varEvent) Then
		eRow	= Ubound(varEvent,3)
	End If
'--------------------------------------------------
' Get Incompleted Timesheet remind
'--------------------------------------------------
	'strReminer=""
	'if cint(intUserID)=cint(intStaffID) then
	
	'	strReminer=RemindIncompleteTimesheet(date()-7,date()-1, intUserID)
	'end if
'--------------------------------------------------
' Get fingerprint scanner data
'--------------------------------------------------	
	strSQL="SELECT DATENAME(dw , CalanderDate) ,CalanderDate,  CONVERT(VARCHAR(5),Enter,108) as Enter, CONVERT(VARCHAR(5),ExitHour,108) as ExitHour, CAST(FORMAT(Duration/60,'00') as nvarchar) +':' +CAST(FORMAT(Duration%60,'00') as nvarchar) as Duration  , Duration as DurationMinute, " & _
				"CASE WHEN Duration IS NOT NULL THEN FORMAT((Duration-(WkHours+OTHours)*60)/60,'00') ELSE '00' END as DailyBalanceHours , " & _
				"CASE WHEN Duration IS NOT NULL THEN FORMAT((Duration-(WkHours+OTHours)*60)%60,'00') ELSE '00' END as DailyBalanceMinute, "&_ 
				" CASE WHEN Duration IS NOT NULL THEN Duration-(WkHours+OTHours)*60 ELSE 0 END as DailyBalance, " & _
				" (WkHours) as WkHours, LeaveHours as OffHour, OTHours as OT, '' as a FROM  GetDaysOfMonth(" & intMonth & "," & intYear & ") a " &_
				"FULL OUTER JOIN (SELECT * FROM rp_AccessDoorAndTimesheet " & _
			"WHERE (MONTH(Tdate) = " & intMonth & ") AND (YEAR(Tdate) =" & intYear & ") AND PersonID=" & intStaffID & ") b ON a.CalanderDate=b.Tdate " &_
				" ORDER BY CalanderDate"
			
	'Call GetRecordset(strSQL,rsFingerprint)
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view timesheet&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:viewdetail()' onMouseOver='self.status=&quot;View timesheet detail&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>View Detail</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
  				  "<a class='c' href='javascript:viewleave()' onMouseOver='self.status=&quot;View annual leave of staff&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>View Leave</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage()' onMouseOver='self.status=&quot;Print timesheet page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Timesheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"

	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "templates/template1/")
%>	

<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" language="javascript" onUnload="return window_onunload();">
<form name="frmtms" method="post">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
	If strError <> "" Then
%>
  <tr height="20">
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td class="red" align="center">&nbsp;<b><%=strError%></b></td>
        </tr>
      </table>    
    </td>
  </tr>  
<%	End if%>      
  <tr> 
    <td valign="top">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td bgcolor="#8FA4D3"> 
            <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
              <tr> 
                <td colspan="2" rowspan="2" class="white" bgcolor="#617DC0"> 
                  <div align="center"> <b>Project </b> </div>
                </td>
                <td width="20%" colspan="<%=intDayNum%>" class="blue-normal" align="right" bgColor="#617DC0">  
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="blue-normal">
                    <tr> 
                      <td width="60%" class="white">&nbsp;&nbsp;<%=strTitle1%></td>
                      <td align="right" width="35%">
					    <select name="lbyear" size="1" class="blue-normal"><%For ii=2000 To Year(Date) +1%>
					      <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option><%Next%>
						</select>
						<select name="lbmonth" size="1" class="blue-normal"><%For ii=1 To 12 %>						
						  <option <%If CInt(intMonth)=ii Then%>selected<%End If%> value="<%=ii%>"><%=SayMonth(ii)%></option><%Next%>
						</select>
                      </td>
                        <td width="5%" align="right">  
                        <table border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa" width="40">
                          <tr> 
                            <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
                              <div align="center" class="blue"><a href="javascript:viewtms();" class="b">Go</a></div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td width="4%" rowspan="2" class="white" bgcolor="#617DC0"> 
                  <div align="center"><b>Total</b></div>
                </td>
              </tr>
              <tr bgcolor="#617DC0">
<%
				For kk=1 To intDayNum
					intWeekDay = WeekDay(strFirstDay+(kk-1))
					strColorOpt = ""
					Select Case intWeekDay
						Case 1
							strColorOpt = DAYCOLOR
						Case 7
							strColorOpt = DAYCOLOR
					End Select
					If Not strColorOpt="" Then%>
				        <td width="2%"><div align="center" class="white"><font color="<%=strColorOpt%>"><b><%=kk%></b></font></div></td>
<%					Else%>
   				        <td width="2%"><div align="center" class="white"><b><%=kk%></b></div></td>
<%					End If
   				Next%>  
			  </tr>

					  
<!--**************************** For Project And SubTask *********************************************-->
					  
<%
	intTotalRow = intRow
	If intTotalRow <= 9 Then
		intTotalRow = 9
	End If	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	strSql="SELECT c.AssignmentID,DateTransfer,c.StaffID,c.FgDelete, d.fgActivate " & _
			"FROM (SELECT ProjectID, MIN(DateTransfer) as DateTransfer " & _
					"FROM ATC_Projectstage GROUP BY ProjectID) a " & _
			"INNER JOIN ATC_Tasks b ON a.ProjectID=b.ProjectID " & _
			"INNER JOIN ATC_Assignments c ON b.SubTaskID=c.SubTaskID " & _
			"INNER JOIN ATC_Projects d ON a.ProjectID=d.ProjectID " & _
			"WHERE c.StaffID=" & intStaffID & " ORDER BY c.AssignmentID"

	
	strSql1="SELECT DateFrom,DateTo FROM ATC_TimesheetApproval WHERE StaffID=" & intStaffID & " AND (Month(DateFrom)=" & intMonth & " OR Month(DateTo)=" & intMonth & ") AND (YEAR(DateFrom)=" & intYear & " OR YEAR(DateTo)=" & intYear & ") " 
'response.write strSql1
	If objDatabase.dbConnect(strConnect) Then
	
		Call GetRecordset(strSql,rsAssignment)
		Call GetRecordset(strSql1,rsApproval)

		For ii = 0 To intTotalRow
			If ii <= intRow Then
				If varTimesheet(intDayCol-1,0,ii) = 0 Then
%>					  	
              <tr> 
                <td height="20" bgcolor="#FFC6C6"><img src="../../images/cross.gif" width="8" height="14"></td>
<%
					If trim(varTimesheet(intDayCol-2,0,ii)) = "S" Then
%>                        
                <td lass="blue" bgcolor="#FFF2F2"><a href="javascript:menufunctions('<%=varTimesheet(intDayCol-4,0,ii)%>','<%=ii%>');" title="<%=varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;&nbsp;&nbsp;- <%=showlabel(varTimesheet(intDayCol-3,0,ii))%></b></a></td>
<%
					ElseIf trim(varTimesheet(intDayCol-2,0,ii)) = "N" Then
%>                        
                <td class="blue" bgcolor="#FFF2F2"><a href="javascript:void(0)" title="<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=showlabel(varTimesheet(0,0,ii))%></b></a></td>
<%	
					ElseIf trim(varTimesheet(intDayCol-2,0,ii)) = "P" Then
%>                        
                <td class="blue" bgcolor="#FFF2F2"><a href="javascript:menufunctions('<%=varTimesheet(intDayCol-4,0,ii)%>','<%=ii%>');" title="<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=showlabel(varTimesheet(0,0,ii))%></b></a></td>
<%
					End If
	'on error resume next
					For kk = 1 To intDayNum
						dblHour = varTimesheet(kk, 0, ii) + varTimesheet(kk, 1, ii)
'if intStaffID=604 then Response.Write kk & ":" & dblHour & "<br>"						
						strHour = "&nbsp;"
						strCurrentDate = CDate(intMonth & "/" & kk & "/" & intYear)									
		
						strParm = CStr(ii) & "," & CStr(kk) & ",'P'" 
						strURLSetHour = "javascript:sethour("& strParm & ");"						

						If kk <= intDayCount Then
'Response.Write kk & "--" & dblHour						
							If dblHour > 0 Then
								strHour = formatnumber(dblHour,1)
								
								if (strCurrentDate>=date()-dateLimit) AND strCurrentDate>strDateLock then
									rsAssignment.MoveFirst									 
									rsAssignment.Find "AssignmentID = " & varTimesheet(intDayCol-4,0,ii)
									if not rsAssignment.EOF then
										DateTransfer=cdate(rsAssignment("DateTransfer")) 
										blnLink= (not rsAssignment("fgDelete")) and (DateTransfer<=strCurrentDate) and (rsAssignment("fgActivate")) and (strCurrentDate>=date()-dateLimit) AND strCurrentDate>strDateLock
										if blnLink then
											if not RecordsetFilter(strCurrentDate) then
												strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this task&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & formatnumber(dblHour,1) & "</a>"
											end if
										end if
									end if
								end if
											
							Else
								If varTimesheet(intDayCol-2, 0, ii) = "N" Then
									strHour = "&nbsp;"
								Else
									strHour = "&nbsp;"
									if (strCurrentDate>=date()-dateLimit) AND strCurrentDate>strDateLock then
										rsAssignment.MoveFirst								 
										rsAssignment.Find "AssignmentID = " & cdbl(varTimesheet(intDayCol-4,0,ii))
										if not rsAssignment.EOF then
											DateTransfer=cdate(rsAssignment("DateTransfer")) 
											blnLink= (not rsAssignment("fgDelete")) and (DateTransfer<=strCurrentDate) and (rsAssignment("fgActivate")) 
											if blnLink then
												if not RecordsetFilter(strCurrentDate) then
													strHour = "<a class='c' href=" & strURLSetHour & ">" & "--" & "</a>"
												end if
											end if
										end if
									end if
											
								End If	
							End If	
						End If
						
						intWeekDay = WeekDay(strFirstDay+(kk-1))
						strColorOpt = ""
						Select Case intWeekDay
							Case 1
								strColorOpt = SUNCOLOR
							Case 7
								strColorOpt = SATCOLOR
						End Select
						If isHoliday(kk) >= 0 Then
							strColorOpt = HOLIDAYCOLOR
						End If	
	%>                        
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal" ><%=strHour%></td>
<%
					Next
					If CDbl(varTimesheet(intDayCol-5, 0, ii)) > 0 Then
						dblTotal = formatnumber(varTimesheet(intDayCol-5, 0, ii),1)
					Else
						dblTotal = "&nbsp;"
					End If		
%>  
                <td bgcolor="#FFF2F2" align="right" class="blue"><b><%=dblTotal%></b>&nbsp;</td>
              </tr>
<%
				End If
			Else
%>                      
              <tr> 
                <td width="1%" height="20" bgcolor="#FFC6C6" class="white">&nbsp;</td>
                <td width="20%" bgcolor="#FFF2F2" class="blue-normal">&nbsp;</td>
<%
				For kk = 1 To intDayNum
					intWeekDay = WeekDay(strFirstDay+(kk-1))
					strColorOpt = ""
					Select Case intWeekDay
						Case 1
							strColorOpt = SUNCOLOR
						Case 7
							strColorOpt = SATCOLOR
					End Select
					If isHoliday(kk) >= 0 Then
						strColorOpt = HOLIDAYCOLOR
					End If	
%>                        
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal" >&nbsp;</td>
<%
				Next
%>			
                <td bgcolor="#FFF2F2" align="right" class="blue-normal">&nbsp;</td>
              </tr>
<%
			End If
		Next	
	End If
	objDatabase.dbDisConnect()																' Disconnect to SQL database	
	Set objDatabase = Nothing	
%>           

<!--**************************** End Of Project And SubTask *******************************-->

<!--**************************** Add Sub-Task *********************************************-->

<%If strError = "" Then%>  
			  <tr>
                <td bgcolor="#FFC6C6" class="white" height="20"><img src="../../images/cross.gif" width="8" height="14"></td>
                <td bgcolor="#FFF2F2" class="blue-normal"><a href="javascript:addsub();" onMouseOver="self.status='Please click here to select project or subtask for writing timesheet';return true" onMouseOut="self.status='';return true">&nbsp;Add SubTask</a></td>
<%
	For kk = 1 To intDayNum
		intWeekDay = WeekDay(strFirstDay+(kk-1))
		strColorOpt = ""
		Select Case intWeekDay
			Case 1
				strColorOpt = SUNCOLOR
			Case 7
				strColorOpt = SATCOLOR
		End Select
		If isHoliday(kk) >= 0 Then
			strColorOpt = HOLIDAYCOLOR
		End If	
%>                        
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal" >&nbsp;</td>
<%	Next%>			
				<td bgcolor="#FFF2F2" align="right" class="blue-normal">&nbsp;</td>
			  </tr>

<!--**************************** End Of Add Sub-Task *********************************************-->
<%End If%>
<!--**************************** For Events and Others *******************************************-->
              <tr>
<%
	For ii = 0 To eRow
		If varEvent(intDayNum+2,0,ii) = -1 Or varEvent(intDayNum+2,0,ii) = -2 Or varEvent(intDayNum+2,0,ii) = -3 Then
%>
                <td height="20" colspan="2" bgcolor="#FFE1E1" class="blue"><b>&nbsp;<%=varEvent(0,0,ii)%></b></td>
<%
		Else
%>          
                <td height="20" colspan="2" class="blue-normal" bgcolor="#FFE1E1">&nbsp;<%=varEvent(0,0,ii)%></td>
<%
		End If

		For kk =1 To intDayNum
			dblHour = varEvent(kk, 0, ii) + varEvent(kk, 1, ii)

			strHour = "&nbsp;"
			
			strCurrentDate = CDate(intMonth & "/" & kk & "/" & intYear)
			strParm = CStr(ii) & "," & CStr(kk) & ",'E'"
			strURLSetHour = "javascript:sethour("& strParm & ");"
'from the first day of month to now

			If kk <= intDayCount Then
				If dblHour > 0 Then
				'Response.Write strCurrentDate & "--" & strDateLock & "****" & (strCurrentDate <=strDateLock) & "==" & trim(varEvent(0,0,ii)) & "----" & (trim(varEvent(0,0,ii)) = "General/Admin" and strCurrentDate<strDateLock)
				
					If trim(varEvent(0,0,ii)) = "Public Holiday" Or varEvent(intDayNum+2,0,ii) = -1 Or varEvent(intDayNum+2,0,ii) = -2 Or varEvent(intDayNum+2,0,ii) = -3 or strCurrentDate<date-dateLimit  Then
						strHour = formatnumber(dblHour,1)
					Else
						'Response.Write strCurrentDate & "--" & strDateLock & "****" & (cdate(strCurrentDate) <=cdate(strDateLock))
						strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this event&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & formatnumber(dblHour,1) & "</a>"	
						if (trim(varEvent(0,0,ii)) <> "General/Admin" AND trim(varEvent(0,0,ii)) <> "Personal Time") and (RecordsetFilter(strCurrentDate) or not fgEnterAnnualy) then
							strHour = formatnumber(dblHour,1)
						end if
					End If		
				Else
					If trim(varEvent(0,0,ii)) = "Public Holiday" Or varEvent(intDayNum+2,0,ii) = -1 Or varEvent(intDayNum+2,0,ii) = -2 Or varEvent(intDayNum+2,0,ii) = -3 or strCurrentDate<date-dateLimit Then					
						strHour = "&nbsp;"						
					ElseIf trim(varEvent(0,0,ii)) = "Annual Holiday" Or trim(varEvent(0,0,ii)) = "Sick Leave" Or trim(varEvent(0,0,ii)) = "Sick Leave with  certificate" Or trim(varEvent(0,0,ii)) = "Other Leave" Or trim(varEvent(0,0,ii)) = "Unpaid Leave" Or trim(varEvent(0,0,ii)) = "Time in lieu of working OT" Then					
	                    
	                    intWeekday = Weekday(strFirstDay + (kk - 1))
	                    
	                    If intWeekday = 1 Or intWeekday = 7 Or isHoliday(kk) >= 0 or strCurrentDate<date-dateLimit Then					' Can not fill in sick day (annual leave day) on Saturday or Sunday or Public Holiday 
							strHour = "&nbsp;"
						Else
							strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this event&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & "--" & "</a>"
							if RecordsetFilter(strCurrentDate) or not fgEnterAnnualy then
								strHour = "&nbsp;"
							end if
						End If
					Else
						strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this event&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & "--" & "</a>"	
						if RecordsetFilter(strCurrentDate) then
							strHour = "&nbsp;"
						end if
					End If	
				End If
'For timesheet in future				
			ElseIf (kk > intDayCount And (CInt(intMonth) = Month(Date) Or CInt(intMonth) > Month(Date)) And CInt(intYear) = Year(Date)) Or (CInt(intYear) > Year(Date)) Then								' Permits to fill in future holidays in advance

				If dblHour > 0 Then
					If varEvent(intDayNum+2,0,ii) = -1 Or varEvent(intDayNum+2,0,ii) = -2 Or varEvent(intDayNum+2,0,ii) = -3 Or trim(varEvent(0,0,ii)) = "Public Holiday" Then
						strHour = formatnumber(dblHour,1)
					ElseIf trim(varEvent(0,0,ii)) = "Annual Holiday" Or trim(varEvent(0,0,ii)) = "Sick Leave" Or trim(varEvent(0,0,ii)) = "Sick Leave with  certificate" Or trim(varEvent(0,0,ii)) = "Other Leave" Or trim(varEvent(0,0,ii)) = "Unpaid Leave" Or trim(varEvent(0,0,ii)) = "Time in lieu of working OT" Then
						strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this event&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & formatnumber(dblHour,1) & "</a>"	

						if RecordsetFilter(strCurrentDate) OR not fgEnterAnnualy then
							strHour = formatnumber(dblHour,1)
						end if
					End If
				Else	
					If trim(varEvent(0,0,ii)) = "Annual Holiday" Or trim(varEvent(0,0,ii)) = "Other Leave" Or trim(varEvent(0,0,ii)) = "Unpaid Leave" Or trim(varEvent(0,0,ii)) = "Sick Leave with  certificate"  Or trim(varEvent(0,0,ii)) = "Time in lieu of working OT" Then
						intWeekday = Weekday(strFirstDay + (kk - 1))
						If intWeekday = 1 Or intWeekday = 7 Or isHoliday(kk) >= 0 Then						' Can not fill in annual leave day on Saturday or Sunday or Public Holiday 
							strHour = "&nbsp;"
						Else	
							strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this event&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & "--" & "</a>"
							if RecordsetFilter(strCurrentDate) or not fgEnterAnnualy then
								strHour = "&nbsp;"
							end if
						End If
					End If
				End If		
				
			End If	
			
			If strError = "No data for your request." And trim(varEvent(0,0,ii)) <> "Annual Holiday" Then
				strHour = "&nbsp;"
			End If
				
			intWeekDay = WeekDay(strFirstDay+(kk-1))
			strColorOpt = ""
			Select Case intWeekDay
				Case 1
					strColorOpt = SUNCOLOR
				Case 7
					strColorOpt = "#D2DAEC"
			End Select
			If isHoliday(kk) >= 0 Then
				strColorOpt = HOLIDAYCOLOR
			End If	
%>                        
            <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#E7EBF5" <%End If%> align="center" class="blue-normal" ><%=strHour%></td>
<%
		Next
		If varEvent(intDayNum+1, 0, ii) > 0 Then
			dblTotal = formatnumber(varEvent(intDayNum+1, 0, ii),1)
		Else
			dblTotal = "&nbsp;"
		End If		
%>
				<td bgcolor="#FFE1E1" align="right" class="blue"><%=dblTotal%>&nbsp;</td>
              </tr> 
<%
	Next
%>                      
<!--**************************** End Of Events and Others *********************************************-->

<!--**************************** For Timekeeping *******************************************-->			
<%	call GetFingerprintByStaff()
	arrTimekeeping=Array("Start Time (hh:mm)","Finish Time (hh:mm)", "Duration (hours)", "<b>Daily Balance</b> (hours)", "<b>Weekly Balance</b> (hours)") 
		intColspan=0
		dblWeeklyBalance=0
		For ii=0 to UBound (arrTimekeeping)%>
			<tr>
					<td colspan="2" bgcolor="#ffc6c6" class="blue" height="20">&nbsp;<%=arrTimekeeping(ii)%></td>
					
<%
			jj=0
			
			Do while (month(arrFingerprint(0,jj))<>intMonth AND day(arrFingerprint(0,jj))<>1)
				if ii>3 then
					dblDailyBalance =cdbl(arrFingerprint(3,jj))- cdbl(arrFingerprint(4,jj))
					dblWeeklyBalance=dblWeeklyBalance + dblDailyBalance
				end if
				
'response.write  dblWeeklyBalance & "<br>"
				jj=jj+1
				
			loop

			For kk =1 To intDayNum
				intWeekDay = WeekDay(strFirstDay+(kk-1))				
				strColorOpt = ""
				Select Case intWeekDay
					Case 1
						strColorOpt = SUNCOLOR
					Case 7
						strColorOpt = "#D2DAEC"
				End Select
				If isHoliday(kk) >= 0 Then
					strColorOpt = HOLIDAYCOLOR
				End If	
%>							
<% if ii<=3 then %>
					<td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#E7EBF5" <%End If%> align="center" class="blue-normal">
						<%if ii<2 then
								response.write arrFingerprint(ii+1,jj)
						else
								if ii=2 then
									response.write Format_hhmm(cdbl(arrFingerprint(ii+1,jj)))
								else
									dblDailyBalance =cdbl(arrFingerprint(ii,jj))- cdbl(arrFingerprint(ii+1,jj))
									response.write Format_hhmm(dblDailyBalance)
									
									
								end if
						end if
						%>
					</td>
<%else
		intColspan=intColspan+1		
		dblWeeklyBalance =dblWeeklyBalance +(cdbl(arrFingerprint(3,jj))- cdbl(arrFingerprint(4,jj)))
'response.write  dblWeeklyBalance & "<br>"		
		if intWeekDay=6 or kk= intDayNum then 
%>		
			<td bgcolor="#E7EBF5" align="right" class="blue" colspan="<%=intColspan%>"><%=Format_hhmm(dblWeeklyBalance)%>&nbsp;
			</td>
			
<%		intColspan=0
		dblWeeklyBalance=0
		end if
end if
			
				jj=jj+1
			Next	%>
				<td bgcolor="#FFE1E1" align="right" class="blue">&nbsp;</td>
			</tr>		
<%		next
%>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
 
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>
<%if strFr<>"" then%>
 <table border="0" cellspacing="5" cellpadding="0" align="center" height="50" name="aa" width="200">
  <tr> 
	<td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
	  <div align="center" class="blue"><a href="javascript:backtoapproval();" class="b">Back to Approval by team</a></div>
	</td>
  </tr>
</table>
<%end if%>
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="txthidden" value="<%=intStaffID%>">
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
<input type="hidden" name="assign" value="<%=Request.Form("assign")%>">
<input type="hidden" name="row" value="">

<input type="hidden" name="txtFrApp" id="txtFrApp" value="<%=strFr%>">
<%if strFr<>"" then%>
	<input type="hidden" id="txtstaffIDs" name="txtstaffIDs" value="<%=intStaffIDs%>">
	<input type="hidden" id="txtstartDate" name="txtstartDate" value="<%=strWeekStart%>">
<%end if%>

</form>
<script language="javascript" src="../../library/library.js"></script>
<script language="javascript" src="../../library/menu.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objNewWindow;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function onLoad()
{
        loadMenus();
}

function loadMenus() 
{
var url_1 = "tms_removetask.asp?m=" + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + "&y=" + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value;

    window.myMenu1 = new Menu();

    myMenu1.addMenuItem("Update","timesheet.asp?act=U", "", "", "", "frmtms");
    if ("<%=intMonth%>" == "<%=intCurMonth%>")
	{
		myMenu1.addMenuItem("Remove", url_1, "", "", "", "frmtms");
	}	
    myMenu1.menuHiliteBgColor = "#617DC0";
	myMenu1.menuItemWidth = 100;
	myMenu1.menuItemHeight = 20;
	myMenu1.writeMenus();
}

function menufunctions(intAssign, intRow)
{
	window.document.frmtms.assign.value = intAssign;
	window.document.frmtms.row.value = intRow;
	window.showMenu(window.myMenu1);
}

function viewtms()
{
	var URL;

	window.document.frmtms.M.value = window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value;
	window.document.frmtms.Y.value = window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value

	URL = "timesheet.asp?act=vmya";

	window.document.frmtms.action = URL;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function backtoapproval()
{
	URL = "tms_approvalByTeam.asp";

	window.document.frmtms.action = URL;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}


function logout()
{
	var url;
	url = "../../logout.asp";
	if (ns)
		document.location = url;
	else
	{
		window.document.frmtms.action = url;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function back_menu()
{
	window.document.frmtms.action = "tms_list_staff.asp?b=1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

var objAddSubWindow, objSetHourWindow, objPrintWindow

function addsub() 
{ //v2.0
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=550,height=600,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objAddSubWindow) && (!objAddSubWindow.closed))
		objAddSubWindow.focus();	
	else 
	{
		objAddSubWindow = window.open('tms_addsubtask.asp?m=' + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + '&y=' + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value + '&act=' + '<%=strAction%>' + '&s=' + '<%=intStaffID%>', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function sethour(row, col, kind)
{
	window.status = "";
	var width=252,height=200;
	
	strFeatures = "top="+(screen.height/2-(height/2))+",left="+(screen.width/2-(width/2))+",width=" +width+",height=250,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=yes,scrollbars=no";

	if((objSetHourWindow) && (!objSetHourWindow.closed))
		objSetHourWindow.close();	

	objSetHourWindow = window.open('tms_writehour.asp?r=' + row + '&c=' + col + '&k=' + kind + '&m=' + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + '&y=' + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value + '&s=' + '<%=intStaffID%>', "MyNewWindow", strFeatures);
	objSetHourWindow.focus();

	window.status = "Opened a new browser window.";  
}

function viewdetail()
{
	window.document.frmtms.action = "tms_viewdetails.asp";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function viewleave()
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

function gopage()
{
	document.frmtms.action = "../../tools/preferences.asp";
	document.frmtms.submit();
}

function selstaff()
{
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=530,height=325,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('tms_select_staff.asp?view=t', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function printpage()
{
	window.status = "";
	
	strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=450,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";

	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();	

	objPrintWindow = window.open('tms_print_preview.asp?m=' + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + '&y=' + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value + '&intStaffID=' + '<%=intStaffID%>', "MyNewWindow", strFeatures);
	objPrintWindow.focus();

	window.status = "Opened a new browser window.";  
}

function window_onunload() 
{
	if((objAddSubWindow) && (!objAddSubWindow.closed))
		objAddSubWindow.close();
		
	if((objSetHourWindow) && (!objSetHourWindow.closed))
		objSetHourWindow.close();
		
	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();
}

//-->
</script>


<%If Request.QueryString("act") = "U" Then%>
<script language="javascript">
	addsub();	
</script>
<%End If%>

</body>
</html>
