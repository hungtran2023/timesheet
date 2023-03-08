<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/library.asp"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->
<%
	Response.Buffer = True
	
	Dim intUserID, intMonth, intYear, intWeekday, intDayNum, intDayCol, intDayCount, intRow, eRow
	Dim intTotalRow, ii, kk, intCurMonth, intWeekend, intStaffID
	Dim dblHour, dblTotal, dblOffHour, dblNormalHour, dblOverTimeHour, dblDayTotal
	Dim lngSalary, lngSalDay, lngSalTotal
	Dim strFirstDay, strParm, strURLSetHour, strColorOpt, strError, strHour, strCurrency
	Dim fgOverTime, varTimesheet, varEvent

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------

	lngSalDay1 = 0
	lngSalDay2 = 0
	lngSalBasic1 = 0
	lngSalBasic2 = 0
	dblDayTotal1 = 0
	dblDayTotal2 = 0
	lngSalTotal = 0
	
	intMonth = Request.QueryString("m")
	intYear	= Request.QueryString("y")

	intCurMonth = Month(Date)
	
	intRow		= -1
	eRow		= -1
	intDayNum	= GetDay(intMonth,intYear)				' Numbers of days in a month
	intDayCol	= intDayNum + 6
	intWeekend	= 0
	
	varEmp		= session("varEmp")
	If IsArray(varEmp) Then
		intEmp	= Ubound(varEmp,2)
	End If
'Response.Write "test:" & intEmp
'Response.End			
	Redim varOverTime1(1,-1)
	Redim varOverTime2(1,-1)
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then%>
<script LANGUAGE="javascript">
<!--
	opener.document.location = "../../message.htm";
	window.close();
	//-->
</script>
	
<%	
	End If					

	intUserID	= session("USERID")
	strFirstDay = FirstOfMonth(intMonth,intYear)		' Get the first day in a month				
	strLastDay = FirstOfMonth(intMonth,intYear) + (intDayNum -1)	' Get the last day in a month
	
	intDayCount	= curDayNum(intDayNum,strFirstDay)		' Numbers of days since the first day in month to now
	


'--------------------------------------------------
' Initialize weekday array
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	If isEmpty(session("varWeekday")) = False Then
		varWeekday = session("varWeekday")
		intNum = Ubound(varWeekday,2)
	Else
		If objDatabase.dbConnect(strConnect) Then			
			strSQL = "SELECT * FROM ATC_WeekDay ORDER BY WeekDayID"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					varWeekday = objDatabase.rsElement.GetRows
					intNum = Ubound(varWeekday,2)					
					session("varWeekday") = varWeekday
					objDatabase.closeRec
				End If
			Else
				Response.Write objDatabase.strMessage
			End If
		Else
			Response.Write objDatabase.strMessage		
		End If
	End If	

%>
<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../../timesheet.css">
<STYLE>BR.page { page-break-after: always }</STYLE>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
'--------------------------------------------------
' The function that initialize array timesheet is called 
' when session("varTimesheet")/session("varEvent") is not initialized
' or user changes month/year to view timesheet    
'--------------------------------------------------

	Dim strOutputSalary

	strCheckDate = CDate(intMonth & "/" & intDayNum & "/" & intYear)

	If intEmp >= 0 Then
'===========================================================================
'for each user
'===========================================================================
		For jj = 0 To intEmp
			'if user is checked
			If CInt(varEmp(5,jj)) = 1 Then
							
				Redim varOverTime1(1,-1)
				Redim varOverTime2(1,-1)

				intStaffID = trim(varEmp(0,jj))
				
				strError	=  tmsInitial(intStaffID,intMonth,intYear)
				
				If strError = "" Then
					varTimesheet = session("varTimesheet")		' Array stores timesheet data
					varEvent	 = session("varEvent")			' Array stores event data
				Else
					varEvent	 = session("varEvent")			' Array stores event data
				End If

				call AnalyseOT(intMonth,intYear,varEvent)

				If isarray(varTimesheet) Then intRow = Ubound(varTimesheet,3)
				
				If isarray(varEvent) Then	eRow = Ubound(varEvent,3)
		
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------
				strTitle	= "<b>" & varEmp(1,jj) & "</b> - " & varEmp(3,jj)

				strTitle1	= "Salary sheet " & intMonth & "/" & intYear 
				strFunction = "<div align='center'>" & strTitle1 & "</div>"
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
				Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../../tms/templates/template1/")			
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

						Response.Write(arrPageTemplate(0))
%>
						<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
						<tr> 
						  <td valign="top">
						    <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
						      <tr> 
						        <td bgcolor="#8FA4D3"> 
						          <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
						            <tr> 
						              <td colspan="2" rowspan="2" class="white" bgcolor="#617DC0"><div align="center"> <b>Project </b> </div></td>
						              <td colspan="<%=intDayNum%>" class="white" align="center" bgColor="#617DC0">
											<table width="100%" border="0" cellspacing="0" cellpadding="0" class="blue-normal">
											  <tr>
											    <td class="white" align="center"><b>Date</b></td>
											  </tr>
											</table></td>      
						              <td rowspan="2" class="white" bgcolor="#617DC0"><div align="center"><b>Total</b></div></td>
						            </tr>
						            <tr bgcolor="#617DC0">
<%										For kk=1 To intDayNum
												intWeekDay = WeekDay(strFirstDay+(kk-1))%>
										        <td width="19"><div align="center" class=<%If(intWeekDay=1 or intWeekDay=7)then%>"holiday"<%else%>"white"<%end if%>><b><%=kk%></b></div></td>
<%						 				Next%> 
									</tr>
					  
<!--**************************** For Project And SubTask *********************************************-->		
<%

									intTotalRow = intRow
									If intTotalRow <= 5 Then intTotalRow = 5
									For ii = 0 To intTotalRow
										If ii <= intRow Then
										
											If varTimesheet(intDayCol-1,0,ii) = 0 Then
												' This row stores subtask name
												If trim(varTimesheet(intDayCol-2,0,ii)) = "S" Then
													strTitle=varTimesheet(intDayCol-3,0,ii)
													strLable=showlabel(varTimesheet(intDayCol-3,0,ii))
												else
													strTitle=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)
													strLable=showlabel(varTimesheet(0,0,ii))
												end if
%>
									<tr> 
									    <td width="8"><img src="../../../images/cross.gif" width="8" height="14"></td>
										<td width="118" class="blue" bgcolor="#FFF2F2">
												<a href="javascript:void(0);" title="<%=strTitle%>" onMouseOver="self.status=&quot;<%=strTitle%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=strLable%></b></a>
										</td>
<%
												For kk = 1 To intDayNum
													dblHour = varTimesheet(kk, 0, ii) + varTimesheet(kk, 1, ii)
													strHour = "&nbsp;"					
													If dblHour > 0 Then	strHour = formatnumber(dblHour,1)
													
													intWeekDay = WeekDay(strFirstDay+(kk-1))
													strColorOpt = "#FFFFFF"
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
						                <td bgcolor="<%=strColorOpt%>" align="center" class="blue-normal" width="19"><%=strHour%></td>
<%
												Next
												dblTotal = "&nbsp;"
												If CDbl(varTimesheet(intDayCol-5, 0, ii)) > 0 Then	dblTotal = formatnumber(varTimesheet(intDayCol-5, 0, ii),1)
%>  
						                <td bgcolor="#FFF2F2" align="right" class="blue"><%=dblTotal%>&nbsp;</td>
									</tr>
<%
											End If
										Else
%>                      
									<tr> 
										<td width="8" bgcolor="#FFC6C6" class="white">&nbsp;</td>
										<td width="118" bgcolor="#FFF2F2" class="blue-normal">&nbsp;</td>
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
										<td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal" width="19">&nbsp;</td>
<%
											Next
%>			
										<td bgcolor="#FFF2F2" align="right" class="blue-normal">&nbsp;</td>
									</tr>
<%
										End If
									Next
										
%>  
<!--**************************** End Of Project And SubTask **************************************-->
  
<!--**************************** For Events and Others *******************************************-->
 <%
									For ii = 0 To eRow
										strClass="blue-normal"
										If varEvent(intDayNum+2,0,ii) = -1 Or varEvent(intDayNum+2,0,ii) = -2 Or varEvent(intDayNum+2,0,ii) = -3 Then strClass="blue"
%>
										<td colspan="2" bgcolor="#FFE1E1" class="<%=strClass%>">&nbsp;<%=varEvent(0,0,ii)%></td>
<%
										For kk =1 To intDayNum
											dblHour = varEvent(kk, 0, ii) + varEvent(kk, 1, ii)
											strHour = "&nbsp;"
											If dblHour > 0 Then	strHour = formatnumber(dblHour,1)
											
											intWeekDay = WeekDay(strFirstDay+(kk-1))
											strColorOpt = "#E7EBF5"
											Select Case intWeekDay
												Case 1
													strColorOpt = SUNCOLOR
												Case 7
													strColorOpt = "#D2DAEC"
											End Select
											If isHoliday(kk) >= 0 Then
												strColorOpt = HOLIDAYCOLOR
											End If%>                        
								        <td bgcolor="<%=strColorOpt%>" align="center" class="blue-normal" width="19"><%=strHour%></td>
<%										Next
										dblTotal = "&nbsp;"
										If varEvent(intDayNum+1, 0, ii) > 0 Then dblTotal = formatnumber(varEvent(intDayNum+1, 0, ii),1)
%>
										<td bgcolor="#FFE1E1" align="right" class="blue"><%=dblTotal%></td>
									</tr> 
<%
								Next
								
%>  
<!--**************************** End Of Events and Others *********************************************-->
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
	arrSalaryStatus =GetSalaryStatus(intStaffID,strFirstDay,strLastDay)
	
	if not IsEmpty(arrSalaryStatus) then
%>
<table width="780" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td>
		<%
			strTemplate=GenerateSalary(intStaffID,strFirstDay,strLastDay,arrSalaryStatus,dblGrantBasic,dblGrantOT)
			
			Response.Write(strTemplate)			
			call Write_summary("templates/template1/",dblGrantBasic,dblGrantOT)

		
		%>
		</td>
	</tr>
</table>
	<%End if%>  
<BR CLASS=page>
       
<%				
			End If	
			session("varTimesheet")=empty		' Array stores timesheet data
			session("varEvent")=empty			' Array stores event data
		Next
	End If		
%>
</body>
</html>
