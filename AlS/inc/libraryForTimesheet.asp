<SCRIPT language="VBScript" RUNAT="SERVER">

Const strExRateNight=0.3

'****************************************************************
'Get Salary array
'****************************************************************
Function GenerateSalary(byval StaffID, ByVal strFirstDay,ByVal strLastDay, Byval arrSalaryStatus,dblGrantBasic,dblGrantOT)
	
	dim strTemplate,strSalary,strOverTime,strFrom,strTo	
	dim dblNetSalary,dblRateHourperMonth,dblRateHourperYear,dblUnpaidDays,dblBasicPay,dblNorTotal,dbloffTotal,dblholTotal,dblOTPay
	dim ii,intNumWorkDays
	
	
	dblGrantBasic=0
	dblGrantOT=0
	
	intNumWorkDays= NumWorkingDate(strFirstDay,strLastDay)
	strTo=strLastDay
	strTemplate=""

	for ii=0 to ubound(arrsalaryStatus,2)
		
		strSalary = ReadSalaryTemplate("templates/template1/","main/ats_salarydetail.htm")
		
		'---------For Net salary and Currency-------------
		dblNetSalary=cdbl(decode(arrsalaryStatus(2,ii),128))
		strSalary=Replace(strSalary,"@@netsalary",formatnumber(dblNetSalary,0))		
		if NOT isnull(arrsalaryStatus(3,ii)) then 
			strSalary=Replace(strSalary,"@@strCurrent",arrsalaryStatus(3,ii))
		else
			strSalary=Replace(strSalary,"@@strCurrent","VND")
		end if
		
		'---------For From and to-------------
		strSalaryDate=arrsalaryStatus(0,ii)
		if Cdate(strSalaryDate)>cdate(strFirstDay) then
			strFrom = strSalaryDate
		else
			strFrom = strFirstDay
		end if				
		strSalary=Replace(strSalary,"@@strFrom",ConvertDate(cdate(strFrom)))
		strSalary=Replace(strSalary,"@@strTo",ConvertDate(cdate(strTo)))
		
		'---------Rate per date and per hour-------------
		dblRateHourperMonth=formatnumber(dblNetSalary/(intNumWorkDays * cdbl(arrsalaryStatus(1,ii))),2)
		dblRateHourperYear=formatnumber((dblNetSalary * cint(NumMonthPerYear))/NumWorkingHour(arrsalaryStatus(1,ii),strTo),2)
		
		strSalary=Replace(strSalary,"@@rateDate",formatnumber(dblRateHourperMonth,2))				
		
		'---------SalaryDays and unpaidDays--------------
		dblUnpaidDays=NumUnpaidDays(StaffID,strFrom,strTo)
		
		call GetSumHour(StaffID,strFrom,strTo,dblWorkingHour,dblOTNormal,dblOTDayoff,dblOTHoliday)
		
		strSalary=Replace(strSalary,"@@numWorkDays",formatnumber(dblWorkingHour,2))
		strSalary=Replace(strSalary,"@@numUnpaidDays",formatnumber(dblUnpaidDays,2))
		strSalary=Replace(strSalary,"@@UnpaidValue",formatnumber(dblUnpaidDays * dblRateHourperMonth * (-1),0))
		
		'If start date =the first day of month then BasicPay=Net salary
		if dblWorkingHour= intNumWorkDays * cdbl(arrsalaryStatus(1,ii)) then
			dblBasicPay=dblNetSalary
		else
			dblBasicPay=dblWorkingHour * dblRateHourperMonth
		end if
			
		strSalary=Replace(strSalary,"@@SalaryValue",formatnumber(dblBasicPay,0))
		strSalary=Replace(strSalary,"@@BasicTotal",formatnumber(dblBasicPay-(dblUnpaidDays * dblRateHourperMonth),0))
		
		dblGrantBasic=dblGrantBasic + dblBasicPay-(dblUnpaidDays * dblRateHourperMonth)
		
		'---------Overtime pay--------------
		if (arrsalaryStatus(4,ii)=true) and (dblOTNormal(1) + dblOTNormal(2)>0 or dblOTDayoff(1) + dblOTDayoff(2)>0 or dblOTHoliday(1) + dblOTHoliday(2)>0 ) then
			'Get total for each row
			call GetTotalOTPay(dblRateHourperYear,dblOTNormal)
			call GetTotalOTPay(dblRateHourperYear,dblOTDayoff)
			call GetTotalOTPay(dblRateHourperYear,dblOTHoliday)			

			strOverTime = ReadOverTimeTemplate("templates/template1/",dblOTNormal,dblOTDayoff,dblOTHoliday,strFrom)

			strOverTime=Replace(strOverTime,"@@rateHour",formatnumber(dblRateHourperYear,2))

			dblOTPay=0
			for kk=3 to 4 
				dblOTPay= dblOTPay + dblOTNormal(kk) + dblOTDayoff(kk) + dblOTHoliday(kk)
			next
			strOverTime=Replace(strOverTime,"@@OverTimeTotal",formatnumber(dblOTPay,0))

			dblGrantOT=dblGrantOT + dblOTPay
		end if

		strTemplate = strSalary & strOverTime & strTemplate
		strTo=cdate(strFrom)- 1
	Next 
	 
	GenerateSalary=strTemplate

End Function

'****************************************************************
'Get Salary array for report
'****************************************************************
Sub GetSalaryStaffForReport(byval StaffID, ByVal strFirstDay,ByVal strLastDay, Byval arrSalaryStatus,dblNet,dblSal,dblUnpaid,dblOverTime,dblProbation)
	
	dim strFrom,strTo,strCurrentcy
	dim dblWorkingHour,dblOTNormal,dblOTDayoff,dblOTHoliday,dblRateNor,dblRateDayoff,dblRateHoliday,jj

	dim dblRateHourperMonth,dblRateHourperYear,dblOTPay
	dim ii,intNumWorkDays
					 
	dblSal=0
	dblUnpaid=0
	dblOverTime=0
	dblProbation=0
	
	intNumWorkDays= NumWorkingDate(strFirstDay,strLastDay)
	strTo=strLastDay
	
	dblNet=0
	
	for ii=0 to ubound(arrsalaryStatus,2)		
		'---------For From and to-------------
		strFrom = strFirstDay
		if Cdate(arrsalaryStatus(0,ii))>cdate(strFirstDay) then	strFrom = arrsalaryStatus(0,ii)
		
		'----------Get Hour and rate for OT --------------
		call GetSumHour(StaffID,strFrom,strTo,dblWorkingHour,dblOTNormal,dblOTDayoff,dblOTHoliday)
		
		'---------For Net salary and Currency-------------
		dblNetSalary=cdbl(decode(arrsalaryStatus(2,ii),128))
		strCurrentcy=arrsalaryStatus(3,ii)
		
		'---------Rate per date and per hour-------------
		dblRateHourperMonth=dblNetSalary/(intNumWorkDays * cdbl(arrsalaryStatus(1,ii)))
		dblRateHourperYear=(dblNetSalary * cint(NumMonthPerYear))/NumWorkingHour(arrsalaryStatus(1,ii),strTo)
		
		'---------Calculate NET salary base on strFrom andstrTo -------
		dblNetTemp= (dblNetSalary/intNumWorkDays) * NumWorkingDate(strFrom,strTo)
		dblNet=dblNet + dblNetTemp

		'---------SalaryDays and unpaidDays--------------
		dblUnpaid= dblUnpaid + (NumUnpaidDays(StaffID,strFrom,strTo) * dblRateHourperMonth)

		'If start date =the first day of month then BasicPay=Net salary		
		if dblWorkingHour= intNumWorkDays * cdbl(arrsalaryStatus(1,ii)) then
			dblSal=dblNetSalary
		else
			dblSal= dblSal + (dblWorkingHour * dblRateHourperMonth)
		end if			

		'---------For probation-------------
		if arrsalaryStatus(8,ii) then dblProbation = dblProbation + dblNetTemp

		'---------Overtime pay--------------
		if (arrsalaryStatus(4,ii)=true) and (dblOTNormal(1) + dblOTNormal(2)>0 or dblOTDayoff(1) + dblOTDayoff(2)>0 or dblOTHoliday(1) + dblOTHoliday(2)>0 ) then
			'Get total for each row
			call GetTotalOTPay(dblRateHourperYear,dblOTNormal)
			call GetTotalOTPay(dblRateHourperYear,dblOTDayoff)
			call GetTotalOTPay(dblRateHourperYear,dblOTHoliday)
	
			dblOTPay=0		
			for jj=3 to 4 
				dblOTPay= dblOTPay + dblOTNormal(jj) + dblOTDayoff(jj) + dblOTHoliday(jj)
			next
			dblOverTime=dblOverTime + dblOTPay
		end if	

		strTo=cdate(strFrom)- 1
	Next 

End Sub

'****************************************************************
'Get Salary template
'****************************************************************
function ReadSalaryTemplate(ByVal strTemplatePath,byval strFilename)

	Dim objFile, objTStream, strPathFile, strPageBaseText, strTemplateLocation
	strTemplateLocation=strTemplatePath

	'--------------------------------------------------
	' Loop through salary detail template file content
	'--------------------------------------------------
		
	Set objFile		= Server.CreateObject("Scripting.FileSystemObject")
	strPathFile		= Server.MapPath(strTemplateLocation & strFilename)
	Set objTStream	= objFile.OpenTextFile (strPathFile, 1, False, False)
	strPageBaseText=""
	While Not objTStream.AtEndOfStream
		strPageBaseText = strPageBaseText & objTStream.ReadLine & vbcrlf
	Wend
	Set objTStream = Nothing
	
	ReadSalaryTemplate = strPageBaseText
end function

'****************************************************************
'Get overtime template
'****************************************************************
function ReadOverTimeTemplate(ByVal strTemplatePath,ByVal dblNormal,ByVal dblDayoff,ByVal dblHoliday, byval strDate)

	dim strPageBaseText
	
	strPageBaseText=ReadSalaryTemplate(strTemplatePath,"main/ats_overtimeDetail.htm")
	
	if cdbl(dblNormal(1))>0 or cdbl(dblNormal(2))>0 then 
		strHtml=GeneralPartofOT(strDate,"&nbsp;&nbsp;Working day",dblNormal)
		strPageBaseText = Replace(strPageBaseText,"#For Normal-->",strHtml)
	end if
	
	if cdbl(dblDayoff(1))>0 or cdbl(dblDayoff(2))>0 then 
		strHtml=GeneralPartofOT(strDate,"&nbsp;&nbsp;Weekend",dblDayoff)
		strPageBaseText = Replace(strPageBaseText,"#For Dayoff-->",strHtml)
	end if
	
	if cdbl(dblHoliday(1))>0 or cdbl(dblHoliday(2))>0 then 
		strHtml=GeneralPartofOT(strDate,"&nbsp;&nbsp;Public Holiday",dblHoliday)
		strPageBaseText = Replace(strPageBaseText,"#For Holiday-->",strHtml)
	end if
	
	ReadOverTimeTemplate = strPageBaseText
end function

'****************************************************************
'Generate part of OT
'****************************************************************
Function GeneralPartofOT(byval strDate,byval strLable,byval dblOT)
	
	dim strHtml
	strHtml=""

	if cdbl(dblOT(1))>0 or cdbl(dblOT(2))>0 then 
	
		if cdate(strDate)<cdate("1-July-2003") then
			strHtml=HTMLOverTime(strLable,dblOT(1),dblOT(0),"@@norTotal")
		else
			strHtml="--><tr><td bgcolor='#eceff5' class='blue-small' colspan='5'>&nbsp;&nbsp;" & strLable & "</td></tr>"
			
			if cdbl(dblOT(1))>0 then strHtml=strHtml & HTMLOverTime("Normal Rate",dblOT(1),dblOT(0),dblOT(3))
			if cdbl(dblOT(2))>0 then strHtml=strHtml & HTMLOverTime("Night Rate",dblOT(2),dblOT(0)+ cdbl(strExRateNight), dblOT(4))
		end if
	end if
	GeneralPartofOT=strHtml
End Function
'****************************************************************
' Get HTMLHoliday
'****************************************************************
function HTMLOverTime(byval strItemName,byval strValue,byval strOTRate,byval strTotalValue)
	dim strHtml
	strHtml="<tr>"	
	strHtml=strHtml & "<td bgcolor='#ffffff' class='blue-normal' align='right'>" & strItemName & "&nbsp;&nbsp;</td>"
	strHtml=strHtml & "<td bgcolor='#ffffff' align='middle' class='blue-normal'>" & strValue & "</td>"
	strHtml=strHtml & "<td bgcolor='#ffffff' align='middle' class='blue-normal'>" & strOTRate & "</td>"
	strHtml=strHtml & "<td bgcolor='#ffffff' align='center' class='blue-normal'>@@rateHour</td>"
	strHtml=strHtml & "<td bgcolor='#ffffff' align='right' class='blue-normal'>" & strTotalValue & "</td>"
	strHtml=strHtml & "</tr>"
	HTMLOverTime=strHtml
end function

'****************************************************************
' Get salary status	
'****************************************************************
Function GetSalaryStatus(ByVal strSaffID,ByVal strFirst, ByVal strLast)

	strConnect = Application("g_strConnect")	' Connection string 				
	Set objDatabase = New clsDatabase 
	
	If objDatabase.dbConnect(strConnect) Then			
		strSQL="GetSalaryDate " & strSaffID & ",'" & strFirst & "','" & strLast & "'"

		If (objDatabase.runQuery(strSQL)) Then
		
			If objDatabase.noRecord = False Then				
				arrSal = objDatabase.rsElement.GetRows
				objDatabase.closeRec
			End If
		Else
			strError = objDatabase.strMessage
		End If
	Else
		strError = objDatabase.strMessage
	End If
	Set objDatabase = Nothing
	
	GetSalaryStatus = arrSal
	
End Function

'****************************************************************
' Get salary status	
'****************************************************************
Function GetWeekdayDetail(byval strTo)

	strConnect = Application("g_strConnect")						' Connection string 				
	Set objDatabase = New clsDatabase 
	
	If objDatabase.dbConnect(strConnect) Then			
		strSQL="sp_GetWeekdayDetail '" &  strTo & "'"

		If (objDatabase.runQuery(strSQL)) Then
		
			If objDatabase.noRecord = False Then				
				arrWeekday = objDatabase.rsElement.GetRows
				objDatabase.closeRec
			End If
		Else
			strError = objDatabase.strMessage
		End If
	Else
		strError = objDatabase.strMessage
	End If
	Set objDatabase = Nothing
	
	GetWeekdayDetail = arrWeekday
	
End Function

'****************************************************************
'Format dd/mm/yyyy
'****************************************************************
Function ConvertDate(ByVal strDate)
	
	dim strDay,strMonth
	If Day(strDate) < 10 Then
		strDay = "0" & Day(strDate)
	Else
		strDay = Day(strDate)
	End If
	If Month(strDate) < 10 Then
		strMonth = "0" & Month(strDate)
	Else
		strMonth = Month(strDate)
	End If  	
	strDate = strDay & "/" & strMonth & "/" & year(strDate)
	
	ConvertDate=strDate
	
End Function

'****************************************************************
'Number working days from strFrom and strTo
'****************************************************************
Function NumWorkingDate(ByVal strFrom,byval strTo)
	
	dim intReturn,strID
	intReturn=0
	
	strID=strFrom
	while (cdate(strID)<=cdate(strTo))
		if (WeekDay(strID)<>1) and (WeekDay(strID)<>7) then intReturn=intReturn+1
		strID=cdate(strID) + 1
	Wend
	
	NumWorkingDate = intReturn
	
End Function

'****************************************************************
'Number working hours per year
'****************************************************************
Function NumWorkingHour(ByVal hourPerDate,byval strTo)
	dim ii
	intWdayPerWeek=0	
	arrWeekdayDetail=GetWeekdayDetail(strTo)
	
	if not isempty(arrWeekdayDetail) then
		for ii=0 to ubound(arrWeekdayDetail,2)			
			if arrWeekdayDetail(3,ii)=false then intWdayPerWeek=intWdayPerWeek + 1
		next		
		
	end if	
	NumWorkingHour= cint(intWdayPerWeek) * cdbl(hourPerDate) * cint(NumWeeksPerYear)
	
End Function

'****************************************************************
'Number unpaid hours
'****************************************************************
Function NumUnpaidDays(byval StaffID,byval strFrom,byval strTo)
	
	dim myCmd,objDatabase
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "GetUnpaidHour"
		Set myParam = myCmd.CreateParameter("staffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("startD",adDate,adParamInput,10)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("endD",adDate,adParamInput,10)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("unpaidHours",adVarChar,adParamOutput,11)
		myCmd.Parameters.Append myParam
					
		myCmd("staffID") = StaffID
		myCmd("startD") = cdate(strFrom)
		myCmd("endD") = cdate(strTo)
		
		myCmd.Execute
		
		'Response.Write cdbl(myCmd("unpaidHours"))/cdbl(hourPerDate)
		
		NumUnpaidDays=cdbl(myCmd("unpaidHours"))
	end if
	
	set myCmd=nothing
	Set objDatabase=nothing
	
End Function

'****************************************************************
'Get sum working hour and over time 
'****************************************************************
Sub GetSumHour(byval StaffID,byval strFrom,byval strTo,dblWorkingHour,dblOTNormal,dblOTDayoff,dblOTHoliday)
	
	'dim myCmd,objDatabase
	dim ii, arrATS
	dblWorkingHour=0
	redim dblOTNormal(4)
	redim dblOTDayoff(4)
	redim dblOTHoliday(4)
	
	for ii=0 to 4
		dblOTNormal(ii)=0
		dblOTDayoff(ii)=0
		dblOTHoliday(ii)=0
	next
	
	arrATS=GetarrATS(StaffID,strFrom,strTo)

	if not isempty(arrATS) then
		for ii=0 to ubound(arrATS,2)
			if isnull(arrATS(3,ii)) then arrATS(3,ii)=0
			dblWorkingHour=dblWorkingHour + cdbl(arrATS(2,ii))
			'For dayoff
			if arrATS(6,ii)=true then
				dblOTDayoff(0)=cdbl(arrATS(5,ii))
				dblOTDayoff(1)=dblOTDayoff(1) + cdbl(arrATS(7,ii))
				dblOTDayoff(2)=dblOTDayoff(2) + cdbl(arrATS(8,ii))
			else
				if not isnull(arrATS(4,ii)) then
					dblOTHoliday(0)=cdbl(arrATS(4,ii))
					dblOTHoliday(1)=dblOTHoliday(1) + cdbl(arrATS(7,ii))
					dblOTHoliday(2)=dblOTHoliday(2) + cdbl(arrATS(8,ii))
				else
					dblOTNormal(0)= cdbl(arrATS(5,ii))
					dblOTNormal(1)=dblOTNormal(1) + cdbl(arrATS(7,ii))
					dblOTNormal(2)=dblOTNormal(2) + cdbl(arrATS(8,ii))
				end if
			end if		
		next	
	end if	
End sub

'****************************************************************
'Get total payment of over time 
'****************************************************************
Sub GetTotalOTPay(byval dblRateHourperYear, byref dblOT)

	dblOT(3)=formatnumber(dblOT(0) * dblOT(1) * dblRateHourperYear,0)
	dblOT(4)=formatnumber((dblOT(0) + cdbl(strExRateNight)) * dblOT(2) * dblRateHourperYear,0)

End Sub

'****************************************************************
'Get sum working hour and over time 
'****************************************************************
function GetarrATS(byval StaffID,byval strFrom,byval strTo)
	
	dim myCmd,objDatabase
	
	strConnect = Application("g_strConnect")	' Connection string 				
	Set objDatabase = New clsDatabase 
	
	If objDatabase.dbConnect(strConnect) Then			
	'Get hour from Timesheet for salary		
		strSQL="GetTimesheetHour " & StaffID & ",'" & strFrom & "','" & strTo & "'"
		If (objDatabase.runQuery(strSQL)) Then
		
			If objDatabase.noRecord = False Then				
				arrATS = objDatabase.rsElement.GetRows
				objDatabase.closeRec
			End If
		Else
			strError = objDatabase.strMessage
		End If
	Else
		strError = objDatabase.strMessage
	End If
	Set objDatabase = Nothing
	
	GetarrATS=arrATS
End function

'****************************************************************
'Get sum working hour and over time 
'****************************************************************
function GetStaffForSalarySheet(byval strSql)
	
	dim myCmd,objDatabase,arrReturn
	
	strConnect = Application("g_strConnect")	' Connection string 				
	Set objDatabase = New clsDatabase 
	
	If objDatabase.dbConnect(strConnect) Then			
		If (objDatabase.runQuery(strSql)) Then		
			If objDatabase.noRecord = False Then				
				arrReturn = objDatabase.rsElement.GetRows
				objDatabase.closeRec
			End If
		Else
			strError = objDatabase.strMessage
		End If
	Else
		strError = objDatabase.strMessage
	End If
	Set objDatabase = Nothing
	
	GetStaffForSalarySheet=arrReturn
End function

'****************************************************************
'Get salary Sheet
'****************************************************************
Function GetSalarySheet(byval arrStaff,ByVal strFirst, ByVal strLast,idDirect,arrSheet)
	dim idxStart,ii,arrStatus
	dim dblSalTotal,dblUnpaidTotal,dblOTTotal,dblNetTotal,dblOTPreTotal
	
	dblSalTotal=0
	dblUnpaidTotal=0
	dblOTTotal=0
	dblNetTotal=0
	dblOTPreTotal=0
	
	idxStart=Ubound(arrSheet,2) + 1
	Redim Preserve arrSheet(13, idxStart + ubound(arrStaff,2) + 1)
	
	'Calulate salary for each staff
	for ii=0 to ubound(arrStaff,2)

		arrSheet(0,idxStart)=ii + 1	'No.
		'
		arrSheet(1,idxStart)=arrStaff(1,ii) 'Fullname
				
		arrSheet(7,idxStart)=arrStaff(3,ii) 'Department
		arrSheet(11,idxStart)=arrStaff(2,ii) 'Taxtype
		'Staff ID
		arrSheet(13,idxStart)=arrStaff(4,ii)	'No.
		
		'Calculate NET salary and working hour
		arrStatus=GetSalaryStatus(arrStaff(0,ii),strFirst,strLast)
		
		if not IsEmpty(arrStatus) then
			'
			arrSheet(2,idxStart)="&nbsp;" & arrStatus(6,0) 'Bank Account 
			arrSheet(8,idxStart)=arrStatus(7,0) 'Bank Detail
			'arrSheet(9,idxStart)=arrStatus(8,0) 'Probation
			
			call GetSalaryStaffForReport(cint(arrStaff(0,ii)),strFirst,strLast,arrStatus,arrSheet(3,idxStart),arrSheet(4,idxStart),arrSheet(5,idxStart),arrSheet(6,idxStart),arrSheet(9,idxStart))
			
			dblNetTotal=dblNetTotal + cdbl(arrSheet(3,idxStart))
			dblSalTotal=dblSalTotal + cdbl(arrSheet(4,idxStart))
			dblUnpaidTotal=dblUnpaidTotal + cdbl(arrSheet(5,idxStart))
	
			dblOTTotal=dblOTTotal + cdbl(arrSheet(6,idxStart))
			
			'--------------------------------------
			'Add OT of last month to salary sheet
			'--------------------------------------
			Call getOTOflastMonth(arrStaff(0,ii),strFirst,strLast,arrSheet(10,idxStart))
			'if arrStaff(0,ii) =609 then Response.Write arrSheet(10,ii) & "--" & ii
			dblOTPreTotal=dblOTPreTotal + cdbl(arrSheet(10,idxStart))
		end if		
		idxStart=idxStart+1
	next 
	
	arrSheet(0,idxStart)=-1
	arrSheet(1,idxStart)=IIF(idDirect=0,"Direct Staff- Total","Indirect Staff - Total")
	arrSheet(3,idxStart)=dblNetTotal
	arrSheet(4,idxStart)=dblSalTotal
	arrSheet(5,idxStart)=dblUnpaidTotal
	arrSheet(6,idxStart)=dblOTTotal
	arrSheet(10,idxStart)=dblOTPreTotal

End function

'****************************************************************
'Write summary salary
'****************************************************************
Sub Write_summary(byval strTemplatePath,byval dblBasic,byval dblOT)
	dim strSummary
	strSummary=ReadSalaryTemplate(strTemplatePath,"main/ats_salsummary.htm")
	
	strSummary=Replace(strSummary,"@@grantBasic",FormatNumber(dblBasic,0))
	strSummary=Replace(strSummary,"@@grantOT",FormatNumber(dblOT,0))
	strSummary=Replace(strSummary,"@@grantTotal",FormatNumber(dblOT +dblBasic,0))
	
	Response.Write strSummary
	
End sub

'****************************************************************
'Get one row for salary Report
'****************************************************************
Function rptTemplateSalPerRow()
	dim strSalaryRow
	
	strSalaryRow ="<tr bgcolor='#FFFFFF'>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='3%'>@@No</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='20%'>@@fullname</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='17%'>@@jobtitle</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='10%'>@@netSal</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='10%'>@@salPay</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='10%'>@@uppaid</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='10%' bgcolor='#FFF2F2'>@@basicSal</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='10%'>@@overTimePay</td>"
    strSalaryRow = strSalaryRow & "<td class='blue' align='center' width='10%' bgcolor='#FFF2F2'>@@Total</td></tr>"
    
    rptTemplateSalPerRow=strSalaryRow
end function

'****************************************************************
'Get one row for Total salary Report
'****************************************************************
Function rptTemplateTotalSalRow()
	
	dim strSalaryRow
	strSalaryRow="<tr bgcolor='#E7EBF5'> "
    strSalaryRow = strSalaryRow & "<td valign='top' width='50%' class='blue' align='right' colspan='4'>&nbsp;</td>"
    strSalaryRow = strSalaryRow & "<td valign='top' width='10%' class='blue' align='right'>&nbsp;</td>"
    strSalaryRow = strSalaryRow & "<td valign='top' width='10%' class='blue' align='right'>&nbsp;</td>"
    strSalaryRow = strSalaryRow & "<td valign='top' width='10%' class='blue' align='right'>&nbsp;</td>"
    strSalaryRow = strSalaryRow & "<td valign='top' width='10%' class='blue' align='right'>&nbsp;</td>"
    strSalaryRow = strSalaryRow & "<td valign='top' width='10%' class='blue' align='right'>&nbsp;</td></tr>"
    
    rptTemplateTotalSalRow=strSalaryRow
end function

'****************************************************************
'Get leave date of User
'****************************************************************
Function LeaveDateUser(byval UserID)
	Dim strSql,rsTemp
	
	strSql="SELECT LeaveDate FROM ATC_Employees WHERE StaffID=" & UserID
	call GetRecordset (strSql, rsTemp)
	
	if rsTemp is nothing then 
		LeaveDateUser = null
	else
		LeaveDateUser = rsTemp("LeaveDate")
	end if	
End Function

'****************************************************************
'Analyse OT
'****************************************************************
Function AnalyseOT(byval intMonth,byval intYear,byref varEvent)
	
	Dim numCol,numRow,ii
	
	strReturn=""
	numCol=Ubound(varEvent,1)
	numRow=Ubound(varEvent,3)
	
	Redim Preserve varEvent(numCol,4,numRow + 2)
	
	varEvent(0,0,numRow + 1)="&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Normal Rate"
	varEvent(0,0,numRow + 2)="&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Night Rate"
	
	'for Initialize normal rate	
	For ii = 1 to numCol
		'For normal rate
		varEvent(ii,0,numRow + 1) = varEvent(ii,2,numRow)
		'For night rate
		varEvent(ii,0,numRow + 2) = varEvent(ii,3,numRow)
	Next
	
End Function

'****************************************************************
' OT Get OT Of last Month
'****************************************************************
Sub GetOTOflastMonth(byval staffID,byval strFirst,byval strLast,byref dblOTLastMonth)

	dim arrStatusLastMonth
	dim arrTemp(4)
	
	'Date of last month
	strFirst= FirstOfMonth(IIF(month(strFirst)=1,12,month(strFirst)-1),IIF(month(strFirst)=1,year(strFirst)-1,year(strFirst)))
	strLast=  DateAdd("m",1,strFirst)-1

	'Calculate NET salary and working hour
	arrStatusLastMonth=GetSalaryStatus(staffID,strFirst,strLast)
		
	if not IsEmpty(arrStatusLastMonth) then		'		
		call GetSalaryStaffForReport(staffID,strFirst,strLast,arrStatusLastMonth,arrTemp(0),arrTemp(1),arrTemp(2),dblOTLastMonth,arrTemp(3))			
	end if		

End Sub

</SCRIPT>