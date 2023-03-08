<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strError, PageSize, fgRight 'view all or Not
	Dim intApproval

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
'Response.Write psize & "--" & rsSrc.RecordCount
	If Not rsSrc.EOF Then
		For i = 1 To psize	
			strApproval=""						
			if rsSrc("approved")>0 then strApproval="<img src='../../images/yes.gif'>"		
			strColor = "#FFF2F2"
			If i Mod 2 = 0 Then strColor = "#E7EBF5"
				
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' class='blue'><a href='javascript:viewtms(" & rsSrc("StaffID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;View Timesheet&quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("Fullname")) & "</a></td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("JobTitle")) & "</td>" &_
			         "<td valign='top' class='blue-normal'>" & Showlabel(rsSrc("Leader")) & "</td>" &_
			         "<td valign='top' class='blue-normal' align='center'>" & strApproval & "</td>" &_
			         "</tr>" & chr(13)
			         			
			rsSrc.MoveNext
			If rsSrc.EOF Then exit for
		Next
		
	End If
	Outbody = strOut
End Function
'--------------------------------------------------
' Adding more row
'--------------------------------------------------
sub AddRowToTimeSheet(byref arr,byref intRow)
	
	intRow = intRow + 1

	Redim Preserve arr(constWeedays + 2,intRow)
	for ii=1 to constWeedays + 2
		arr(ii,intRow)="&nbsp;"
	next
end sub
'--------------------------------------------------
' Get Project Timesheet Array
'--------------------------------------------------
sub GetProTimeSheetArray(byval userID,byval fromDate, byval toDate)
	
	dim strSql,rsProATS,numRows
	dim strCurTask,dblSubtotal
	ReDim arrProjectATS(constWeedays + 2, 0)
	
	numRows=0
	
	strSql="SELECT upper(c.ProjectID) as ProjectID,c.SubTaskID,c.TaskID,c.Parent,c.SubTaskName,a.Tdate,a.Hours,a.OverTime " & _
			"FROM " & GetUnionTimesheetSQL(fromDate,toDate) & _
			"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
			"INNER JOIN (SELECT a1.SubTaskID, a1.ProjectID, a1.SubTaskName, a1.TaskID, a1.ChainID, b1.SubTaskName AS Parent " & _
						"FROM ATC_Tasks a1 LEFT OUTER JOIN ATC_Tasks b1 ON a1.TaskID = b1.SubTaskID) c ON b.SubTaskID=c.SubtaskID " & _
			"WHERE a.EventID=1 AND a.staffID=" & userID & " AND a.Tdate BETWEEN '" & fromDate & "' AND '" & toDate & "' ORDER BY ProjectID,TaskID,SubtaskName"

	Call GetRecordset(strSql,rsProATS)
	strCurTask="#"
	if not rsProATS.EOF then
		do while not rsProATS.EOF
			if strCurTask<>rsProATS("ProjectID") & "#" & rsProATS("SubTaskName") then
				
				arrProjectATS(constWeedays + 2,numRows)="<b>" & FormatNumber(dblSubtotal,1) & "</b>"
				dblSubtotal=0
				call AddRowToTimeSheet(arrProjectATS,numRows)
				if IsNull(rsProATS("TaskID")) then
					arrProjectATS(1,numRows)="<a href='javascript:void(0);' title='" & rsProATS("SubTaskName") & "' class='c'><b>&nbsp;" & rsProATS("ProjectID") & "</b></a>"
				else

					if left(strCurTask,Instr(1,strCurTask,"#")-1)<>rsProATS("ProjectID") then
						arrProjectATS(1,numRows)="<a href='javascript:void(0);' title='" & rsProATS("Parent") & "' class='c'><b>&nbsp;" & rsProATS("ProjectID") & "</b></a>"
						call AddRowToTimeSheet(arrProjectATS,numRows)
					end if
					
					arrProjectATS(1,numRows)="<a href='javascript:void(0);' title='" & rsProATS("SubTaskName") & "' class='c'><b>&nbsp;&nbsp;&nbsp;- &nbsp; " & rsProATS("SubTaskName") & "</b></a>"
				end if
				
				strCurTask=rsProATS("ProjectID") & "#" & rsProATS("SubTaskName")
				
			end if
			arrProjectATS(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,numRows)=FormatNumber(cdbl(rsProATS("hours")) + cdbl(rsProATS("OverTime")),1)
			if Instr(1,rsProATS("ProjectID"),"01000_ATL_")>0 then
				dblTotalOtherHours=dblTotalOtherHours + cdbl(rsProATS("hours")) + cdbl(rsProATS("OverTime"))
			else
				dblTotalProHours=dblTotalProHours + cdbl(rsProATS("hours")) + cdbl(rsProATS("OverTime"))
			end if
			
			arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,2)=arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,2) + cdbl(rsProATS("hours"))
			arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,3)=arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,3) + cdbl(rsProATS("OverTime"))
			arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,1)=arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,3) + arrTotal(Weekday(cdate(rsProATS("Tdate")),vbMonday) + 1,2)
			
			dblSubtotal=dblSubtotal + cdbl(rsProATS("hours")) + cdbl(rsProATS("OverTime"))
			rsProATS.MoveNext
		loop
		arrProjectATS(constWeedays + 2,numRows)="<b>" & FormatNumber(dblSubtotal,1) & "</b>"
	end if
'For decoration	
	do while numRows<=10
		call AddRowToTimeSheet(arrProjectATS,numRows)
	loop

end sub
'--------------------------------------------------
' Get Event Timesheet Array
'--------------------------------------------------
sub GetEventTimeSheetArray(byval userID,byval fromDate, byval toDate)
	dim strSql,rsEventATS,numRows
	dim strEventName,dblSubtotal
	ReDim arrEventATS(constWeedays + 2, 0)
	
	numRows=0
	strEventName=""
	strSql="SELECT c.EventID,c.EventName,Tdate,ISNULL(hours,0) as hours,ISNULL(OverTime,0) as Overtime " & _
	"FROM ATC_Events c " & _
		"LEFT JOIN (SELECT TDate,EventID,hours,OverTime FROM " & GetUnionTimesheetSQL(fromDate,toDate)  & _
					" WHERE EventID>1 AND StaffID = " & userID & " AND Tdate BETWEEN '" & fromDate & "' AND '" & toDate & "') b ON c.EventID=b.EventID " & _
	"WHERE c.EventID<>1 ORDER BY c.EventID"
'Response.Write strSql
'Response.End
	Call GetRecordset(strSql,rsEventATS)
	if not rsEventATS.EOF then
		do while not rsEventATS.EOF
			if strEventName<>rsEventATS("EventName")then
								
				arrEventATS(constWeedays + 2,numRows)="<b>" & IIF(dblSubtotal=0,"&nbsp;",FormatNumber(dblSubtotal,1)) & "</b>"
				dblSubtotal=0
				call AddRowToTimeSheet(arrEventATS,numRows)				
				arrEventATS(1,numRows)="&nbsp;" & rsEventATS("EventName")				
				strEventName=rsEventATS("EventName")
			end if
			if not isnull(rsEventATS("Tdate")) then	
				arrEventATS(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,numRows)=FormatNumber(cdbl(rsEventATS("hours")) + cdbl(rsEventATS("OverTime")),1)
				
				if rsEventATS("EventID")=3 then
					dblTotalOtherHours=dblTotalOtherHours + cdbl(rsEventATS("hours")) + cdbl(rsEventATS("OverTime"))
				else
					dblTotalLeaveHours=dblTotalLeaveHours + cdbl(rsEventATS("hours")) + cdbl(rsEventATS("OverTime"))
				end if
				
				arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,2)=arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,2) + cdbl(rsEventATS("hours"))
				arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,3)=arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,3) + cdbl(rsEventATS("OverTime"))
				arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,1)=arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,3) + arrTotal(Weekday(cdate(rsEventATS("Tdate")),vbMonday) + 1,2)

			end if
			dblSubtotal=dblSubtotal + cdbl(rsEventATS("hours")) + cdbl(rsEventATS("OverTime"))			
			
			rsEventATS.MoveNext
		loop
		arrEventATS(constWeedays + 2,numRows)="<b>" & IIF(dblSubtotal=0,"&nbsp;",FormatNumber(dblSubtotal,1)) & "</b>"
	end if
	
	
end sub
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intDepartmentID = Request.Form("lbdepartment")
	fgSort = Request.Form("S")
	
	intApproval=Request.Form("lbApproval")
	if intApproval="" then
		intApproval=0
	else
		intApproval=cint(intApproval)
	end if
	'if Request.Form("lbApproval")<>"" then
		'intApproval=cint(Request.Form("lbApproval"))
	 
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
			strSQL = "SELECT * FROM ATC_Department ORDER BY Department"

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
' Initialize appoval timesheet records
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 


	dateToday=date()
	'the last Monday from today
	dateFrom=dateToday-(Weekday(dateToday,2) + 6)			
	'the last Sunday from today
	dateTo=dateFrom + 6

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
			Set rsStaff = Server.CreateObject("ADODB.Recordset")
			rsStaff.CursorLocation = adUseClient			' Set the Cursor Location to Client

			Set myCmd = Server.CreateObject("ADODB.Command")
			Set myCmd.ActiveConnection = objDatabase.cnDatabase
			myCmd.CommandType = adCmdStoredProc
			myCmd.CommandText = "sp_getListEmp"

			Set myParama = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
			myCmd.Parameters.Append myParama
			Set myParamb = myCmd.CreateParameter("level",adTinyInt,adParamInput)
			myCmd.Parameters.Append myParamb
			Set myParamc = myCmd.CreateParameter("strSQL", adVarChar,adParamInput, 5000)
			myCmd.Parameters.Append myParamc
			Set myParamd = myCmd.CreateParameter("fgCheck", adTinyInt,adParamInput)
			myCmd.Parameters.Append myParamd
					
			myCmd("StaffID") = session("USERID")
			myCmd("level") = 0

			strSQL = "SELECT a.StaffID, FirstName + ' '+ ISNULL(LastName,'') AS FullName, a.DepartmentID, c.JobTitle,a.DirectLeaderID,e.Fullname as Leader,ISNULL(d.staffID,0) as approved" & _
						" FROM ATC_Employees a LEFT JOIN ATC_PersonalInfo b ON a.StaffID = b.PersonID " & _
						" LEFT JOIN HR_CurrentJobtitle c ON a.StaffID = c.StaffID " & _
						" LEFT JOIN (SELECT staffID FROM ATC_TimesheetApproval WHERE DateFrom='" & dateFrom & "' AND DateTo='" & dateTo & "') d ON a.StaffID = d.StaffID " & _
						" LEFT JOIN (SELECT PersonID,FirstName + ' ' + ISNULL(MiddleName,'') + ' '+ ISNULL(LastName,'') AS FullName FROM ATC_PersonalInfo) e ON  e.PersonID=a.DirectLeaderID " & _
						" WHERE b.fgDelete = 0 AND a.fgIndirect=0"
			
			
'Response.Write(fgRight)
						
			If fgRight Then						' View all		  
				myCmd("fgCheck") = 0
			Else
				strSQL = strSQL & " AND a.StaffID " '& session("USERID")
				myCmd("fgCheck") = 1 
			End If
'Response.Write strSQL
			myCmd("strSQL") = strSQL
			
'Response.Write myCmd("StaffID") & "," & myCmd("level") & "," & myCmd("strSQL") & "," & myCmd("fgCheck")
			On Error Resume Next	
			rsStaff.Open myCmd,,adOpenStatic,adLockBatchOptimistic
			If Err.number > 0 then
				strError = Err.Description
			End If
			Err.Clear
'Response.Write rsStaff.RecordCount			
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
				Case "vpsr"											' Sort by department		
					
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
						rsStaff.Sort = "Leader ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "Leader DESC"
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
					intDepart = Request.Form("lbdepartment")
					rsStaff.Filter=""
					strFilter=""
					
					strFilter=iif(strSName <> "","FullName LIKE '%" & strSName & "%'","")
					strFilter= strFilter & iif(CInt(intDepart) <> 0,iif(strFilter<>"", " AND ","") & "DepartmentID=" & intDepart,"")
					if intApproval>0 then
						strFilter=strFilter & iif(strFilter<>"", " AND ","")
						strFilter= strFilter & iif(intApproval = 1,"approved >0","approved =0")
					end if
					rsStaff.Filter=strFilter
'Response.Write strFilter					
					'If strSName <> "" And CInt(intDepart) <> 0 Then
					'	If InStr(1,Request.Form("txtname"),"'") = 0 Then
					'		rsStaff.Filter = "FullName LIKE '%" & strSName & "%' AND DepartmentID=" & intDepart
					'	Else
					'		rsStaff.Filter = "FullName LIKE #" & strSName & "# AND DepartmentID=" & intDepart
					'	End If		
					'ElseIf strSName = "" And CInt(intDepart) <> 0 Then
					'	rsStaff.Filter = "DepartmentID=" & intDepart
					'ElseIf strSName <> "" And CInt(intDepart) = 0 Then
					'	If InStr(1,Request.Form("txtname"),"'") = 0 Then
					'		rsStaff.Filter = "FullName LIKE '%" & strSName & "%'"
					'	Else
					'		rsStaff.Filter = "FullName LIKE #" & strSName & "#"
					'	End If	
					'End If

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

					intDepartmentID = 0
					strName = ""
					strStatus = Request.Form("txtstatus")
			End Select	 
			
			strLast = Outbody(rsStaff, intPageSize)
		End If		
	End If
'==================================================================================
'For appoval individually
'==================================================================================
	dim arrProjectATS, arrEventATS,arrTotal
	intStaffID=request.form("txthidden")
	
	curDate=request.form("selecteddate")
	if curDate="" then curDate=Date()
	
	strLastSat = curDate - Weekday(curDate, vbSaturday) +1
	strNextFri = strLastSat + 6
	
	strTitle1	= "Timesheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
	
	if intStaffID<>"" then
		call GetProTimeSheetArray(intStaffID,strLastSat,strNextFri)
		call GetEventTimeSheetArray(intStaffID,strLastSat,strNextFri)
	end if

	
	
	
'response.write strLastSat & "--" & strNextFri
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
	If strChoseMenu = "" Then strChoseMenu = "AA"
	
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
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">
<head>
<meta http-equiv="Content-type" content="text/html;charset=UTF-8"/>
<meta http-equiv="Content-Language" content="en"/>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Atlas Industries Timesheet System</title>

<link href="../../bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../bootstrap/css/dataTables.bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../bootstrap/css/bootstrap-datepicker.css" rel="stylesheet" type="text/css">
<link href="../../css/timesheet.css" rel="stylesheet" >
<link href="../../css/style.css" rel="stylesheet" type="text/css">    
<style>
.modal-dialog{
   width: 60%;
   margin: auto;
}

.datepicker { font-size: 10px; }

</style>

</head>
<body data-pinterest-extension-installed="cr1.39.1">

<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
<div class="container-fluid" >
<%If strError <> "" Then%>  
	<div class="row">	
			<div class="<%if strError="Update successfull." then %>alert alert-danger<%else%>alert alert-success<%end if%>">
				<strong>Error:</strong><%=strError%>
			</div>
		</div>
<% End If%>	
	<div class="row" style="padding:20px 0px 0px 20px;"><h3>Approval register</h3> 
	From <b>Mon, <%=day(dateFrom) & "/" & month(dateFrom) & "/" & year(dateFrom)%></b> to <b>Sun,<%=day(dateTo) & "/" & month(dateTo) & "/" & year(dateTo)%></b></div>
	<div class="row">
        <div class="col-sm-6 col-sm-offset-3">
            <div id="imaginary_container"> 
			<form name="searchform" method="post" action="tms_listfor_approval.asp">
				<div class="form-group">
					<div class="input-group">
						<input type="text" name="txtSearch" id="txtSearch" onkeyup="myFunction()" class="form-control" placeholder="Filter">
						<div class="input-group-btn">
							<button type="button" id="btnFilter" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
								<span id="filterLable">By Fullname</span>
								<span class="caret"></span>
							</button>
							<ul class="dropdown-menu">
								<li><a  href="#" id="filterBy">By Manager</a></li>
							</ul>
						</div>
					</div>
				</div>
			</form>
            </div>
        </div>
	</div>
	<div class="row">
		<form id="frmList" method="post" action="tms_listfor_approval.asp">
			<div class="table-responsive">	
				<table class="table table-hover" id="tblList">
					<thead  class="thead-inverse tableheaderblue">
						<tr>
							<th>Full Name</th>
							<th>Jobtitle</th>
							<th>Department</th>
							<th>Report To</th>
							<th>Approved</th>
						</tr>
					</thead>					
					
				</table>	
				<input type="hidden" id="txthidden" name="txthidden" value="">
			</div>		
		</form>		
    </div> 
	<div class="row">
		<div class="text-center">
			<%'=Pagination()
			%>
		</div>
	</div>

</div>
 <!-- Modal HTML -->
    <div id="myModal" class="modal">
        <div class="modal-dialog">
            <div class="modal-content">
			<form id="detailform" class="form-horizontal" action="tms_listfor_approval.asp?act=save" method="post">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                    <h4 class="modal-title">Approval register</h4>
                </div>
                <div class="modal-body">   
					<div class="row">
						 <div class='col-sm-3' >
							<div class="form-group" style="padding-left:10px">
								<div class='input-group date' id='datetimepicker' style="border-style: ridge;border-width:1px;border-color:grey">
									<input type="hidden" id="selecteddate">
								</div>
							</div>
						</div>
						<div class='col-sm-9'>							
							<table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
								<tr> 
									<td colspan="2" rowspan="2" class="white" bgcolor="#617DC0"> 
										<div align="center"> <b>Project </b> </div></td>								
									<td colspan="7" class="blue-normal" align="right" bgColor="#617DC0"> 
										<table width="100%" border="0" cellspacing="0" cellpadding="0" class="blue-normal">
											<tr> 
												<td class="white">&nbsp;&nbsp; Timesheet of <b>Andy North - Chief Operating Officer <%=strTitle1%></b></td>										
											</tr>
										</table>
									</td>
									<td rowspan="2" class="white" bgcolor="#617DC0"> <div align="center"><b>Total</b></div></td>
								</tr>
								<tr bgcolor="#617DC0"> 
									<%for ii=strLastSat to strNextFri%>
									<td align="center" class="white">
										<b><%=iif(Weekday(ii,vbMonday)<=5,day(ii) & "-" & MonthName(month(ii),True),"<font color='#FF9999'>" & day(ii) & "-" & MonthName(month(ii),True) & "</font>")%></b></td>
									<%next%>
								</tr>
								
							</table>
						</div>
					</div>
					
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                    <button type="submit" class="btn btn-primary" id="btnSave">Approve</button>
                </div>
				</form>
            </div>
        </div>
    </div>
</div>       
      
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

<<script type="text/javascript" src="../../js/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/dataTables.bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap-datepicker.js" charset="UTF-8"></script>

<script type="text/javascript" src="../../js/library.js"></script>
<script src="../../bootstrap/js/bootstrap.min.js"></script>

<script language="javascript">
<!--

$.extend( true, $.fn.dataTable.defaults, {
     "paging": false,
		"searching": false,
		"processing": true,
		"ordering": false,
		"info":     false
} );

$(document).ready(function() {

    $('#tblList').DataTable( {       
        "serverSide": true,
        "ajax": "../../inc/JSON_listStaffs.asp",
        "columns": [
            { "data": "Fullname" },
            { "data": "JobTitle" },
            { "data": "Department" },
            { "data": "ReportTo" },
			{ "data": "approved" }
        ]
    });	
	
	
	    // Align modal when it is displayed
    $(".modal").on("shown.bs.modal", alignModal);
    
    // Align modal when user resize the window
    $(window).on("resize", function(){
        $(".modal:visible").each(alignModal);
    });   
	
	$('#tblList tbody').on('click', 'tr', function (e) {
	
		e.preventDefault();
		var table = $('#tblList').DataTable();
        var dataid = table.row(this).id();
		
		$("#txthidden").val(dataid);
		$("#frmList" ).submit(); 
		$('#myModal').modal({backdrop: "static"});
        
        //alert( 'You clicked on '+data+'\'s row' );
    });
	
	$('#filterBy').click(function(e){ 
		var cur;
		e.preventDefault();
		cur=$(this).text();
		if (cur=="By Manager" )
		{
			$("#filterLable").text("By Manager");
			$(this).text("By Fullname");
		}
		else
		{
			$("#filterLable").text("By Fullname");
			$(this).text("By Manager");
		}

		$("#btnFilter").dropdown("toggle");
		myFunction();
		return false; 
	});	
	
	var date = new Date();
	var today = new Date(date.getFullYear(), date.getMonth(), date.getDate());

	$('#datetimepicker').datepicker({
		todayHighlight: true		
	});
	
	$('#datetimepicker').datepicker('setDate', today);
	
	activeWeek();
	
	$('#datetimepicker').on('changeDate', function() {
		$('#selecteddate').val($('#datetimepicker').datepicker('getFormattedDate'));
			activeWeek();
		//alert($('#selecteddate').val());
	});
	
});

 function activeWeek() {
	$('.day.active').closest('tr').find('.day').addClass('active');
		
}

function alignModal(){

	var modalDialog = $(this).find(".modal-dialog");	
	// Applying the top margin on modal dialog to align it vertically center
	modalDialog.css("margin-top", Math.max(0, ($(window).height() - modalDialog.height()) / 2));
}

function myFunction() {
   var input, filter, table, tr, td, i, idx;
  input = document.getElementById("txtSearch");
  filter = input.value.toUpperCase();
  table = document.getElementById("tblList");
  tr = table.getElementsByTagName("tr");
	idx=0;
  if ($("#filterLable").text()=="By Manager")
	idx=3;
  // Loop through all table rows, and hide those who don't match the search query
  for (i = 0; i < tr.length; i++) {
    td = tr[i].getElementsByTagName("td")[idx];
    if (td) {
      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
        tr[i].style.display = "";
      } else {
        tr[i].style.display = "none";
      }
    } 
  }
}
//-->
</script>
</body>
</html>