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
	Dim strFirstDay, strParm, strURLSetHour, strColorOpt, strError, varTimesheet, varEvent,dateLimit
		
	Dim dateToday,dateFrom,dateTo
	dim arrCalendar,monthNamesEn,dayNamesEn,constWeedays,arrHoliday
	Dim dblTotalProHours,dblTotalLeaveHours,dblTotalOtherHours,fgUnlock
	
	dim arrProjectATS, arrEventATS,arrTotal
	
	dayNamesEn = array("", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun") 
	constWeedays=7
	
	ReDim arrTotal(constWeedays + 2, 3)
	
	arrTotal(1,1)="<b>&nbsp;Total Hours</b>"
	arrTotal(1,2)="<b>&nbsp;Normal Hours</b>"
	arrTotal(1,3)="<b>&nbsp;Overtime Hours</b>"
	
	for jj=1 to 3
		for ii=2 to constWeedays + 2
			arrTotal(ii,jj)=0
		next
	next
	
	dblTotalProHours=0
	dblTotalLeaveHours=0
	dblTotalOtherHours=0
'------------------------------------------------
'
'------------------------------------------------	
Function CheckHoliday(checkDate)
	dim blnReturn
	blnReturn=false
	if IsArray(arrHoliday) then
		For ii = 0 To Ubound(arrHoliday,2)
			If arrHoliday(0,ii) = checkDate Then
				blnReturn=true
				Exit For
			End If
		Next
	end if	
	CheckHoliday=blnReturn		
end function
'------------------------------------------------
'
'------------------------------------------------
function GetUnionTimesheetSQL(byval fromDate,byval toDate)
	dim strSql
	strSql=""
	if year(fromDate) =	year(toDate) then
		strSql=selectTable(Year(fromDate)) & " a "
	else
		strSql="("
		for ii=year(fromDate) to year(toDate)
			if strSql<> "(" then strSql=strSql & " UNION ALL "
			strSql=strSql & "SELECT StaffID,Tdate,Hours,OverTime,AssignmentID,EventID FROM " & selectTable(ii)
		next
		strSql=strSql & ") a "
	end if

GetUnionTimesheetSQL=strSql
end function

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
' Get Calendar Array
'--------------------------------------------------	
sub GetCalendarArray(byval today)
	
	dayOfMonth=GetDay(month(today),year(today))
	lastdate=cdate(month(today) & "/" & dayOfMonth & "/" & year(today))

	'Weekday of the first date current month
	weekDayOftheFirst=Weekday(today-day(today)+1,vbMonday)
	'WeekdayOftheLast=Weekday(today+day(today)+1,vbMonday)

	if (weekDayOftheFirst>0) then startDate=(today- day(today) + 1) - weekDayOftheFirst + 1

	numMembers=cint(dayOfMonth) + cint(weekDayOftheFirst - 1) + cint(7 - Weekday(lastdate,vbMonday))
	Redim arrCalendar(numMembers)
	
	for ii=1 to numMembers
		arrCalendar(ii)=startDate
		startDate = startDate + 1
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

	strAct=Request.QueryString("act")
	if strAct="go" then
		selMonth=Request.Form("lbmonth")
		selYear=Request.Form("lbyear")
		
		dateToday=cdate(selMonth & "/1/" & selYear)
		
		do while Weekday(dateToday,vbMonday)<>1
			dateToday=dateToday + 1
		loop
		dateFrom=dateToday
		dateToday=dateToday+7
	else
		dateToday=Request.Form("D")
		if dateToday="" then
			dateToday=Date()
		else
			dateToday=cdate(dateToday)
		end if	
		'the last Monday from today
		dateFrom=dateToday-(Weekday(dateToday,2) + 6)
	end if
	
	'dateToday=cdate("23-Jan-2006")
	
	'the last Sunday from today
	dateTo=dateFrom + 6
	
call GetCalendarArray(dateFrom)

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					
'--------------------------------------------------
' Check Approving Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgUnlock = False
	Else
		varGetRight = session("RightOn")
		fgUnlock = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "Unlock Timesheet" Then
				fgUnlock = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------	
	intUserID	= session("USERID")
	intStaffID  = Request.Form("txthidden")
	
call GetProTimeSheetArray(intStaffID,dateFrom,dateTo)
call GetEventTimeSheetArray(intStaffID,dateFrom,dateTo)


'--------------------------------------------------
' Get holiday
'--------------------------------------------------
strConnect = Application("g_strConnect")	' Connection string 	
	Set objDatabase = New clsDatabase 

	If objDatabase.dbConnect(strConnect) Then
		strSQL = "exec GetListHolidays null, null, '" & dateFrom & "', '" & dateTo & "', 1"
		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.noRecord = False Then
				arrHoliday = objDatabase.rsElement.GetRows
				objDatabase.closeRec
			End If
		Else
			strError = objDatabase.strMessage
		End If
	Else
		strError = objDatabase.strMessage		
	End If
	
'--------------------------------------------------
' Initialize appoval timesheet records
'--------------------------------------------------
	
	If objDatabase.dbConnect(strConnect) Then			
		strSQL = "SELECT * FROM ATC_TimesheetApproval WHERE DateFrom='" & dateFrom & "' AND DateTo='" & dateTo & "' AND StaffID=" & intStaffID

		Set rsTmsApproval = Server.CreateObject("ADODB.Recordset")
		Set rsTmsApproval.ActiveConnection = objDatabase.cnDatabase
		rsTmsApproval.CursorLocation = adUseClient
		
		rsTmsApproval.LockType=3
		
		rsTmsApproval.Open strSQL
		
		If Err.number =>0 then	
			strError = Err.Description
		else
			set rsTmsApproval.ActiveConnection=nothing
		end if
	Else
		Response.Write objDatabase.strMessage		
	End If

if strAct="app" then
	rsTmsApproval.AddNew()
	rsTmsApproval("DateFrom")=dateFrom
	rsTmsApproval("DateTo")=dateTo
	rsTmsApproval("StaffID")=intStaffID
	rsTmsApproval("ApprovalID")=intUserID

	set rsTmsApproval.ActiveConnection = objDatabase.cnDatabase
	rsTmsApproval.UpdateBatch
	rsTmsApproval.Requery()
	set rsTmsApproval.ActiveConnection=nothing
elseif strAct="unlock" then

	if rsTmsApproval.RecordCount=1 then
		rsTmsApproval.Delete(adAffectCurrent)
		set rsTmsApproval.ActiveConnection = objDatabase.cnDatabase
		rsTmsApproval.UpdateBatch
		rsTmsApproval.Requery()
		set rsTmsApproval.ActiveConnection=nothing
	end if
	
end if

'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view timesheet&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Timesheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"

	if not rsTmsApproval.EOF then 
		objEmployee.SetFullName(rsTmsApproval("ApprovalID"))
		varFullName = split(objEmployee.GetFullName,";")
		strAppFullName="<b>" & varFullName(0) & "</b>" 
	end if
	
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

<script language="javascript" src="../../library/library.js"></script>
<script language="javascript" src="../../library/menu.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objNewWindow;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function viewtms()
{
	var URL;
	var dToday,yToday;
	var selectMonth,selectYear;
	mToday=<%=month(Date())%>;
	yToday=<%=year(Date())%>;
	
	selectMonth=window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value;
	selectYear= window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value;
	
	
	if ((selectMonth<=mToday && selectYear==yToday)||(selectYear<yToday))
	{
	//alert (window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + "/" + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value);
		URL = "tms_approval.asp?act=go";

		window.document.frmtms.action = URL;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}
	else
		alert ("Please ensure the month fields is earlier or equal to the current date.");
	
}

function approvetms()
{
	var url;
	
	var agree=confirm("WARNING: Do NOT just press enter! Before continuing make sure that you really checked this timesheet.");
	if (agree)
	{	
		URL = "tms_approval.asp?act=app"
		window.document.frmtms.action = URL;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();

	}	
}

function unlocktms()
{
	var url;
	var agree=confirm("Are you sure that you want to unlock this timesheet?");
	if (agree)
	{
		URL = "tms_approval.asp?act=unlock"
		window.document.frmtms.action = URL;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}
}
function logout()
{
	var url;
	url = "../../logout.asp";
		window.document.frmtms.action = url;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
}

function back_menu()
{
	window.document.frmtms.action = "tms_listfor_approval.asp?b=1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function selectDate(day,month,year)
{
	window.document.frmtms.D.value=month + "/" + day + "/" + year;
	window.document.frmtms.action = "tms_approval.asp";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function viewdetail()
{
	window.document.frmtms.action = "tms_viewdetails.asp";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
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
		objNewWindow = window.open('tms_select_staff.asp?view=a', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
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
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
<%
	If strError <> "" Then
%>
  <tr height="20">
    <td>
      <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td class="red" align="center">&nbsp;<b><%=strError%></b></td>
        </tr>
      </table>    
    </td>
  </tr>  
<%	End if%>  
	<tr> 
		<td width="156" align="center" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" class="monthtableouter">
				<tr>
					<td>
						<table border="0" cellpadding="1" cellspacing="1" class="monthtableinner">
							<tr bgcolor="#FFFFFF">
								<td class="red" colspan="7" align=center><b><%=SayMonth(Month(dateFrom)) & " " & year(dateFrom)%></b></td>
							</tr>
							<tr>
								<%For ii = 1 To UBound(dayNamesEn)%>
								<td class="days" align=center width="14%"><%=dayNamesEn(ii)%></td>
								<%next%>
								
							</tr>
							<%
							numRows=Ubound(arrCalendar) \ constWeedays
							if (Ubound(arrCalendar) mod constWeedays >0) then numRows=numRows +1
							For ii=1 to numRows%>
							<tr>
								<%for jj=1 to constWeedays
									'Class for background
									strClass="normalday"
									if jj>=constWeedays-1 then strClass="weekend"																		
									if ((ii-1) * constWeedays + jj<=Ubound(arrCalendar)) then
										dayIdx=arrCalendar((ii-1) * constWeedays + jj)
										if dayIdx<=date() or month(dayIdx)<=month(date()) then
										strID =iif(cdate(dayIdx)=cdate(dateToday),"id = 'today'","")											
										strUrl="<a href='javascript:selectDate(" & day(dayIdx+7) & "," & month(dayIdx+7) & "," & year(dayIdx+7) & ");' class='day'>" & IIF(month(dayIdx)=month(dateFrom),day(dayIdx),"<font color='silver'>" & day(dayIdx) & "</font>") & "</a>"
										%>
										<td align="center" <%=IIf(dayIdx>=dateFrom and dayIdx <=dateTo,"bgcolor='#d6dbef'","class='" & strClass &"'")%><%'=strID%>>
											<%=strUrl%></td>										
									<%	else%>
										<td  class="<%=strClass%>">&nbsp;</td>
									<%	end if
									end if
								next%>
							</tr>
							<%next%>		
							<tr>
								<td align="center" colspan="5" nowrap>
									<select name="lbmonth" class="month">
										<%for ii=1 to 12%>
										<option value="<%=ii%>" <%=IIf(ii=month(dateFrom),"selected","")%>><%=SayMonth(ii)%>&nbsp;</option>
										<%next%>
									</select>
									<select name="lbyear" class="month">
										<%for ii=2000 to year(date())%>
										<option value="<%=ii%>" <%=IIf(ii=year(dateFrom),"selected","")%>><%=ii%>&nbsp;</option>
										<%next%>
									</select>
								</td>
								<td colspan="2" nowrap>
									<table border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa" width="40">
										<tr> 
											<td bgcolor="#f6f6f6" onMouseOver="this.style.backgroundColor='#dbdbdb';" onMouseOut="this.style.backgroundColor='#f6f6f6';" height="20"> 
												<div align="center" class="day"><a href="javascript:viewtms();" onMouseOver="self.status='View timesheet';return true" onMouseOut="self.status='';return true">Go</a></div>
											</td>
										 </tr>
									 </table>
								</td>
							</tr>
							<tr bgcolor="#FFFFFF">
								<td align="left" colspan="7" class="blue-normal">
									*Choose either a date or month for which you want to see details
								</td>
							</tr>								
						</table>
					</td>
				</tr>
			</table>
			<table width="100%"  bgcolor="#FFFFFF">
				<tr> 
					<td class='blue' colspan="4"><br>&nbsp;Summary of hours:</td>
				</tr>
				<tr  class="blue-normal"> 
				  <td width="8%" align="right"><img height='5' src='../../images/dot.gif' width='5'></td>
					<td width="53%">Project Hours:</td>
					<td width="37%" align="right"> <%=FormatNumber(dblTotalProHours,1)%></td>
					<td width="2%">&nbsp;</td>
				</tr>	
				<tr class="blue-normal"> 
				    <td align="right"><img height='5' src='../../images/dot.gif' width='5'></td>
					<td>Others:</td>
					<td align="right"><%=FormatNumber(dblTotalOtherHours,1)%></td>
					<td>&nbsp;</td>
				</tr>									
				<tr class="blue-normal"> 
					<td  align="right"><img height='5' src='../../images/dot.gif' width='5'></td>
					<td >Leave Hours:</td>
					<td align="right"> <%=FormatNumber(dblTotalLeaveHours,1)%></td>
					<td>&nbsp;</td>
				</tr>
				
				<tr class="blue"> 
				    <td >&nbsp;</td>
					<td align="right">Total:</td>
					<td align="right"><%=FormatNumber(dblTotalOtherHours + dblTotalProHours + dblTotalLeaveHours,1)%></td>
					<td>&nbsp;</td>
				</tr>	
<%if rsTmsApproval.EOF then%>
				<tr> 
					<td class='blue' colspan="4">
						<table border="0" cellspacing="5" cellpadding="0" height="20" name="aa" width="100%">
						  <tr> 
						    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
						      <div align="center" class="blue"><a href="javascript:approvetms();" class="b" onMouseOver="self.status='Approved Timesheet';return true" onMouseOut="self.status='';return true">Approve</a></div>
						    </td>
						  </tr>
						</table>
					</td>
				</tr>
				<tr> 
					<td class='blue-normal' colspan="4"><i>(From <%=day(dateFrom) & "/" & month(dateFrom) & "/" & year(dateFrom)%> to <%=day(dateTo) & "/" & month(dateTo) & "/" & year(dateTo)%>)</i></td>
				</tr>
<%else%>				
				<tr> 
					<td class='blue' colspan="4">&nbsp;Approved By:</td>
				</tr>
				<tr>
					<td  class="red" colspan="4" align="center"><%=strAppFullName%></td>
				</tr>
				<tr> 
					<td class='blue-normal' colspan="4" align="center"><%=rsTmsApproval("DateApproval")%><br></td>
				</tr>
				<%if fgUnlock then%>
				<tr> 
					<td class='blue' colspan="4">
						<table border="0" cellspacing="5" cellpadding="0" height="20" name="aa" width="100%">
						  <tr> 
						    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
						      <div align="center" class="blue"><a href="javascript:unlocktms();" class="b" onMouseOver="self.status='Unlock Timesheet';return true" onMouseOut="self.status='';return true">Unlock Timesheet</a></div>
						    </td>
						  </tr>
						</table>
					</td>
				</tr>
				<%End if%>

<%end if%>
			</table>
		</td>
		
		
          <td bgcolor="#8FA4D3" valign="top" width="624"> 
            <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
                <tr> 
                  <td width="178" colspan="2" rowspan="2" class="white" bgcolor="#617DC0"> 
                    <div align="center"> <b>Project </b> </div></td>								
                  <td width="350" colspan="7" class="blue-normal" align="right" bgColor="#617DC0"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="blue-normal">
                      <tr> 
                        <td width="57%" class="white">&nbsp;&nbsp;<%=strTitle1%></td>
                        
                      </tr>
                    </table></td>
                  <td width="96" rowspan="2" class="white" bgcolor="#617DC0"> <div align="center"><b>Total</b></div></td>
                </tr>
                <tr bgcolor="#617DC0"> 
					<%for ii=dateFrom to dateTo%>
						<td width="50" align="center" class="white">
							<b><%=iif(Weekday(ii,vbMonday)<=5,day(ii) & "-" & MonthName(month(ii),True),"<font color='#FF9999'>" & day(ii) & "-" & MonthName(month(ii),True) & "</font>")%></b></td>
					<%next%>
                </tr>
                <!--**************************** For Project And SubTask *********************************************-->
                <%
					numOfRows=UBound(arrProjectATS,2)
					numOfCols=UBound(arrProjectATS,1)
					for idxRow=1 to numOfRows%>
						<tr>
							<td width="8" bgcolor="#FFC6C6" class="white"><%if arrProjectATS(1,idxRow)<>"&nbsp;" then%>
								<img src="../../images/cross.gif" width="8" height="14"><%else%>&nbsp;<%end if%></td>
				<%		for idxCol=1 to numOfCols
							strColor="#FFFFFF"
							'For ProjectName and Total
							if idxCol=1 or idxCol=numOfCols then
								strColor="#FFF2F2"
							'for Sat
							elseif (idxCol=numOfCols-1) then
								strColor="#C2CCE7"
							'for Sun
							elseif (idxCol=numOfCols-2) then
								strColor="#E7EBF5"
							end if
							'For public holiday
							if idxCol>1 and idxCol<numOfCols then
								if CheckHoliday(dateFrom + (idxCol-2)) then strColor="#FFC6C6"
							end if
							%>
							<td <%=IIF(idxCol=1,"width='170'","")%>bgcolor="<%=strColor%>" align="<%=IIF(idxCol=1,"left","center")%>" class="blue-normal"><%=arrProjectATS(idxCol,idxRow)%></td>
				<%		next%>		
						</tr>
				<%	next%>                
                <!--**************************** End Of Project And SubTask *******************************-->                
                <!--**************************** For Events and Others ************************************-->
				<%
				numOfRows=UBound(arrEventATS,2)
				for idxRow=1 to numOfRows%>
				<tr>
				<%
					for idxCol=1 to numOfCols
						strColor="#E7EBF5"
						'For EventName and Total
						if idxCol=1 or idxCol=numOfCols then
							strColor="#FFE1E1"
						'for Sat
						elseif (idxCol=numOfCols-1) then
							strColor="#C2CCE7"
						'for Sun
						elseif (idxCol=numOfCols-2) then
							strColor="#D2DAEC"
						end if
						'For public holiday
						if idxCol>1 and idxCol<numOfCols then
							if CheckHoliday(dateFrom + (idxCol-2)) then strColor="#FFC6C6"
						end if%>
                <td <%=IIF(idxCol=1,"colspan='2'","")%> class="blue-normal" bgcolor="<%=strColor%>" align="<%=IIF(idxCol=1,"left","center")%>"><%=arrEventATS(idxCol,idxRow)%></td>
					<%next%>
                </tr>
                <%Next%> 
                
                 <!--**************************** For Total rows ************************************-->
                 <%
				numOfRows=UBound(arrTotal,2)
				for idxRow=1 to numOfRows%>
				<tr>
				<%
					for idxCol=1 to numOfCols
						strColor="#E7EBF5"
						'For Total
						if idxCol=1 or idxCol=numOfCols then
							strColor="#FFE1E1"
						'for Sat
						elseif (idxCol=numOfCols-1) then
							strColor="#C2CCE7"
						'for Sun
						elseif (idxCol=numOfCols-2) then
							strColor="#D2DAEC"
						end if
						'For public holiday
						if idxCol>1 and idxCol<numOfCols then
							if CheckHoliday(dateFrom + (idxCol-2)) then strColor="#FFC6C6"
							arrTotal(numOfCols,idxRow)=arrTotal(numOfCols,idxRow) + arrTotal(idxCol,idxRow)
						end if
						%>						
                <td <%=IIF(idxCol=1,"colspan='2'","")%> class="blue-normal" bgcolor="<%=strColor%>" align="<%=IIF(idxCol=1,"left","center")%>">
						<%if idxCol=1 then Response.Write arrTotal(idxCol,idxRow) else Response.Write IIF(arrTotal(idxCol,idxRow)<>0,"<b>" & FormatNumber(arrTotal(idxCol,idxRow),1) & "</b>","&nbsp;") end if%></td>
					<%next%>
                </tr>
                <%Next%> 
                <!--**************************** End Of Events and Others *********************************************-->
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
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="D" value="<%=dateToday%>">
<input type="hidden" name="txthidden" value="<%=intStaffID%>">

<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
<input type="hidden" name="assign" value="<%=Request.Form("assign")%>">
<input type="hidden" name="row" value="">
</form>
</body>
</html>
