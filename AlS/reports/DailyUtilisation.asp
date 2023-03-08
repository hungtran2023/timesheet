<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
'*********************************************************
'Generate report
'*********************************************************
Function ATSsql(byval strDate)
	dim strATS	
	strATS=selectTable(year(strDate))
	'if strATS="ATC_Timesheet2004" then strATS="ATC_Timesheet"
	ATSsql=strATS
End function
'*********************************************************
'Initial Data
'*********************************************************
Function InitialData(byval strDate)
	dim strSql
	if isEmpty(session("HoursProject")) or isEmpty(session("ListProject")) or isEmpty(session("NonBillableG")) or isEmpty(session("StaffOfManager"))then

		strConnect = Application("g_strConnect")	
		set DbConn = Server.CreateObject("ADODB.Connection")	
		DbConn.Open strConnect		
		strSql="SELECT ISNULL(c.FirstName + ' ' + c.LAstName,'yUnassign') as FullName, projectName,projectKey2,StaffCount,a.ProjectID,a.ManagerID " & _
				"FROM ATC_Projects a INNER JOIN " & _ 
					"(SELECT a1.DateTransfer,Projectkey2,a1.projectID FROM ATC_ProjectStage a1 " & _
					"INNER JOIN (SELECT MAX(DateTransfer) AS DateTransfer,ProjectID FROM ATC_ProjectStage GROUP BY projectID) b1 " & _
					"ON (a1.ProjectID=b1.ProjectID AND a1.DateTransfer=b1.DateTransfer)) b ON a.ProjectID=b.ProjectID " & _
				"LEFT JOIN ATC_PersonalInfo c ON a.ManagerID=c.PersonID " & _
				"INNER JOIN " & _
					"(SELECT COUNT(DISTINCT StaffID) as StaffCount,e.ProjectID " & _
						"FROM ATC_Assignments d " & _
						"INNER JOIN ATC_Tasks e ON d.SubTaskID=e.SubTaskID " & _
						"INNER JOIN ATC_PersonalInfo f ON f.PersonID=d.StaffID " & _
						"WHERE f.fgDelete=0 AND d.fgDelete=0 " & _
						"GROUP BY e.ProjectID) AS StaffCount ON StaffCount.ProjectID=a.ProjectID " & _
				"WHERE a.fgActivate=1 and a.fgDelete=0 AND projectKey2<>7 ORDER BY FullName"
'Response.Write strSql			
		Set rsProList = Server.CreateObject ("ADODB.Recordset")
		rsProList.CursorLocation=adUseClient
		rsProList.Open strSql,dbConn,adOpenStatic,adLockBatchOptimistic
		set rsProList.ActiveConnection=nothing

		set session("ListProject")=rsProList.Clone

		strSQl="SELECT ISNULL(SUM(a.Hours+OVerTime)/numberManager,0) as Hours,a.StaffID FROM " & ATSsql(strDate) & " a " & _
				"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
				"INNER JOIN ATC_Tasks c ON b.SubtaskID=c.SubtaskID " & _
				"INNER JOIN (SELECT a1.DateTransfer,Projectkey2,a1.projectID FROM ATC_ProjectStage a1 " & _
							"INNER JOIN (SELECT MAX(DateTransfer) AS DateTransfer,ProjectID FROM ATC_ProjectStage GROUP BY projectID) b1 " & _
							"ON (a1.ProjectID=b1.ProjectID AND a1.DateTransfer=b1.DateTransfer)) d ON d.ProjectID=c.ProjectID " & _
				"INNER JOIN (SELECT COUNT(DISTINCT(d.ManagerID)) as numberManager,a.StaffID FROM ATC_Assignments a " & _
									"INNER JOIN ATC_Tasks b ON b.SubtaskID=a.SubtaskID " & _
									"INNER JOIN  ATC_Projects d ON d.ProjectID=b.ProjectID " & _
							"WHERE a.fgDelete=0 AND d.fgdelete=0 AND d.fgActivate=1 " & _
							"GROUP BY a.StaffID) e ON e.StaffID=a.StaffID	" & _
				"WHERE Tdate='" & strDate & "' AND (a.EventID<>1 or projectkey2=7) " & _
				"GROUP BY  a.StaffID,numberManager " & _
				"ORDER BY a.StaffID"
'Response.Write strSql		
		Call GetRecordset(strSql,rsNonBillG)
		set session("NonBillableG")=rsNonBillG.Clone
		
		strSQl="SELECT DISTINCT(c.StaffiD),ManagerID " & _
				"FROM ATC_Projects a INNER JOIN ATC_Tasks b ON a.ProjectID=b.ProjectID " & _
					"INNER JOIN ATC_Assignments c ON b.SubTaskID=c.SubTaskID " & _
				"WHERE c.fgDelete=0 AND a.fgActivate=1 AND StaffID IN " & _
					"(SELECT StaffID FROM ATC_Employees WHERE LeaveDate IS NULL) " & _
				"ORDER BY ManagerID"
		Call GetRecordset(strSql,rsStaffOfManager)
		set session("StaffOfManager")=rsStaffOfManager.clone
		
		strSql="SELECT Hours+OverTime as Hours,f.ManagerID,projectkey2,c.ProjectID " & _
				"FROM " & ATSsql(strDate) & " a " & _
					"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
					"INNER JOIN ATC_Tasks c ON b.SubtaskID=c.SubtaskID " & _
					"INNER JOIN (SELECT a1.DateTransfer,Projectkey2,a1.projectID " & _
								"FROM ATC_ProjectStage a1 INNER JOIN " & _
									"(SELECT MAX(DateTransfer) AS DateTransfer,ProjectID FROM ATC_ProjectStage GROUP BY projectID) b1 " & _
									"ON (a1.ProjectID=b1.ProjectID AND a1.DateTransfer=b1.DateTransfer)) d " & _
						"ON d.ProjectID=c.ProjectID INNER JOIN ATC_Projects f ON f.ProjectID =c.ProjectID " & _
					"WHERE Tdate='" & strDate & "' AND ((projectkey2=5) OR(projectkey2=1 AND a.AssignmentID<>1)) "
'Response.Write strSql
		Set rsHourProject = Server.CreateObject ("ADODB.Recordset")
		rsHourProject.CursorLocation=adUseClient
		rsHourProject.Open strSql,dbConn,adOpenStatic,adLockBatchOptimistic
		set rsHourProject.ActiveConnection=nothing
		DbConn.Close()
		set DbConn=nothing		
		set session("HoursProject")=rsHourProject.Clone
	else
		set rsProList=session("ListProject").Clone
		set rsNonBillG=session("NonBillableG").Clone
		set rsStaffOfManager=session("StaffOfManager").clone
		set rsHourProject=session("HoursProject").Clone
	end if
End function
'*********************************************************
'Update billable status temporary
'*********************************************************
Function UpdateBillableProject(byval strUpdateProject)
	dim arrProjectID
	arrProjectID=split(strUpdateProject,",")
	for i=0 to UBound(arrProjectID)-1
		rsProList.Filter= "ProjectID='" & arrProjectID(i) & "'"
		if not rsProList.EOF then
			if rsProList("Projectkey2")=5 then
				rsProList("Projectkey2")=1
			else
				'rsHourProject("Projectkey2")=5
				rsProList("Projectkey2")=5
			end if
			rsProList.Update
		end if
		rsProList.Filter=""
		rsHourProject.Filter= "ProjectID='" & arrProjectID(i) & "'"
		if not rsHourProject.EOF then
			do while not rsHourProject.EOF
				if rsHourProject("Projectkey2")=5 then
					rsHourProject("Projectkey2")=1
				else
					rsHourProject("Projectkey2")=5
				end if
				rsHourProject.MoveNext
			loop
			rsHourProject.UpdateBatch
		end if
		rsHourProject.Filter= ""
	next
End function
'*********************************************************
'Calculate Utilisation
'*********************************************************
Function CalculateUtilisation(byval intManagerID,byval strF,byref dblBillable,byref dblNonBillablePro,byref dblNonBillableGeneral,byref intStaffs)
	dim strSQl
	dblBillable=GetBillableHours(intManagerID,strF)
	dblNonBillablePro=GetNoneBillableHoursProject(intManagerID,strF)
	dblNonBillableGeneral=GetNoneBillableHoursGeneral(intManagerID,strF,intStaffs)	
End function
'*********************************************************
'Get Billable Hours
'*********************************************************
Function GetBillableHours(byval intManagerID,byval strDate)
	dim strSQl,rsTemp
	dim dblValue
	dblValue=0

	rsHourProject.Filter="ManagerID=" & intManagerID & " AND projectkey2=1" 
	do while not rsHourProject.EOF
		dblValue=dblValue + cdbl(rsHourProject("Hours"))
		rsHourProject.MoveNext
	loop	
	rsHourProject.Filter=""
	GetBillableHours=dblValue
	
End function

'*********************************************************
'Get NoneBillable Hours for Project (5)
'*********************************************************
Function GetNoneBillableHoursProject(byval intManagerID,byval strDate)
	dim dblValue
	dblValue=0
	
	rsHourProject.Filter="ManagerID=" & intManagerID & " AND projectkey2=5" 
	do while not rsHourProject.EOF
		dblValue=dblValue + cdbl(rsHourProject("Hours"))
		rsHourProject.MoveNext
	loop	
	rsHourProject.Filter=""	
	
	GetNoneBillableHoursProject=dblValue
	
End function

'*********************************************************
'Get NoneBillable Hours for Genneral (7 and Event)
'*********************************************************
Function GetNoneBillableHoursGeneral(byval intManagerID,byval strDate,byref intStaffs)
	dim dblValue
	dblValue=0
	if rsNonBillG.recordcount>0 then
		rsStaffOfManager.Filter="ManagerID=" & intManagerID
		intStaffs=rsStaffOfManager.recordCount
		Do while not rsStaffOfManager.EOF
			rsNonBillG.MoveFirst
			rsNonBillG.Find "StaffID=" & rsStaffOfManager("StaffID")
			if not rsNonBillG.EOF then dblValue=dblValue+ cdbl(rsNonBillG("Hours"))
			rsStaffOfManager.MoveNext
		Loop
		rsStaffOfManager.Filter=""
	end if
	GetNoneBillableHoursGeneral=FormatNumber(dblValue,2)
	
End function
'*********************************************************
'Get Book Hour Project
'*********************************************************
Function GetBookHourProject(projectID)
	dim dblValue
	dblValue=0
	
	rsHourProject.Filter="ProjectID='" & projectID & "'"
	do while not rsHourProject.EOF
		dblValue=dblValue + cdbl(rsHourProject("Hours"))
		rsHourProject.MoveNext
	loop	
	rsHourProject.Filter=""	
	
	GetBookHourProject=dblValue	
End function
'*********************************************************
'Generate report
'*********************************************************
Function GenerateReport(strdate)
	dim strSql, strATS,strReturn,arrlongmon
	dim strManagerFirst,dblBillable,dblNonBillablePro,dblNonBillableGen,intStaffs
	dim dblTotalBillable,dblTotalNonBillablePro,dblTotalNonBillableGen
	dim strProID,strSID,ii

	dblTotalBillable=0
	dblTotalNonBillableGen=0
	dblTotalNonBillablePro=0
	
	if rsProList.RecordCount>0 then
		rsProList.MoveFirst
		strManagerFirst=""	
		Do while not rsProList.EOF
			strBgcolor="#FFFFFF"
			'if rsPro("CSOCompleted") then strBgcolor="#F1DADB"
			dblBillable=""
			dblNonBillablePro=""
			dblNonBillableGen=""
			dblUtilisation=""
								
			if strManagerFirst <> rsProList("FullName") then
				if strManagerFirst<>"" then
					strReturn=strReturn & "<tr valign='top' bgcolor='#FFF2F2' class='blue-normal'> " & vbCrLf 
					strReturn=strReturn & "<td colspan='4' align='right'>Number of Staffs</td>"
					strReturn=strReturn & "<td colspan='5' align='left'><b>" & intStaffs & "</b></td></tr>"
				end if
				strReturn=strReturn & "<tr valign='top' bgcolor='" & strBgcolor & "' class='blue-normal'> " & vbCrLf 		
				strManagerFirst = rsProList("FullName")
				if strManagerFirst ="yUnassign" then 
					strReturn=strReturn & "<td><b>Unassigned</b></td>" & vbCrLf
				else
					strReturn=strReturn & "<td><b>" & strManagerFirst & "</b></td>" & vbCrLf
				end if
				Call CalculateUtilisation(rsProList("ManagerID"),strDate,dblBillable,dblNonBillablePro,dblNonBillableGen,intStaffs)
				
				dblTotalBillable=dblTotalBillable + dblBillable
				dblTotalNonBillableGen=dblTotalNonBillableGen + dblNonBillableGen
				dblTotalNonBillablePro=dblTotalNonBillablePro + dblNonBillablePro
				
				if cdbl(dblBillable) + cdbl(dblNonBillablePro) + cdbl(dblNonBillableGen)>0 then
					dblUtilisation= FormatNumber(cdbl(dblBillable) * 100 /(cdbl(dblBillable) + cdbl(dblNonBillablePro) + cdbl(dblNonBillableGen)),2)
				end if
			else
				strReturn=strReturn & "<tr valign='top' bgcolor='" & strBgcolor & "' class='blue-normal'> " & vbCrLf 		
				strReturn=strReturn & "<td></td>" & vbCrLf
			end if
			'strTemp=split(rsProList("ProjectID"),"_")				
			strReturn=strReturn & "<td>" & ucase(Left(rsProList("ProjectID"),10)) & " - " & rsProList("projectName") & "</td>" & vbCrLf

			strReturn=strReturn & "<td align='center'><input type='checkbox' name='chkActivate' value='" & rsProList("projectKey2") & "'"
			if cint(rsProList("projectKey2"))=1 then strReturn=strReturn & "checked"
			strReturn=strReturn & "><input type='hidden' name='txtProjectID' value='" & rsProList("ProjectID") & "'></td>" & vbCrLf
				
			strReturn=strReturn & "<td align='right'>" & rsProList("StaffCount") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>" & vbCrLf
			strReturn=strReturn & "<td align='right'><b>" & GetBookHourProject(rsProList("ProjectID")) & "</b></td>" & vbCrLf
			strReturn=strReturn & "<td align='right'><b>" & dblBillable & "</b></td>" & vbCrLf
			strReturn=strReturn & "<td align='right'><b>" & dblNonBillablePro & "</b></td>" & vbCrLf
			strReturn=strReturn & "<td align='right'><b>" & dblNonBillableGen & "</b></td>" & vbCrLf
				
			strReturn=strReturn & "<td align='right'><b>" & dblUtilisation & "</b></td>" & vbCrLf
			rsProList.MoveNext
			strReturn=strReturn & "</tr>" & vbCrLf
		Loop
		strReturn=strReturn & "<tr valign='top' bgcolor='#FFF2F2' class='blue-normal'> " & vbCrLf 
		strReturn=strReturn & "<td colspan='4' align='right'>Number of Staffs</td>"
		strReturn=strReturn & "<td colspan='5' align='left'><b>" & intStaffs & "</b></td></tr>"
		
		'For total
		
		if cdbl(dblTotalBillable) + cdbl(dblTotalNonBillablePro) + cdbl(dblTotalNonBillableGen)>0 then
			dblTotalUtilisation= FormatNumber(cdbl(dblTotalBillable) * 100 /(cdbl(dblTotalBillable) + cdbl(dblTotalNonBillablePro) + cdbl(dblTotalNonBillableGen)),2)
		end if
		
		strReturn=strReturn & "<tr valign='top' bgcolor='#FFF2F2' class='blue-normal'> " & vbCrLf 		
		strReturn=strReturn & "<td colspan='5' align='right'><b>Total</b></td>" & vbCrLf
		strReturn=strReturn & "<td align='right'><b>" & dblTotalBillable & "</b></td>" & vbCrLf
		strReturn=strReturn & "<td align='right'><b>" & dblTotalNonBillablePro & "</b></td>" & vbCrLf
		strReturn=strReturn & "<td align='right'><b>" & dblTotalNonBillableGen & "</b></td>" & vbCrLf
				
		strReturn=strReturn & "<td align='right'><b>" & dblTotalUtilisation & "</b></td></tr>" & vbCrLf
		
	end if
	
	GenerateReport=strReturn
End Function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If
'-------------------------------
' Calculate pagesize
'-------------------------------
	if not isEmpty(session("Preferences")) then
		arrPre = session("Preferences")
		if arrPre(1, 0)>0 then PageSize = arrPre(1, 0) else PageSize = PageSizeDefault
		set arrPre = nothing
	else
		PageSize = PageSizeDefault
	end if
	
'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 
	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../welcome.asp")
	end if
'----------------------------------
' Get report
'----------------------------------	
Dim intDate,intMonth,intYear,arrlstDay(2),strprintdate,strfromto,strfrom,strTo
Dim gMessage,intnumMonths
Dim rsNonBillG,rsStaffOfManager,rsHourProject,rsProList

intDate=Request.Form("lstday")
if Request.Form("lstday")="" then intDate=Day(date - 1)
intMonth=Request.Form("lstmonth")
if Request.Form("lstmonth")="" then intMonth=month(date - 1)
intYear=Request.Form("lstyear")
if Request.Form("lstyear")="" then intYear=year(date - 1)

strDate=cdate(intMonth & "/" & intDate & "/" & intYear)

strprintdate=FormatDateTime(cdate(month(date) & "/" & Day(date) & "/" & year(date)),1)

if Request.QueryString("act")="" then
	session("ListProject")=empty
	session("NonBillableG")=empty
	session("StaffOfManager")=empty
	session("HoursProject")=empty
end if

call InitialData(strDate)

if Request.QueryString("act")="recal" then
	strUpdateProject=Request.Form("txtBillable")
	call UpdateBillableProject(strUpdateProject)
end if
'Get List for date
arrlstDay(0) = selectmonth("lstmonth",cint(intMonth), -1)
arrlstDay(1) = selectday("lstday", cint(intDate), -1)
arrlstDay(2) = selectyear("lstyear", cint(intYear),year(now())-2, year(now()), -1)


strLast = GenerateReport(strDate)
if not isEmpty(session("rpt_forecast")) then session("rpt_forecast")=empty
session("rpt_forecast")=strLast
'----------------------------------
' Get Company Information
'----------------------------------
if isEmpty(session("arrInfoCompany")) then
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
		strQuery = "SELECT a.CompanyName, isnull(Address,'') Address, isnull(City,'') City, isnull(b.CountryName,'') Country, " &_
					"isnull(Phone,'') Phone, isnull(Fax,'') Fax, isnull(c.Logo,'') Logo FROM ATC_Companies a " &_
					"LEFT JOIN ATC_Countries b On a.CountryID = b.CountryID " &_
					"LEFT JOIN ATC_CompanyProfile c ON a.CompanyID = c.CompanyID " &_
					"WHERE a.CompanyID = " & session("Inhouse")
		If objDb.runQuery(strQuery) Then
			If not objDb.noRecord then
				arrInfoCompany = objDb.rsElement.getRows
				session("arrInfoCompany") = arrInfoCompany
				objDb.closerec
			end if
		Else
			gMessage = objDb.strMessage
		end if
		objDb.dbDisconnect
	Else
		gMessage = objDb.strMessage
	End if
	set objDb = nothing
end if
'----------------------------------
' Get Full Name and Job Title
'----------------------------------

	Set objEmployee = New clsEmployee
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'><a href='../welcome.asp?choose_menu=B' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;" &_
				"<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing	
	
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_report.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
if not isEmpty(session("arrInfoCompany")) then
	arrTmp = session("arrInfoCompany")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", arrTmp(0, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", arrTmp(1, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", arrTmp(2, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", arrTmp(3, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", arrTmp(4, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", arrTmp(5, 0))
	if arrTmp(6, 0)<>"" then
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "<img src='../images/" & arrTmp(6, 0) & "' border='0'>" )
	else
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;" )
	end if
	set arrTmp = nothing
else
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;")
end if
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script>
var objWindowSumPro;

function _print() { //v2.0
var str1 = "<%=strfromto%>";
	str1 = escape(str1);
var str2 = "<%=strprintdate%>";
	str2 = escape(str2);
var fgprint = 10;
if (fgprint!=0) {
	window.status = "";
	strFeatures = "top="+(screen.height/2-275)+",left="+(screen.width/2-390)+",width=800,height=550,toolbar=no," 
	            + "menubar=yes,location=no,resizable=yes,directories=no,scrollbars=yes,status=yes";
	if ((objWindowSumPro) && (!objWindowSumPro.closed)) {
		objWindowSumPro.focus();
	
	} else {
		objWindowSumPro = window.open("p_DailyUtilisation.asp?fromto=" + str1 + "&printdate=" + str2 + "&num=" + 7, "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";
  }
else
	alert("No data for your request.")
}

function window_onunload() {
	if((objWindowSumPro) && (!objWindowSumPro.closed))
		objWindowSumPro.close();
}

function checkdata()
{
	var dateTranfer=document.frmreport.lstday.value + "/" + document.frmreport.lstmonth.value + "/" + document.frmreport.lstyear.value;

	if(isdate(dateTranfer)==false) {
		alert("The date tranfer (" + dateTranfer + ") is invalid.");
		document.frmreport.lstday.focus();
		return false;
	}
	return true;
}

function _submit() {
	if(checkdata()==true) {
		document.frmreport.action = "DailyUtilisation.asp";
		document.frmreport.target = "_self" ;
		document.frmreport.submit();
	}
}

function checkass () {
  strID="";
  len = document.frmreport.chkActivate.length;
  for(var ii=0; ii<len; ii++) {
	if ((document.frmreport.chkActivate[ii].checked && (document.frmreport.chkActivate[ii].value=='5'))||
		(!document.frmreport.chkActivate[ii].checked && (document.frmreport.chkActivate[ii].value=='1'))) {
		strID = strID + document.frmreport.txtProjectID[ii].value + ",";
	}
  }
  return(strID)
}

function _ReCalculate() {	
	strID=checkass();
	if (strID!="")
	{
		document.frmreport.txtBillable.value = strID
		document.frmreport.action = "DailyUtilisation.asp?act=recal";
		document.frmreport.target = "_self" ;
		document.frmreport.submit();
	}
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE=javascript onunload="return window_onunload()">
<form name="frmreport" method="post">
	<%
	'--------------------------------------------------
	' Write the header of HTML page
	'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
	%>


	
  <table width="90%" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%"  align="center" >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top" >
     
	  <table width="780" border="0" cellspacing="0" cellpadding="0">

          <tr style="background-color:Gray"> 
            <td ><img src="../IMAGES/dot1px.gif" width="1" height="10"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="60%"><table width="100%" border="0" cellpadding="0" cellspacing="0" >
                <tr> 
                  <td width="18%" class="blue-normal">Report Date</td>
                  <td width="32%">
<%
Response.Write arrlstDay(1)
Response.Write arrlstDay(0)
Response.Write arrlstDay(2)
%>                   
                  </td>
                  <td width="15%"  class="blue-normal" align="left">
					<table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
						<tr> 
							<td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
								<a href="javascript:_submit();" class="b">Refresh</a> </td>
						</tr>
					</table>
				</td>
				<td width="20%" class="blue-normal" align="left">
				</td>
				<td width="15%" class="blue-normal" align="left"><img src="../IMAGES/dot1px.gif" width="1" height="10"></td>
                  
                </tr>
              </table></td>
            <td width="40%"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr> 
				<td >&nbsp;&nbsp;</td>
			</tr>
			<tr> 
				<td bgcolor="8CA0D1"><img src="../IMAGES/DOT-01.GIF" width="1" height="1"></td>
			</tr>
			<tr> 
				<td class="red">&nbsp;&nbsp;<b><%=gMessage%></b> </td>
			</tr>
        </table>
<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "UTILISATION REPORT<br><span class='blue'>(Daily Utilisation -" & intDate & "/" & intMonth & "/" & intYear & ")</span>")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", "")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
%>
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr> 
		<td><table width="80" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
			<tr> 
				<td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
					<a href="javascript:_ReCalculate();" class="b">Re-Calculate</a> </td>
			</tr>
		</table></td>
	</tr>
	<tr> 
		<td ><img src="../IMAGES/DOT1.GIF" width="1" height="5"></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="5">
                <tr> 
                <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">Manager</td>
                <td class="blue" align="center" width="25%" bgcolor="#E7EBF5" rowspan="2">Ultilzation previous month</td>
                <td class="blue" align="center" width="25%" bgcolor="#E7EBF5" rowspan="2">Ultilzation previous month</td>
                <td class="blue" align="center" width="30%" bgcolor="#E7EBF5" colspan="4" >Ultilzation per week / Month </td>
                <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">Average</td>
              </tr>
              <tr> 
                <td class="blue" align="center" width="8%" bgcolor="#E7EBF5">W1</td>
                <td class="blue" align="center" width="8%" bgcolor="#E7EBF5">W2</td>
                <td class="blue" align="center" width="8%" bgcolor="#E7EBF5">W3</td>
                <td class="blue" align="center" width="8%" bgcolor="#E7EBF5">W4</td>
              </tr>
<%'=strLast%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<%			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
%>
<input type="hidden" name="txtBillable" value="">
</form>
</body>
</html>