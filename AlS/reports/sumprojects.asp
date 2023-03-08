<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
'*********************************************************
'Generate report
'*********************************************************
Function ATSsql(byval strF,byval strT)
	dim strATS
	strATS="SELECT tdate, assignmentID, staffID, Hours, Overtime FROM "
	If year(strF)<> year(strT) then
		
		For ii=year(strF) To year(strT)
			strATS=strATS & selectTable(ii)
			If ii<>	year(strT) then
				strATS=strATS & " UNION ALL SELECT tdate, assignmentID, staffID, Hours, Overtime FROM "
			end if
		Next
	else
		strATS=strATS & selectTable(year(strT))
	end if
	ATSsql=strATS
End function
'****************************************
' Function: outbody
' Description: 
' Parameters: array data, page size, which page
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function OutBody(ByRef arrSrc, ByVal PSize, ByVal Whichpage)
Dim strFullname, strsavename,intcolspan,strAPK
	strOut = ""
	i = (Whichpage - 1)*PSize	
	lastU = Ubound(arrSrc, 2)
	cnt = 0
	intcolspan=3
	if intDetail=1 then intcolspan=4
	Do Until i>lastU
		cnt = cnt + 1
		if arrSrc(5, i) = 0 then
			strColor = "#FFFFFF"
			if intDetail=1 then
				strAPK=showlabel(Mid(arrSrc(0, i), 1, Instr(arrSrc(0, i), chr(13)))) & _
					 showlabel(Mid(arrSrc(0, i), Instr(arrSrc(0, i), chr(13))+1, len(arrSrc(0, i))))
			else
				strAPK=trim(mid(arrSrc(0, i),1,Instr(arrSrc(0, i), chr(13))))
				if strAPK<>"" then strAPK = ucase(replace(right(strAPK,len(strAPK)-10),"_", " ")) & "<br>"
			end if			
			
			strTmp = "<td valign='top' class='blue-normal'>&nbsp;" & strAPK & "</td>" &_
					"<td valign='top' class='blue-normal'>&nbsp;" & showlabel(arrSrc(1, i)) & "</td>" &_
					"<td valign='top' class='blue-normal'>&nbsp;" & showlabel(arrSrc(2, i)) & "</td>" 
			if intDetail=1 then	strTmp=strTmp & "<td valign='top' class='blue-normal'>&nbsp;" & showlabel(arrSrc(3, i)) & "</td>" 
			
			strTmp=strTmp & "<td valign='top' class='blue-normal' align='right'>" & FormatNumber(arrSrc(4, i), 2) & "</td>"
		else
			strColor = IIF(arrSrc(5, i) = 1, "#FFE1E1","#FFF2F2")
			strTmp = "<td valign='top' colspan='"& intcolspan &"' class='blue-normal' align='right'>" & arrSrc(3, i) & "</td>" &_
					"<td valign='top' width='12%' class='blue' align='right'>" & FormatNumber(arrSrc(4, i), 2) & "</td>"			
		end if
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" & strTmp & "</tr>"
		i = i + 1
		if cnt = pSize then exit do
	Loop
	Outbody = strOut
End Function

'****************************************
' Function: selectfullmonth
' Description: 
' Parameters:
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectfullmonth(ByVal vname, Byval vselected)
	arrlongmon  = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	strlst = "<select name='" & vname & "' size='1' height='26px' width='75px' " &_
			"style='width:75px;height=24px;' class='blue-normal' onClick='document.frmreport.opttime[1].checked=true;'>"
	strTmp = ""
	For i = 1 to 12
		strTmp1 = CStr(i)
		if len(strTmp1) = 1 then strTmp1 = "0" & strTmp1 end if
		if i = vselected then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & strTmp1 & "' " & strSel & ">" & arrlongmon(i-1) & "</option>"
	Next
	strlst = strlst & strTmp & "</select>"
	set arrlongmon = nothing
	selectfullmonth = strlst
end function
'****************************************
' Function: selectyear
' Description: 
' Parameters:
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectlistyear(ByVal vname, Byval vselected, Byref arrSrc)
	strlst = "<select name='" & vname & "' size='1' height='26px' width='70px' " &_
			"style='width:70px;height=24px;' class='blue-normal' onClick='document.frmreport.opttime[1].checked=true;'>"
	strTmp = ""
	For i = 0 to ubound(arrSrc)
		if arrSrc(i) = int(vselected) then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & arrSrc(i) & "' " & strSel & ">" & arrSrc(i) & "</option>"
	Next
	strlst = strlst & strTmp & "</select>"
	selectlistyear = strlst
end function

	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage
	Dim intDetail,intDepart
	dim strBillableTask

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
'--------------------------------------------------
' Check VIEWALL project right
' User can update all project
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgViewAll = False
	Else
		varGetRight = session("RightOn")
		fgViewAll = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "View all projects" Then
				fgViewAll = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If	
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
' Get Full Name and Job Title
'----------------------------------

	Set objEmployee = New clsEmployee
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'><a href='../welcome.asp?choose_menu=B' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				"<a class='c' href='javascript:_print();' onMouseOver='self.status=&quot;Print report&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Print</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Preparing data
'--------------------------------------------------
	Call freeListpro
	Call freeProInfo
	Call freeAssignment
	Call freeAssignRight
	Call freeListEmp
	Call freeShort
	Call freeSinglepro

if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
	Call freeSumpro
end if

stract = Request.QueryString("act")
if stract = "REFRESH" then session("READYSUMPRO") = false

stropt = Request.Form("opttime")
strtypepro = Request.Form("lsttypepro")
if strtypepro = "" then strtypepro = "0"
select case strtypepro
	case "0" strtype = ""
	case "1" strtype = "AND f.Projectkey2 = 1" 'billable
	case "2" strtype = "AND (f.Projectkey2 = 5 OR f.Projectkey2 = 7)"	'non-billable
end select

intDepart=Request.Form("lbdepartment")
if intDepart="" then intDepart=0

intDetail=1

strprojectkey = trim(Request.Form("txtsearch"))
if strprojectkey <> "" then strtype = strtype & " AND e.ProjectID LIKE '%" & strprojectkey & "%'"
if cint(intDepart)>0 then strtype = strtype & " AND e.DepartmentID= " & intDepart 

if not fgViewAll then strtype = strtype & " AND " & getWherePhase("e",session("USERID"))

strwhere = ""
gMessage = ""

strfrom = ""
strto = ""
if stropt = "0" then 'from to
	strfrom = Request.Form("txtfrom")
	strto = Request.Form("txtto")
	arrTmp = split(strfrom, "/")	
	strfrom_I = cdate(arrTmp(1) & "/" & arrTmp(0) & "/" & arrTmp(2))
	arrTmp = split(strto, "/")	
	strto_I = cdate(arrTmp(1) & "/" & arrTmp(0) & "/" & arrTmp(2))

else 'month
	strmonth = Request.Form("lstmonth")
	if strmonth = "" then strmonth = month(date())
	arrlongmon  = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	strnameofmonth = arrlongmon(int(strmonth)-1)
	set arrlongmon = nothing
	stryear = Request.Form("lstyear")
	if stryear = "" then stryear = Cstr(year(date())) ' for first run
	strfrom_I = cdate(strmonth & "/1/" & stryear)
	strto_I=DateAdd("m",1 , strfrom_I) -1
	
end if

strwhere="(SELECT * FROM (" & ATSsql(strfrom_I,strto_I) & ") ATS WHERE Tdate BETWEEN '" & strfrom_I & "' AND '" & strto_I & "') a "

if stropt = "0" then 
	strfromto = "FROM " & strfrom & " to " & strto
else 
	strfromto = strnameofmonth & " - " & stryear
end if
strprintdate = FormatDateTime(Date, 1) 'day(date()) & "/" & month(date()) & "/" & year(date())
strlstmonth = selectfullmonth("lstmonth", int(strmonth))

If isEmpty(session("READYSUMPRO")) or session("READYSUMPRO")=false then
  Dim objConn                 '-- The ADO Connection to the Database
  Dim objRs1             '-- The ADO Parent Recordset (Publishers)
  Dim objRs2         '-- The ADO child Recordset (Titles)
  Dim strShape                '-- The SHAPE Syntax
  Dim strConn                 '-- Connection String to the Database
  '-- Create the ADO Objects
  set objConn = Server.CreateObject("ADODB.Connection")
  set objRs1 = Server.CreateObject("ADODB.Recordset")
  set objRs2 = Server.CreateObject("ADODB.Recordset")

  strConn = Application("g_strConnectShape")

  '-- Define the Shape Provider
  objConn.Provider = "MSDataShape"

  '-- Open the Connection
  objConn.Open strConn
  strShape = "SHAPE(SHAPE {SELECT ISNULL(c.TaskID, 0) fg, e.projectID, e.projectname, h.DepartmentCode, d.Firstname + ' ' + isnull(d.middlename,'')+ ' ' + isnull(d.lastname,'') as Fullname, " &_
			"c.Subtaskname,c.FgBillable, a.Tdate, a.Hours + a.Overtime as sumhour, f.Projectkey2, CONVERT(varchar(50), f.DateTransfer, 106) DateTransfer " &_
			"FROM " & strwhere &_
			"LEFT JOIN ATC_Assignments b ON a.AssignmentID = b.AssignmentID " &_
			"LEFT JOIN ATC_Tasks c On b.SubTaskID = c.SubTaskID " &_
			"LEFT JOIN ATC_PersonalInfo d ON a.StaffID = d.PersonID " &_
			"LEFT JOIN ATC_Projects e ON c.ProjectID = e.ProjectID " &_
			"INNER JOIN ATC_Department h ON e.DepartmentID = h.DepartmentID " &_
			"LEFT JOIN ATC_ProjectStage f ON c.ProjectID = f.ProjectID " &_
			"AND f.DateTransfer = (SELECT max(DateTransfer) FROM ATC_ProjectStage WHERE DateTransfer " &_
			"<= a.Tdate and ProjectID = c.ProjectID) WHERE a.AssignmentID>1 and e.ProjectID is not null " & strtype & " " &_
			"ORDER BY e.projectID} as rsdetail " &_
			"COMPUTE rsdetail, ANY(rsdetail.fg) fg, ANY(rsdetail.ProjectName) ProjectName, ANY(rsdetail.DepartmentCode) DepartmentCode, " &_
			"SUM(rsdetail.sumHour) as perhour BY SubTaskname,FgBillable, Fullname, DateTransfer, Projectkey2, ProjectID) rsTask " &_
			"COMPUTE rsTask, SUM(rsTask.perhour) as vhour BY DateTransfer, Projectkey2, ProjectID"

  objRs1.ActiveConnection = objConn
  objRs1.Open strShape
  
  if not objRs1.EOF then 
	objRs1.Sort = "ProjectID, DateTransfer ASC"
	Dim arrData()
	cnt = -1  
	fgPass = true
	Overall = 0
		  dblBillableHours=0
	Do Until objRs1.EOF
	  Set objRs2 = objRs1("rsTask").Value
	  
	  strTaskName=""
	  strBillableTask=""
	  'Total hours per subtask
	  dblHourSubTask=0

	  	  
	  Do Until objRs2.EOF

	  	  	
	  	if strTaskName<> objRs2("subtaskname") then
	  						
	  		if  strTaskName<>"" then		
	  			cnt = cnt + 1 	
				Redim preserve arrData(5, cnt)
				arrData(0, cnt) = ""
				arrData(1, cnt) = ""
				arrData(2, cnt) = ""
				arrData(3, cnt) = "Sub total of <b>" & strTaskName & " (" & strBillableTask & ")</b> :"
				arrData(4, cnt) = dblHourSubTask
				arrData(5, cnt) = 2 'flag notify that this row is subtotal line	
			end if
					
			strTaskName= objRs2("subtaskname") 
			if cint(objRs2("fgBillable")) = 0 then
				strBillableTask = "Non - Billable"
			elseif cint(objRs2("fgBillable"))=1 then
				strBillableTask = "Billable"
			else
				strBillableTask = "Risked Billable"
			end if
	
			dblHourSubTask=0
	  	end if
        
        
	  	cnt = cnt + 1 	
	  	
		Redim preserve arrData(5, cnt)
		
		arrData(2, cnt) = objRs2("Fullname")
		if objRs2("fg") = 0 and fgPass=false then
			arrData(3, cnt) = "_"
		else
			arrData(3, cnt) = objRs2("subtaskname")  
		end if
		arrData(4, cnt) = objRs2("perhour")
		
		dblHourSubTask=dblHourSubTask + objRs2("perhour")
		
		if cint(objRs2("fgBillable"))=1 then 
		    dblBillableHours=dblBillableHours +  objRs2("perhour")
		    'response.write "<br>" & strTaskName & ":" & objRs2("perhour") & "--" & 	dblBillableHours
		    
		end if


	
		arrData(5, cnt) = 0
		if fgPass then
			'arrData(0, cnt) = objRs2("DepartmentCode") & "_" & objRs2("Projectkey2") & "_" & objRs2("projectID") & chr(13) & "(" & objRs2("DateTransfer") & ")"
			arrData(0, cnt) = objRs2("projectID") '& chr(13) & "(" & objRs2("DateTransfer") & ")"
			arrData(1, cnt) = objRs2("Projectname")
			arrData(3, cnt) = objRs2("subtaskname")
		
			fgPass = false
		end if			
'Response.Write 		strTaskName & "--" & 	strBillableTask & "<br>"		
	    objRs2.MoveNext
	  Loop
	  
		cnt = cnt + 1 	
		Redim preserve arrData(5, cnt)
		arrData(0, cnt) = ""
		arrData(1, cnt) = ""
		arrData(2, cnt) = ""
		arrData(3, cnt) = "Sub total of <b>" & strTaskName & " (" & strBillableTask & ")</b> :"
		arrData(4, cnt) = dblHourSubTask
		arrData(5, cnt) = 2 'flag notify that this row is subtotal line
	  
	  cnt = cnt + 1
	  Redim preserve arrData(5, cnt)
		arrData(0, cnt) = ""
		arrData(1, cnt) = ""
		arrData(2, cnt) = ""
		arrData(3, cnt) = "Total: <b></b>"
		arrData(4, cnt) = objRs1("vhour")
		arrData(5, cnt) = 1 'flag notify that this row is summary line
		Overall = Overall + objRs1("vhour")
		fgPass = true
	  objRs1.MoveNext
	  
	Loop
	'row for overall total
	cnt = cnt + 1
	Redim preserve arrData(5, cnt)
	arrData(0, cnt) = ""
	arrData(1, cnt) = ""
	arrData(2, cnt) = ""
	arrData(3, cnt) = "<b>Overall Total: </b>"
	arrData(4, cnt) = Overall
	arrData(5, cnt) = 1
	
	'row for total billable hours
	cnt = cnt + 1
	Redim preserve arrData(5, cnt)
	arrData(0, cnt) = ""
	arrData(1, cnt) = ""
	arrData(2, cnt) = ""
	arrData(3, cnt) = "<b>Total billable hours: </b>"
	arrData(4, cnt) = dblBillableHours
	arrData(5, cnt) = 1
	
	session("arrSumPro") = arrData
	session("NumPageSumPro") = PageCount(arrData, PageSize)
	session("CurpageSumPro") = 1
	


	on error resume next
	objRs1.Close
	objRs2.Close
	objConn.Close
	set objRs1 = nothing
	set objRs2 = nothing
	set objConn = nothing

	if Err.number>0 then
		gMessage = Err.description
		Err.Clear
	end if
  else
	session("NumPageSumPro") = 0
	session("CurpageSumPro") = 0
	session("arrSumPro") = empty
  end if 'test have data
  session("READYSUMPRO") = true
end if

'--------------------
' get table ATC_Index
'--------------------
if isEmpty(session("arryearValid")) then
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	Dim arryearValid()
	cnt = -1
	If objDb.dbConnect(strConnect) then
		strQuery = "SELECT start_date FROM ATC_Index"
		If objDb.runQuery(strQuery) Then
			If not objDb.noRecord then
				Do until objDb.rsElement.EOF
					cnt = cnt + 1
					Redim preserve arryearValid(cnt)
					arryearValid(cnt) = year(objDb.rsElement(0))
					objDb.MoveNext
				Loop
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
	'For current year
	cnt = cnt + 1
	Redim preserve arryearValid(cnt)
	arryearValid(cnt) = year(now())
	session("arryearValid") = arryearValid
end if
arrtmp = session("arryearValid")
strlstyear = selectlistyear("lstyear", stryear, arrTmp)
set arrtmp = nothing

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

varNavi = Request.QueryString("navi")
if varNavi <> "" then
	tmpi = session("CurPageSumPro")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPageSumPro") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPageSumPro")
			end if
	End select
	session("CurPageSumPro") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPageSumPro") = CInt(varGo)

if not isEmpty(session("arrSumPro")) then
  arrData = session("arrSumPro")
  strLast = OutBody(arrData, PageSize, session("CurpageSumPro")) 
end if

'--------------------------------------------------
' Initialize department array
'--------------------------------------------------	
If Not isEmpty(session("varDepartment")) Then
		varDepartment = session("varDepartment")
Else
		varDepartment = GetDepartment()
		if not isEmpty(varDepartment) then	session("varDepartment") = varDepartment
End If
if IsArray(varDepartment) then intNum = Ubound(varDepartment,2)
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

<link href="../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>
<script src="../jQuery/jquery.min.js"></script>
<script src="../jQuery/jquery-ui.min.js"></script>
<link href="../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<script>
	$(function() {
		var dates = $( "#txtfrom, #txtto" ).datepicker({
		dateFormat: "dd/mm/yy",
		minDate:"1/1/2000",
		maxDate: "31/12/<%=year(Date())%>",
        onSelect: function( selectedDate ) {
                                    var option = this.id == "txtfrom" ? "minDate" : "maxDate",
                                        instance = $( this ).data( "datepicker" ),
                                        date = $.datepicker.parseDate(
                                            instance.settings.dateFormat ||
                                            $.datepicker._defaults.dateFormat,
                                            selectedDate, instance.settings );
                                    dates.not( this ).datepicker( "option", option, date );
                                }
            });
	});
</script>

<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script>
var objWindowSumPro;

function _print() { //v2.0
var str1 = "<%=strfromto%>";
	str1 = escape(str1);
var str2 = "<%=strprintdate%>";
	str2 = escape(str2);
var fgprint = <%=session("NumPageSumPro")%>;
if (fgprint!=0) {
	window.status = "";
	strFeatures = "top="+(screen.height/2-275)+",left="+(screen.width/2-390)+",width=800,height=550,toolbar=no," 
	            + "menubar=yes,location=no,directories=no,scrollbars=yes,status=yes";
	if ((objWindowSumPro) && (!objWindowSumPro.closed)) {
		objWindowSumPro.focus();
	
	} else {
		objWindowSumPro = window.open("p_sumprojects.asp?fromto=" + str1 + "&printdate=" + str2, "MyNewWindow", strFeatures);
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
	if (document.frmreport.opttime[0].checked)
	{
		if (isnull(document.frmreport.txtfrom.value)==true)
		{
			alert("Please enter startdate before click here.")
			document.frmreport.txtfrom.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtfrom.value)==false)
			{			
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtfrom.focus();
				return false;
			}
		}
		
		if (isnull(document.frmreport.txtto.value)==true)
		{
			alert("Please enter enddate before click here.")
			document.frmreport.txtto.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtto.value)==false)
			{
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtto.focus();
				return false;
			}
		}
		
		if (comparedate(document.frmreport.txtfrom.value,document.frmreport.txtto.value)==false)
		{
			alert("The startdate must be less than the finishdate.")
			document.frmreport.txtfrom.focus();
			return false;
		}
	}	
	return true;
}

function _submit() {
	if(checkdata()==true) {
		document.frmreport.action = "sumprojects.asp?act=REFRESH";
		document.frmreport.target = "_self" ;
		document.frmreport.submit();
	}
}

function next() {
var curpage = <%=session("CurPageSumPro")%>;
var numpage = <%=session("NumPageSumPro")%>;
	if (curpage < numpage) {
		document.frmreport.action = "sumprojects.asp?navi=NEXT"
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageSumPro")%>;
var numpage = <%=session("NumPageSumPro")%>;
	if (curpage > 1) {
		document.frmreport.action = "sumprojects.asp?navi=PREV";
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageSumPro")%>;
	var curpage = <%=session("CurPageSumPro")%>;
	var intpage = document.frmreport.txtpage.value;
	intpage = parseInt(intpage, 10);
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmreport.action = "sumprojects.asp?Go=" + intpage;
		document.frmreport.target = "_self";
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
	
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%"  align="center" >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="31%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td valign="top" width="0%" class="blue" height="33">&nbsp;</td>
                  <td width="4%" class="blue" align="right" height="33"> 
                    <input type="radio" name="opttime" value="0" <%if stropt="0" then%>checked<%end if%>>
                  </td>
                  <td class="blue-normal" width="8%" height="33">From </td>
                  <td class="blue-normal" width="40%" height="33"> 
                    <input type="text" name="txtfrom" id="txtfrom" class="blue-normal"  size="5" style="width:60" value="<%=strfrom%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                  <td width="8%" class="blue-normal" height="33"> To </td>
                  <td class="blue-normal" height="33" width="40%"> 
                    <input type="text" name="txtto" id="txtto" class="blue-normal" size="5" style="width:60" value="<%=strto%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                </tr>
                <tr> 
                  <td valign="top" width="0%" class="blue">&nbsp;</td>
                  <td width="4%" class="blue" align="right"> 
                    <input type="radio" name="opttime" value="1" <%if stropt="1" then%>checked<%end if%>>
                  </td>
                  <td class="blue-normal" width="8%">Month </td>
                  <td class="blue-normal" width="40%"> 
<%Response.Write strlstmonth%>
                  </td>
                  <td width="8%" class="blue-normal"> Year </td>
                  <td class="blue-normal" width="40%">
<%Response.Write strlstyear%>
                  </td>
                </tr>
              </table>
            </td>
            <td width="49%"> 
			<table width="100%" border="0" cellpadding="1" cellspacing="0">
                <tr> 
                  <td class="blue-normal" width="25%" height="33">Search for</td>
                  <td class="blue-normal" colspan="3" >
					<input type="text" name="txtsearch" class="blue-normal" size="10" style="width:100px;" value="<%=strprojectkey%>">
					
					</td>
                </tr>
               <tr> 
                  <td class="blue-normal" height="33">Project Type</td>
                  <td class="blue-normal" height="33">
					<select name="lsttypepro" style="width:100px" class="blue-normal">
						<option value="0" <%if strtypepro="0" then%>selected<%end if%>>&nbsp;</option>
						<option value="1" <%if strtypepro="1" then%>selected<%end if%>>Billable project</option>
						<option value="2" <%if strtypepro="2" then%>selected<%end if%>>Non-billable project</option>
					</select> </td>
                  <td class="blue-normal" >Department</td>
                  <td class="blue-normal" >
					<select id="lbdepartment" size="1" name="lbdepartment" style="width:150px" class="blue-normal">
					  <option value="0" <%if cint(intDepart)=0 then%>selected<%end if%>>&nbsp;</option>
						<%
								If intNum >= 0 Then
								    For ii = 0 To intNum
						%>                    
											  <option <%If CInt(intDepart)=CInt(varDepartment(0,ii)) Then%> selected <%End If%> value="<%=varDepartment(0,ii)%>"><%=showlabel(varDepartment(1,ii))%></option>
						<%
									Next
								End If	
						%>					

				    </select></td>
                </tr>       
              </table>
             
            </td>
            <td width="20%"> 
              <table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
                <tr> 
                  <td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
                      <a href="javascript:_submit();" class="b">Submit</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
			<td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		  </tr>
          <tr> 
            <td bgcolor="8CA0D1"><img src="../IMAGES/DOT-01.GIF" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>&nbsp; </td>
          </tr>
        </table>
    		<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Projects")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strfromto)
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
			%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="3">
<%if intDetail=1 then%>              
                <tr> 
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5">Project ID </td>
                  <td class="blue" align="center" width="22%" bgcolor="#E7EBF5">Project Name </td>
                  <td class="blue" align="center" width="25%" bgcolor="#E7EBF5">Full Name </td>
                  <td class="blue" align="center" width="26%" bgcolor="#E7EBF5">Sub-task Description </td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5">Hours<br> Worked </td>
                </tr>
<%else%>
				<tr> 
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5">APK </td>
                  <td class="blue" align="center" width="35%" bgcolor="#E7EBF5">Description </td>
                  <td class="blue" align="center" width="35%" bgcolor="#E7EBF5">Full Name </td>
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5">Hours</td>
				</tr>
<%end if%>

<%Response.Write strLast%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<%if session("NumPageSumPro")>0 then%>
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="20" align="center">
    <tr> 
      <td align="right" bgcolor="#E7EBF5"> 
        <table width="70%" border="0" cellspacing="1" cellpadding="0" height="18">
          <tr> 
            <td align="right" valign="middle" width="37%" class="blue-normal">Page 
            </td>
            <td align="center" valign="middle" width="13%" class="blue-normal"> 
              <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageSumPro")%>" size="2" style="width:50">
            </td>
            <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();"  onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
            </td>
            <td align="right" valign="middle" width="15%" class="blue-normal">Pages 
               <%=session("CurpageSumPro")%>/<%=session("NumpageSumPro")%>&nbsp;&nbsp;</td>
            <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();"  onMouseOver="self.status='Previous page'; return true;" onMouseOut="self.status='';">Previous</a> 
              /<a href="javascript:next();" onMouseOver="self.status='Next page'; return true;" onMouseOut="self.status='';"> Next</a>&nbsp;&nbsp;&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
</table>
<%end if%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>
</body>
</html>