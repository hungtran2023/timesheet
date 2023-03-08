<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
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
	lastU = Ubound(arrSrc, 2)
	strOut = ""
	i = (Whichpage - 1)*PSize
	cnt = 0
	strsavetask = ""
	Do Until i>lastU
		cnt = cnt + 1
		if arrSrc(7, i) = 0 then
			strColor = "#FFFFFF"
			if arrSrc(3, i) = strsavetask then
				strTaskname = "_"
			else
				strTaskname = arrSrc(3, i)
				strsavetask = arrSrc(3, i)
			end if
			strTmp = "<td valign='top' class='blue-normal'>&nbsp;" & showlabel(arrSrc(0, i)) & "</td>" &_
					"<td valign='top' class='blue-normal'>&nbsp;" & showlabel(arrSrc(1, i)) & "</td>" &_					
					"<td valign='top' class='blue-normal'>" & showlabel(arrSrc(2, i)) & "</td>" &_
					"<td valign='top' class='blue-normal'>&nbsp;" & showlabel(strTaskname) & "</td>" &_
					"<td valign='top' class='blue-normal' align='right'>" & FormatNumber(arrSrc(4, i), 2) & "</td>" &_
					"<td valign='top' class='blue-normal' align='right'>" & FormatNumber(arrSrc(5, i), 2) & "</td>" &_
					"<td valign='top' class='blue-normal' align='right'>" & FormatNumber(arrSrc(6, i), 2) & "</td>"
		else
			strColor = "#FFE1E1"
			strsavetask = ""
			if arrSrc(7, i) <= 2 then
			strColor = IIF(arrSrc(5, i) = 1, "#FFE1E1","#FFF2F2")
			strTmp = "<td valign='top' colspan='4' class='blue-normal' align='right'>" & arrSrc(3, i) & "</td>" &_
					"<td valign='top' class='blue' align='right'>" & FormatNumber(arrSrc(4, i), 2) & "</td>" &_
					"<td valign='top' class='blue' align='right'>" & arrSrc(5, i) & "</td>" &_
					"<td valign='top' class='blue' align='right'>" & FormatNumber(arrSrc(6, i), 2) & "</td>"
			else
			strTmp = "<td valign='top' colspan='6' class='blue' align='right'>" & arrSrc(3, i) & "</td>" & _
					"<td valign='top' class='blue' align='right'>" & arrSrc(6, i) & "</td>"
			end if
			
		end if
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" & strTmp & "</tr>"
		i = i + 1
		if cnt = pSize then exit do
	Loop
	Outbody = strOut
End Function
'****************************************
' Function: selectpro
' Description: 
' Parameters: vselected is a value need to be selected
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectpro(ByRef rsSrc, ByVal vname, Byval vselected)
	strlst = "<select name='" & vname & "' size='1' height='26px' width='190px' " &_
			"style='width:190px;height=24px;' class='blue-normal'>"
			
	if vselected = "" then
		strTmp = "<option value='' selected>-- Select a project --</option>"
	end if
	rsSrc.MoveFirst
	Do until rsSrc.EOF
		if rsSrc(0) = vselected then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" &  rsSrc(0) & "' " & strSel & ">" &  rsSrc(0) & " (" &  rsSrc(1) & ")</option>"
		rsSrc.MoveNext
	Loop
	strlst = strlst & strTmp & "</select>"
	selectpro = strlst
end function

	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage
	dim strBillableTask
	Dim dblBillableHours,dblNonBillableHours,dblOverrunHours
	dim arrSubtaskType  
	
	arrSubtaskType= Array("Non-Billable","Billable","Risked Billable")


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
'		Response.Clear
		Response.Redirect("../welcome.asp")
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
'----------------------------------
' Get Full Name and Job Title
'----------------------------------
	Set objEmployee = New clsEmployee	
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
	Call freeSumpro
if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
	Call freeSinglepro
end if

stract = Request.QueryString("act")
if stract = "REFRESH" then 
	session("READYSINGLEPRO") = false
	strproID = Request.Form("lstpro")
end if
strprintdate = FormatDateTime(Date, 1)	'day(date()) & "/" & month(date()) & "/" & year(date())

If isEmpty(session("READYSINGLEPRO-L")) or session("READYSINGLEPRO-L")=false then
	strConnect = Application("g_strConnect")
	rsPro = Server.CreateObject("ADODB.Recordset")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
	strQuery = "SELECT ProjectID, ProjectName,CSOCompleted,CSOMainHours FROM ATC_Projects WHERE fgDelete = 0 AND left(ProjectID, 1) <> '_'"
	if not fgViewAll then strQuery = strQuery & " AND " & getWherePhase("ATC_Projects",session("USERID"))


	
'Response.Write 	strQuery
'Response.End
	If objDb.openRec(strQuery) Then
	  objDb.recDisConnect
	  if not objDb.noRecord then
		set rsPro = objDb.rsElement.Clone
		session("READYSINGLEPRO-L") = TRUE
		rsPro.MoveFirst
		set session("rsProSINGLEPRO-L") = rsPro
		session("fgShowPRO-L") =  0 ' show all
	  else
		set rsPro = nothing
		rsPro = Empty
	  End if
	  objDb.CloseRec
	Else
	  gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
Else
	set rsPro = session("rsProSINGLEPRO-L")
	
	if not IsEmpty(session("rsResultSINGLEPRO-L")) then
		set rsResult = session("rsResultSINGLEPRO-L")
	end if 
End if

If (isEmpty(session("READYSINGLEPRO")) or session("READYSINGLEPRO")=false) and strproID<>"" and gMessage = "" then
  strConnect = Application("g_strConnect") 
  Set objDb = New clsDatabase
  If objDb.dbConnect(strConnect) then
	strQueryExtra = "SELECT TDate, StaffID, AssignmentID, Hours, OverTime FROM ATC_TimeSheet WHERE AssignmentID>1"
	strQuery = "SELECT TMS_Table FROM ATC_Index"
	If not objDb.runQuery(strQuery) Then
		gMessage = objDb.strMessage
	Else
		Do until objDb.rsElement.EOF
			strQueryExtra = strQueryExtra & " UNION SELECT TDate, StaffID, AssignmentID, Hours, OverTime FROM " & objDb.rsElement(0) &_
							" WHERE AssignmentID>1 "
			objDb.rsElement.MoveNext
		Loop
		objDb.rsElement.Close()
	End if
	objDb.dbDisconnect
  Else
	gMessage = objDb.strMessage
  End if
  set objDb = nothing

'Response.Write strQueryExtra
  
  if gMessage = "" then 'no error
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

	  strShape = "SHAPE(SHAPE {SELECT isnull(c.TaskID, 0) fg, e.projectID, e.projectname, e.CSOCompleted,e.CSOMainHours,h.DepartmentCode, d.Firstname + ' ' + isnull(d.middlename,'')+ ' ' + isnull(d.lastname,'') as Fullname, " &_
				"c.Subtaskname,c.fgBillable, a.Tdate,a.Hours,a.Overtime, a.Hours + a.Overtime as sumhour, f.Projectkey2, f.DateTransfer DateTransferSort, CONVERT(varchar(50), f.DateTransfer, 106) DateTransfer, " &_
				"HourTransfer, g.Firstname + ' ' + isnull(g.middlename,'')+ ' ' + isnull(g.lastname,'') as Ownersystem FROM (" & strQueryExtra & ") a " &_
				"INNER JOIN ATC_Assignments b On a.AssignmentID = b.AssignmentID " &_
				"INNER JOIN ATC_Tasks c On b.SubTaskID = c.SubTaskID " &_
				"INNER JOIN ATC_PersonalInfo d On a.StaffID = d.PersonID " &_
				"INNER JOIN ATC_Projects e On c.ProjectID = e.ProjectID " &_
				"INNER JOIN ATC_Department h On e.DepartmentID = h.DepartmentID " &_
				"INNER JOIN (SELECT Projectkey2, ProjectID, DateTransfer, HourTransfer, StaffID FROM ATC_ProjectStage) f On c.ProjectID = f.ProjectID " &_
				"AND f.DateTransfer = (SELECT max(k.DateTransfer) FROM (SELECT DateTransfer FROM ATC_ProjectStage WHERE DateTransfer " &_
				"<= a.Tdate AND ProjectID = c.ProjectID) k) LEFT JOIN ATC_PersonalInfo g On f.StaffID = g.PersonID " &_
				"WHERE e.ProjectID is not null AND e.ProjectID = '" & strproID & "'} as rsdetail " &_
				"COMPUTE rsdetail, ANY(rsdetail.fg) fg, ANY(rsdetail.ProjectName) ProjectName, ANY(rsdetail.HourTransfer) HourTransfer, ANY(rsdetail.DepartmentCode) DepartmentCode," &_
				"ANY(rsdetail.Ownersystem) Ownersystem, ANY(rsdetail.DateTransfer) DateTransfer,SUM(rsdetail.Hours) as normalhour,SUM(rsdetail.Overtime) as OThour, SUM(rsdetail.sumHour) as perhour BY SubTaskname,fgBillable, Fullname, DateTransferSort, Projectkey2, ProjectID) rsTask " &_
				"COMPUTE rsTask, ANY(rsTask.HourTransfer) HourTransfer, ANY(rsTask.DateTransfer) DateTransfer, ANY(rsTask.Ownersystem) Ownersystem,SUM(rsTask.normalhour) as vnorhour,SUM(rsTask.OThour) as vOThour, SUM(rsTask.perhour) as vhour BY DateTransferSort, Projectkey2, ProjectID"

	  objRs1.ActiveConnection = objConn
	  objRs1.Open strShape
	  if not objRs1.EOF then 
		objRs1.Sort = "DateTransferSort"
		Dim arrData()
		cnt = -1  
		fgPass = true
		Overall = 0
		OverallOT=0
		intlastsum = -1
		
		dblBillableHours=0
		dblNonBillableHours=0
		dblOverrunHours=0
		
		Do Until objRs1.EOF
		  Set objRs2 = objRs1("rsTask").Value
		  objRs2.Sort = "fgBillable,SubTaskname, Fullname ASC"
		  
		  strTaskName=""
		  strBillableTask=""
		  'Total hours per subtask
		  dblHourSubTask=0
		  dblHourSubTaskOT=0
		  
		  Do Until objRs2.EOF
		  
			if strTaskName<> objRs2("subtaskname") then
	  			if  strTaskName<>"" then
	  				cnt = cnt + 1 	
	  				
	  				Redim preserve arrData(7, cnt)
					arrData(0, cnt) = ""
					arrData(1, cnt) = ""
					arrData(2, cnt) = ""
					arrData(3, cnt) = "Sub total of <b>" & strTaskName & " (" & strBillableTask & ")</b> :"
					arrData(4, cnt) = dblHourSubTask
					arrData(5, cnt) = dblHourSubTaskOT
					arrData(6, cnt) = dblHourSubTask + dblHourSubTaskOT
					if strBillableTask = arrSubtaskType(0) then
						dblNonBillableHours=dblNonBillableHours + dblHourSubTask + dblHourSubTaskOT
					elseif strBillableTask = arrSubtaskType(1) then
						dblBillableHours=dblBillableHours + dblHourSubTask + dblHourSubTaskOT
					else
						dblOverrunHours=dblOverrunHours + dblHourSubTask + dblHourSubTaskOT
					end if
					
					arrData(7, cnt) = 2

				end if
				strTaskName= objRs2("subtaskname")
				
				strBillableTask=arrSubtaskType(cint(objRs2("fgBillable")))
				
	
				dblHourSubTask=0
				dblHourSubTaskOT=0
	  		end if
		  
		  
		  	cnt = cnt + 1
			Redim preserve arrData(7, cnt)
			arrData(2, cnt) = objRs2("Fullname")
			if objRs2("fg") = 0 and fgPass=false then
				arrData(3, cnt) = "_"
			else
				arrData(3, cnt) = objRs2("subtaskname")
			end if
			arrData(4, cnt) = objRs2("normalhour")
			arrData(5, cnt) = objRs2("OThour")
			arrData(6, cnt) = objRs2("perhour")
			
			dblHourSubTask = dblHourSubTask + objRs2("normalhour")			
			dblHourSubTaskOT = dblHourSubTaskOT + objRs2("OThour")
			
			arrData(7, cnt) = 0
			if fgPass then
				'arrData(0, cnt) = objRs2("DepartmentCode") & "_" & objRs2("Projectkey2") & "_" & objRs2("projectID")
				arrData(0, cnt) = objRs2("projectID")
				arrData(1, cnt) = objRs2("Projectname")
				arrData(3, cnt) = objRs2("subtaskname")
				fgPass = false
			end if			
		    objRs2.MoveNext
		  Loop
		  
		  cnt = cnt + 1 	
	  				
	  		Redim preserve arrData(7, cnt)
			arrData(0, cnt) = ""
			arrData(1, cnt) = ""
			arrData(2, cnt) = ""
			arrData(3, cnt) = "Sub total of <b>" & strTaskName & " (" & strBillableTask & ")</b> :"
			arrData(4, cnt) = dblHourSubTask
			arrData(5, cnt) = dblHourSubTaskOT
			arrData(6, cnt) = dblHourSubTask + dblHourSubTaskOT
			arrData(7, cnt) = 2
		  
		  	if strBillableTask = arrSubtaskType(0) then
				dblNonBillableHours=dblNonBillableHours + dblHourSubTask + dblHourSubTaskOT
			elseif strBillableTask = arrSubtaskType(1) then
				dblBillableHours=dblBillableHours + dblHourSubTask + dblHourSubTaskOT
			else
				dblOverrunHours=dblOverrunHours + dblHourSubTask + dblHourSubTaskOT
			end if
		  cnt = cnt + 1
		  Redim preserve arrData(7, cnt)
			arrData(0, cnt) = ""
			arrData(1, cnt) = ""
			arrData(2, cnt) = ""
			arrData(4, cnt) = objRs1("vnorhour") 'worked hours
			arrData(7, cnt) = 1
			if intlastsum = -1 then
				arrData(3, cnt) = "(Registered by <b>" & objRs1("Ownersystem") & ", " & objRs1("DateTransfer") & "</b>) Total: "
				'arrData(5, cnt) = ""
				'arrData(6, cnt) = arrData(4, cnt) 'total hours
				arrData(5, cnt) = objRs1("vOThour") 'worked hours
				arrData(6, cnt) = objRs1("vhour") 'worked hours
				intlastsum = cnt
			else
				arrData(3, cnt) = "(Transferred by <b>" & objRs1("Ownersystem") & ", " & objRs1("DateTransfer") & "</b>) Total: "
				'arrData(5, cnt) = objRs1("HourTransfer") 'in/out
				'arrData(6, cnt) = CSng(arrData(4, cnt)) + CSng(arrData(5, cnt)) 'total hours
				arrData(5, cnt) = objRs1("vOThour") 'worked hours
				arrData(6, cnt) = objRs1("vhour") 'worked hours
				'minus at "lastsum" row if hourTransfer>0
				'if CSng(objRs1("HourTransfer"))>0 then
					'arrData(5, intlastsum) = arrData(5, intlastsum) & " <" & objRs1("HourTransfer") & ">"
					'arrData(6, intlastsum) = CSng(arrData(6, intlastsum)) - CSng(objRs1("HourTransfer"))
				'end if
				intlastsum = cnt
			end if
			Overall = Overall + objRs1("vnorhour")
			OverallOT=OverallOT + objRs1("vOThour")
			fgPass = true
		  objRs1.MoveNext
		Loop
		'row for overall total
		cnt = cnt + 1
		Redim preserve arrData(7, cnt)
		arrData(0, cnt) = ""
		arrData(1, cnt) = ""
		arrData(2, cnt) = ""
		arrData(3, cnt) = "<b>Overall Total:</b> "
		arrData(4, cnt) = Overall
		arrData(5, cnt) = OverallOT
		arrData(6, cnt) = Overall + OverallOT
		arrData(7, cnt) = 1
		
		if dblNonBillableHours<>0 then
			cnt = cnt + 1
			Redim preserve arrData(7, cnt)
			arrData(0, cnt) = ""
			arrData(1, cnt) = ""
			arrData(2, cnt) = ""
			arrData(3, cnt) = "<b> " & arrSubtaskType(0) & " hours: </b> "
			arrData(4, cnt) = ""
			arrData(5, cnt) = ""
			arrData(6, cnt) = FormatNumber(dblNonBillableHours,2)
			arrData(7, cnt) = 3
		end if

		if dblBillableHours<>0 then
			cnt = cnt + 1
			Redim preserve arrData(7, cnt)
			arrData(0, cnt) = ""
			arrData(1, cnt) = ""
			arrData(2, cnt) = ""
			arrData(3, cnt) = "<b> " & arrSubtaskType(1) & " hours: </b> "
			arrData(4, cnt) = ""
			arrData(5, cnt) = ""
			arrData(6, cnt) = FormatNumber(dblBillableHours,2)
			arrData(7, cnt) = 3
		end if

		if dblOverrunHours<>0 then
			cnt = cnt + 1
			Redim preserve arrData(7, cnt)
			arrData(0, cnt) = ""
			arrData(1, cnt) = ""
			arrData(2, cnt) = ""
			arrData(3, cnt) = "<b> " & arrSubtaskType(2) & ": </b> "
			arrData(4, cnt) = ""
			arrData(5, cnt) = ""
			arrData(6, cnt) = FormatNumber(dblOverrunHours,2)
			arrData(7, cnt) = 3
		end if	
      
		'row for CSO Man-Hours
		if Mid(strproID,11,1)="L" Then
		    rsPro.Filter="ProjectID='" & strproID & "'"
		    cnt = cnt + 1
		    Redim preserve arrData(7, cnt)
		    arrData(0, cnt) = ""
		    arrData(1, cnt) = ""
		    arrData(2, cnt) = ""
		    arrData(3, cnt) = ""
		    if Mid(strproID,11,1)="L" Then	arrData(3, cnt) = "<b>CSO Man-Hours</b>"
		    arrData(4, cnt) = ""
		    arrData(5, cnt) = ""
		    arrData(6, cnt)= ""
		    if not isnull(rsPro("CSOMainHours")) AND Mid(strproID,11,1)="L" then arrData(6, cnt) = FormatNumber(rsPro("CSOMainHours"),2)
		    arrData(7, cnt) = 3
		    
		    cnt = cnt + 1
		    Redim preserve arrData(7, cnt)
		    arrData(0, cnt) = ""
		    arrData(1, cnt) = ""
		    arrData(2, cnt) = ""
		    arrData(3, cnt) = ""
		    if Mid(strproID,11,1)="L" Then	arrData(3, cnt) = "<b>CSO Man-Days</b>"
		    arrData(4, cnt) = ""
		    arrData(5, cnt) = ""
		    arrData(6, cnt)= ""
		    if not isnull(rsPro("CSOMainHours")) AND Mid(strproID,11,1)="L" then arrData(6, cnt) = FormatNumber(CDbl(rsPro("CSOMainHours"))/8,2)
		    arrData(7, cnt) = 3
		end if
Response.Write rsPro("CSOMainHours")
		if not IsNull(rsPro("CSOMainHours")) AND Mid(strproID,11,1)="L" then
			if cdbl(dblBillableHours) >cdbl(rsPro("CSOMainHours")) AND cdbl(rsPro("CSOMainHours"))>0 then
				cnt = cnt + 1
				Redim preserve arrData(7, cnt)
				arrData(0, cnt) = ""
				arrData(1, cnt) = ""
				arrData(2, cnt) = ""
				arrData(3, cnt) = "<b>Overrun:</b>"
				arrData(4, cnt) = ""
				arrData(5, cnt) = ""
				arrData(6, cnt) = FormatNumber(((dblBillableHours-cdbl(rsPro("CSOMainHours")))/cdbl(rsPro("CSOMainHours")))* 100,1) & "%"
				arrData(7, cnt) = 3
			end if
		end if
		session("arrSinglePro") = arrData
		session("NumPageSinglePro") = PageCount(arrData, PageSize)
		session("CurpageSinglePro") = 1			
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
	  else 'no data available
		session("NumPageSinglePro") = 0
		session("CurpageSinglePro") = 0
		session("arrSinglePro") = empty
	  end if 'test have data
	  session("READYSINGLEPRO") = true
  end if 'gMessage
else
	if isempty(session("NumPageSinglePro")) then  session("NumPageSinglePro") = 0
	if isempty(session("CurpageSinglePro")) then session("CurpageSinglePro") = 0
	if isempty(session("arrSinglePro")) then session("arrSinglePro") = empty
end if


	
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

if strproID = "" then strproID = Request.Form("txtproid")
varNavi = Request.QueryString("navi")
if varNavi <> "" then
	tmpi = session("CurPageSinglePro")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPageSinglePro") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPageSinglePro")
			end if
	End select
	session("CurPageSinglePro") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPageSinglePro") = CInt(varGo)

varSearch = ""
varSearch = Request.QueryString("search")
if varSearch<>""  then
	'making custom recordser
	if not IsEmpty(rsResult) then
		rsResult.Close
		set rsResult = nothing
		rsResult = Empty
		session("rsResultSINGLEPRO-L") = empty
	end if
	
	if not IsEmpty(rsPro) then	
		set rsResult = rsPro.Clone
		
		varSearch = replace(varSearch, "%", "")
		varSearch = replace(varSearch, "#", "")
		criteria = trim(varSearch)
		if criteria <> "" then
			if Instr(criteria, "'")>0 then
				criteria = "#" & criteria & "#"
			else
				criteria = "'%" & Replace(criteria, "'", "''") & "%'"
			end if
			rsResult.Filter = "ProjectID Like " & criteria
		else
			rsResult.MoveLast
			rsResult.MoveNext
		end if
		If rsResult.EOF then ' no result
			rsResult.Close
			set rsResult = nothing
			rsResult = Empty
			session("fgShowSINGLEPRO-L") = 0
			gMessage = "No results found."
			strproID = ""
			varSearch = ""
				
		else
			session("fgShowSINGLEPRO-L") = 1 ' show the result
			rsResult.MoveFirst
			set session("rsResultSINGLEPRO-L") = rsResult
			'reset
			strproID = ""
			session("NumPageSinglePro") = 0
			session("CurpageSinglePro") = 0
			session("arrSinglePro") = empty
		end if
	else
		gMessage = "No results found."
	end if
end if


if Request.QueryString("act")="vra3" then
	varFilter = "1"
else
	varFilter =  ""
end if

if varFilter<>"" then
	if isEmpty(session("filter")) then
		session("fgShowSINGLEPRO-L") = 0
	else
		if not isEmpty(rsResult) then
			rsResult.Close
			set rsResult = nothing
			rsResult = Empty
			session("rsResultSINGLEPRO-L") = empty
		end if
		strQuery = "SELECT DISTINCT ProjectID, ProjectName, fgActivate, ProjectKey2 FROM (" &_
			"SELECT a.ProjectID, a.ProjectName,a.CSOCompleted,a.CSOMainHours, a.fgActivate, b.ProjectKey2, c.ProjectTypeID " &_
			"ProjectType FROM ATC_Projects a " &_
			"LEFT JOIN (SELECT ProjectKey2, ProjectID, DateTransfer FROM ATC_ProjectStage) b On a.projectID = b.projectID " &_
			"AND b.DateTransfer = (SELECT min(k.DateTransfer) FROM (SELECT DateTransfer FROM ATC_ProjectStage WHERE " &_
			"ProjectID = a.ProjectID) k) " &_
			"LEFT JOIN ATC_projectPrjType c ON c.ProjectID = a.ProjectID WHERE a.fgDelete = 0 AND left(a.ProjectID, 1) <> '_') AA " &_
			"WHERE " & session("Filter")

		strConnect = Application("g_strConnect")
		Set objDb = New clsDatabase
		objDb.recConnect(strConnect)
		If objDb.openRec(strQuery) Then
			objDb.recDisConnect
			set rsResult = objDb.rsElement.Clone
			If not rsResult.EOF then
				session("fgShowSINGLEPRO-L") = 1 ' show the result
				rsResult.MoveFirst
				set session("rsResultSINGLEPRO-L") = rsResult
				'reset
				strproID = ""
				session("NumPageSinglePro") = 0
				session("CurpageSinglePro") = 0
				session("arrSinglePro") = empty
			else
				gMessage = "No results found."
				strproID = ""
				rsResult.Close
				set rsResult = nothing
				rsResult = Empty
				session("fgShowSINGLEPRO-L") = 0
			End if
				objDb.closeRec
		Else
		  gMessage = objDb.strMessage
		End if
		set objDb = nothing
	end if
end if
'Response.Write strproID
If session("fgShowSINGLEPRO-L") = 1 then 'after filter or searching
	strlstpro = selectpro(rsResult, "lstpro", strproID)
Else
	'= selectpro(rsPro, "lstpro", strproID)
	strlstpro  = "<select name='" & vname & "' size='1' height='26px' width='190px' " &_
			"style='width:190px;height=24px;' class='blue-normal'>"	
	strlstpro = strlstpro & "<option value='' selected>-- Select a project --</option>"
	strlstpro=strlstpro & "</select>"
	
end if
	
If gMessage="" then
	if not isEmpty(session("arrSinglePro")) then
	  arrGet = session("arrSinglePro")
	  strLast = OutBody(arrGet, PageSize, session("CurpageSinglePro"))
	  set arrGet = nothing
	end if
Else
	strLast = ""
End if
session("singleReport")=strLast
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
var objWindowSinglePro;
var objNewWindow;

function _print() { //v2.0
var str2 = "<%=strprintdate%>";
	str2 = escape(str2);
var fgprint = <%=session("NumPageSinglePro")%>;
if (fgprint!=0) {
	window.status = "";
	strFeatures = "top="+(screen.height/2-275)+",left="+(screen.width/2-390)+",width=800,height=550,toolbar=no," 
	            + "menubar=yes,location=no,directories=no,scrollbars=yes,status=yes";
	if ((objWindowSinglePro) && (!objWindowSinglePro.closed)) {
		objWindowSinglePro.focus();
	
	} else {
		objWindowSinglePro = window.open("p_singleproject.asp?printdate=" + str2, "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";
  }
else
	alert("No data for your request.")
}

function window_onunload() {
	if((objWindowSinglePro) && (!objWindowSinglePro.closed))
		objWindowSinglePro.close();
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.close();
}

function submitpro() {
	if(document.frmreport.lstpro.options[document.frmreport.lstpro.selectedIndex].value!="") {
		document.frmreport.action = "singleproject.asp?act=REFRESH";
		document.frmreport.target = "_self" ;
		document.frmreport.submit();
	}
	else {
		alert("Select a project, please.")
	}
}

function next() {
var curpage = <%=session("CurPageSinglePro")%>;
var numpage = <%=session("NumPageSinglePro")%>;
	if (curpage < numpage) {
		document.frmreport.action = "singleproject.asp?navi=NEXT"
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageSinglePro")%>;
var numpage = <%=session("NumPageSinglePro")%>;
	if (curpage > 1) {
		document.frmreport.action = "singleproject.asp?navi=PREV";
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageSinglePro")%>;
	var curpage = <%=session("CurPageSinglePro")%>;
	var intpage = document.frmreport.txtpage.value;
	intpage = parseInt(intpage, 10);
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmreport.action = "singleproject.asp?Go=" + intpage;
		document.frmreport.target = "_self";
		document.frmreport.submit();		
	}
}

function search() {
	var tmp = document.frmreport.txtsearch.value;
	tmp = escape(tmp);
	if (alltrim(tmp) != "") {
		document.frmreport.action = "singleproject.asp?search=" + tmp;
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function filter() { //v2.0
  window.status = "";
  strFeatures = "top="+(screen.height/2-105)+",left="+(screen.width/2-126)+",width=252,height=210,toolbar=no," 
              + "menubar=no,location=no,directories=no";
  if ((objNewWindow) && (!objNewWindow.closed)) {
	objNewWindow.focus();
	
  } else {
	objNewWindow = window.open("../management/project/n_profilter.asp", "MyNewWindow", strFeatures);
  }
  window.status = "Opened a new browser window.";
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
            <td width="36%"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td class="blue-normal" width="27%">&nbsp;Select Project</td>
                  <td class="blue-normal" width="73%"> 
<%Response.Write strlstpro%>
                  </td>
                </tr>
              </table>
            </td>
            <td width="11%"> 
              <table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
                <tr> 
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
                    <div align="center"> 
                      <p class="blue"><a href="javascript:submitpro();" onMouseOver="self.status='View report'; return true;" onMouseOut="self.status=''" class="b">Submit</a> 
                    </div>
                  </td>
                </tr>
              </table>
            </td>
            <td width="53%" align="right"> 
              <table border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td class="blue-normal" align="right" width="42%" valign="middle"> 
                    Search for&nbsp; </td>
                  <td align="right" width="18%" valign="middle"> 
                    <input type="text" name="txtsearch" class="blue-normal" size="15" style="width:150" value="<%=varsearch%>">
                  </td>
                  <td class="blue" align="right" width="21%" valign="middle"> 
                    <table width="80" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
                      <tr> 
                        <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
                          <div align="center"> 
                            <p class="blue"><a href="javascript:search();" onMouseOver="self.status='Search for ProjectID'; return true;" onMouseOut="self.status=''" class="b">Search</a> 
                          </div>
                        </td>
                        
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr align="center"> </tr>
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
			<td>&nbsp;</td>
          </tr>
        </table>
    		<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary by Project")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", "")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
			%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                <tr> 
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5">ProjectID </td>
                  <td class="blue" align="center" width="20%" bgcolor="#E7EBF5">Project Name </td>
                  <td class="blue" align="center" width="17%" bgcolor="#E7EBF5">Full Name </td>
                  <td class="blue" align="center" width="22%" bgcolor="#E7EBF5">Sub-task Description </td>
                  <td class="blue" align="center" width="9%" bgcolor="#E7EBF5">Hours</td>
                  <td class="blue" align="center" width="8%" bgcolor="#E7EBF5">Overtime</td>
                  <td class="blue" align="center" width="9%" bgcolor="#E7EBF5">Total<br> Hours </td>
                </tr>
<%Response.Write strLast%>
				
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<%if session("NumPageSinglePro")>0 then%>
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="20" align="center">
    <tr> 
      <td align="right" bgcolor="#E7EBF5"> 
        <table width="70%" border="0" cellspacing="1" cellpadding="0" height="18">
          <tr> 
            <td align="right" valign="middle" width="37%" class="blue-normal">Page 
            </td>
            <td align="center" valign="middle" width="13%" class="blue-normal"> 
              <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageSinglePro")%>" size="2" style="width:50">
            </td>
            <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();"  onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
            </td>
            <td align="right" valign="middle" width="15%" class="blue-normal">Pages 
               <%=session("CurpageSinglePro")%>/<%=session("NumpageSinglePro")%>&nbsp;&nbsp;</td>
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
<input type="hidden" name="txtproid" value="<%=strproID%>">
</form>
</body>
</html>