<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->
<!-- #include file = "../../../inc/library.asp"-->

<%
	Dim strUserName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError, intPageSize, fgRight 'view all or Not
	Dim varEmp, varFullName

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
	If Not rsSrc.EOF Then
		For i = 1 To psize
			If i Mod 2 = 0 Then
				strColor = "#E7EBF5"
			Else
				strColor = "#FFF2F2"
			End If
			
			If Not IsEmpty(session("varEmp")) Then
				varEmp = session("varEmp")

				For kk = 0 To Ubound(varEmp,2)
					If trim(rsSrc("StaffID")) = trim(varEmp(0,kk)) Then
						If CInt(varEmp(5,kk)) = 1 Then
							strCheck = "checked"
							Exit For
						Else
							strCheck = ""	
						End If	
					End If	
				Next
			Else
				strCheck = ""	
			End If
								
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='middle' width='35%' class='blue'><a href='javascript:viewsal(" & rsSrc("StaffID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;Generate Salary&quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("Fullname")) & "</a></td>" &_
			         "<td valign='middle' width='30%' class='blue-normal'>" & Showlabel(rsSrc("JobTitle")) & "</td>" &_
			         "<td valign='middle' width='30%' class='blue-normal'>" & Showlabel(rsSrc("Department")) & "</td>" &_
			         "<td valign='middle' width='5%' class='blue-normal' align='center'>" & _
					 "<input type='checkbox' name='chkget' value='" & rsSrc("StaffID") & ";" & rsSrc.Bookmark & "'" & strcheck & "></td>" & _
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	End If
	Outbody = strOut
End Function

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intDepartmentID = Request.Form("lbdepartment")
	fgSort = Request.Form("S")
	fgCheckbox = 0
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
		
	intCurPage = trim(Request.Form("P"))
	If intCurPage = "" Then
		intCurPage = 1
	End If		
	strName = Request.Form("name")
	intDepart = Request.Form("depart")
	Redim varEmp(4,-1)
	
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../../message.htm")
	End If					

	intUserID = session("USERID")
	
'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = 8'PageSizeDefault
		Set arrPre = Nothing
	Else
		intPageSize = 8'PageSizeDefault
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
		Response.Redirect("../../../welcome.asp")
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
		
		If objDatabase.dbConnect(strConnect) Then

'--------------------------------------------------
' Check right on page
'--------------------------------------------------
			
			strSQL = "exec sp_CheckRight " & intUserID
			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord Then				 
					strError1 = "You have no access rights."
				End If
				objDatabase.closeRec
			Else
				strError = "1. " & objDatabase.strMessage
			End If

'--------------------------------------------------
' End of checking right on page
'--------------------------------------------------

			Set rsStaff = Server.CreateObject("ADODB.Recordset")
			rsStaff.CursorLocation = adUseClient										' Set the Cursor Location to Client

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

			strSQL = "SELECT a.StaffID, FirstName + ' ' + ISNULL(MiddleName,'') + ' '+ ISNULL(LastName,'') AS FullName, a.DepartmentID, ISNULL(c.JobTitle,'') AS JobTitle, ISNULL(d.Department,'') AS Department, 0 AS fgExist" & _
						" FROM ATC_Employees a LEFT JOIN ATC_PersonalInfo b ON a.StaffID = b.PersonID LEFT JOIN ATC_JobTitle c " & _
						" ON a.JobTitleID = c.JobTitleID LEFT JOIN ATC_Department d ON a.DepartmentID = d.DepartmentID WHERE b.fgDelete = 0" 'AND LeaveDate IS NULL"
						
			If fgRight Then						' View all
				myCmd("fgCheck") = 0
			Else
				strSQL = strSQL & " AND a.StaffID "
				myCmd("fgCheck") = 1 
			End If
			myCmd("strSQL") = strSQL

'			On Error Resume Next
			
			rsStaff.Open myCmd,,adOpenStatic,adLockBatchOptimistic
				
			If Err.number > 0 then
				strError = Err.Description
			End If
			Err.Clear
			If Not rsStaff.EOF Or rsStaff.RecordCount > 0 Then
				intTotalPage = pageCount(rsStaff, intPageSize)
				rsStaff.MoveFirst

				'rsStaff.Sort="FullName ASC"

				varEmp = rsStaff.GetRows
				session("varEmp") = varEmp
				
'Response.Write varEmp(0,1) & "--" & varEmp(1,1) & "--" & varEmp(2,1) & "--" & varEmp(3,1) & "--" & varEmp(4,1) & "--" & varEmp(5,1)
'Response.Write rsStaff.Fields(0).Name & "--" & rsStaff.Fields(1).Name & "--" &rsStaff.Fields(2).Name & "--" & rsStaff.Fields(3).Name & "--" & rsStaff.Fields(4).Name & "--" & rsStaff.Fields(5).Name

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
	
		intMonth = Request.Form("lbmonth")
		intYear = Request.Form("lbyear")
		
		Set rsStaff = session("rsStaff")

		rsStaff.MoveFirst
		If recCount(rsStaff) > 0 Then
			intTotalPage = pageCount(rsStaff, intPageSize)
		
			Select Case strAct
				Case "vpr"
'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					intSelect = Request.Form("chkget").Count
					
					for ii=1 to intSelect
						varTemp=split(Request.Form("chkget")(ii),";")
						intPosition = CInt(varTemp(1)) - 1
						varEmp(5,intPosition) = 1
					next
										
					'For ii = 1 To intPageSize
					'	chkget = "chkget" & CStr(ii)
					'	vcheck = Trim(Request.Form(chkget))
					'	If vcheck <> "" Then
							'Response.Write chkget & ":" & Request.Form(chkget) & "<br>"
					'		varTemp = split(Request.Form(chkget),";")
					'		intPosition = CInt(varTemp(1)) - 1
					'		varEmp(5,intPosition) = 1
					'	Else
					'		intPosition = rsStaff.Bookmark - 1
					'		varEmp(5,intPosition) = 0
					'	End If					
					'	rsStaff.MoveNext
					'	If rsStaff.EOF Then Exit For
					'Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
'--------------------------------------------------
				
					varEmp = session("varEmp")
					If IsArray(varEmp) Then
						intEmp = Ubound(varEmp,2)					
					End If
					
					If intEmp >= 0 Then	
						For ii = 0 To intEmp
							If varEmp(5,ii) = 1 Then
								fgCheckbox = 1
								Exit For
							End If	
						Next
					End If
					
					If fgCheckbox = 0 Then
						strError = "Please choose checkbox before click button 'Print'" 
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
	
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

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------
					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					For ii = 1 To intPageSize
						chkget = "chkget" & CStr(ii)
						vcheck = Trim(Request.Form(chkget))
						If vcheck <> "" Then
							varTemp = split(Request.Form(chkget),";")
							intPosition = CInt(varTemp(1)) - 1
							varEmp(5,intPosition) = 1
						Else
							intPosition = rsStaff.Bookmark - 1
							varEmp(5,intPosition) = 0
						End If					
						rsStaff.MoveNext
						If rsStaff.EOF Then Exit For
					Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
'--------------------------------------------------

					If fgSort = "D" Then
						rsStaff.Sort = "FullName ASC"
						fgSort = "A"
					ElseIf fgSort = "N" Or fgSort = "A"	Then
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

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------
					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					For ii = 1 To intPageSize
						chkget = "chkget" & CStr(ii)
						vcheck = Trim(Request.Form(chkget))
						If vcheck <> "" Then
							varTemp = split(Request.Form(chkget),";")
							intPosition = CInt(varTemp(1)) - 1
							varEmp(5,intPosition) = 1
						Else
							intPosition = rsStaff.Bookmark - 1
							varEmp(5,intPosition) = 0
						End If					
						rsStaff.MoveNext
						If rsStaff.EOF Then Exit For
					Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
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
					
				Case "vpsd"											' Sort by department		

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

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------
					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					For ii = 1 To intPageSize
						chkget = "chkget" & CStr(ii)
						vcheck = Trim(Request.Form(chkget))
						If vcheck <> "" Then
							varTemp = split(Request.Form(chkget),";")
							intPosition = CInt(varTemp(1)) - 1
							varEmp(5,intPosition) = 1
						Else
							intPosition = rsStaff.Bookmark - 1
							varEmp(5,intPosition) = 0
						End If					
						rsStaff.MoveNext
						If rsStaff.EOF Then Exit For
					Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
'--------------------------------------------------

					If fgSort = "N" Or fgSort = "D" Then
						rsStaff.Sort = "Department ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "Department DESC"
						fgSort = "D"				
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
				Case "vpa1"											' When user click button "Go"

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------
					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					For ii = 1 To intPageSize
						chkget = "chkget" & CStr(ii)
						vcheck = Trim(Request.Form(chkget))
						If vcheck <> "" Then
							varTemp = split(Request.Form(chkget),";")
							intPosition = CInt(varTemp(1)) - 1
							varEmp(5,intPosition) = 1
						Else
							intPosition = rsStaff.Bookmark - 1
							varEmp(5,intPosition) = 0
						End If					
						rsStaff.MoveNext
						If rsStaff.EOF Then Exit For
					Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
'--------------------------------------------------				
				
					If CInt(Request.Form("txtpage")) <= CInt(intTotalPage) Then
						intCurPage = Request.Form("txtpage")
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
					
				Case "vpa2"											' When user click Previous link	
				
'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------
					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					For ii = 1 To intPageSize
					
						chkget = "chkget" & CStr(ii)
						vcheck = Trim(Request.Form(chkget))
						If vcheck <> "" Then
							varTemp = split(Request.Form(chkget),";")
							intPosition = CInt(varTemp(1)) - 1
							varEmp(5,intPosition) = 1
						Else
							intPosition = rsStaff.Bookmark - 1
							varEmp(5,intPosition) = 0
						End If					
						rsStaff.MoveNext
						If rsStaff.EOF Then Exit For
					Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
'--------------------------------------------------				
				
					If CInt(intCurPage) > 1 Then
						intCurPage = CInt(intCurPage) - 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")

				Case "vpa3"											' When user click Next link		

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------
					
					varEmp = session("varEmp")

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					For ii = 1 To intPageSize
						chkget = "chkget" & CStr(ii)
						vcheck = Trim(Request.Form(chkget))
						If vcheck <> "" Then
							varTemp = split(Request.Form(chkget),";")
							intPosition = CInt(varTemp(1)) - 1
							varEmp(5,intPosition) = 1
						Else
							intPosition = rsStaff.Bookmark - 1
							varEmp(5,intPosition) = 0
						End If					
						rsStaff.MoveNext
						If rsStaff.EOF Then Exit For
					Next
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End 
'--------------------------------------------------
					
					If CInt(intCurPage) < CInt(intTotalPage) Then
						intCurPage = CInt(intCurPage) + 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
					
				Case "vra1"											' When user click button "Search"
				
'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------

					varEmp = session("varEmp")
					If IsArray(varEmp) Then
						intEmp = Ubound(varEmp,2)					
					End If
					
					If intEmp >= 0 Then	
						For ii = 0 To intEmp
							varEmp(5,ii) = 0
						Next
					End If
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End
'--------------------------------------------------
				
					strSName = Request.Form("txtname")
					intDepart = Request.Form("lbdepartment")

					If strSName <> "" And CInt(intDepart) <> 0 Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%' AND DepartmentID=" & intDepart
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "# AND DepartmentID=" & intDepart
						End If		
					ElseIf strSName = "" And CInt(intDepart) <> 0 Then
						rsStaff.Filter = "DepartmentID=" & intDepart
					ElseIf strSName <> "" And CInt(intDepart) = 0 Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%'"
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "#"
						End If	
					End If

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
				
'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------

					varEmp = session("varEmp")
					If IsArray(varEmp) Then
						intEmp = Ubound(varEmp,2)					
					End If
					
					If intEmp >= 0 Then	
						For ii = 0 To intEmp
							varEmp(5,ii) = 0
						Next
					End If
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End
'--------------------------------------------------

					rsStaff.Filter = ""
					intTotalPage = pageCount(rsStaff, intPageSize)

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)

					intDepartmentID = 0
					strName = ""

					strStatus = Request.Form("txtstatus")
					
				Case "vca1"											' When user click button "Select all"	

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------

					varEmp = session("varEmp")
					If IsArray(varEmp) Then
						intEmp = Ubound(varEmp,2)					
					End If
					
					If intEmp >= 0 Then	
						For kk = 0 To intEmp
							rsStaff.MoveFirst
							Do While Not rsStaff.EOF
								If trim(rsStaff("StaffID")) = trim(varEmp(0,kk)) Then
									varEmp(5,kk) = 1
									Exit Do
								End If	
								rsStaff.MoveNext
							Loop	
						Next
					End If

					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End
'--------------------------------------------------

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					strStatus = Request.Form("txtstatus")				
					
				Case "vca2"											' When user click button "Clear all"

'--------------------------------------------------
' Get users that was checked
'--------------------------------------------------

					varEmp = session("varEmp")
					If IsArray(varEmp) Then
						intEmp = Ubound(varEmp,2)					
					End If
					
					If intEmp >= 0 Then	
						For kk = 0 To intEmp
							rsStaff.MoveFirst
							Do While Not rsStaff.EOF
								If trim(rsStaff("StaffID")) = trim(varEmp(0,kk)) Then
									varEmp(5,kk) = 0
									Exit Do
								End If	
								rsStaff.MoveNext
							Loop	
						Next
					End If
				
					session("varEmp") = varEmp	
					Set varEmp = Nothing

'--------------------------------------------------
' End
'--------------------------------------------------

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					strStatus = Request.Form("txtstatus")				
					
			End Select	 
			
			strLast = Outbody(rsStaff, intPageSize)
		End If		
	End If
		
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
				"<img src='../../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu = "" then strChoseMenu = "AD"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../../timesheet.css">
<script language="javascript" src="../../../library/library.js"></script>

<script language="javascript">
<!--

function sort(kind)
{
	document.frmreport.action = "sal_list_staff.asp?act=vps" + kind;
	document.frmreport.submit();
}

function viewpage(kind)
{
	var intpage = parseInt(window.document.frmreport.txtpage.value,10);
	var curpage = "<%=CInt(intCurPage)%>";
	var pagetotal = "<%=CInt(intTotalPage)%>";
	
	if (kind == 1)
	{
		window.document.frmreport.txtpage.value = intpage
		if ((intpage > 0) & (intpage <= pagetotal) & (intpage != curpage)) 
		{
			document.frmreport.action = "sal_list_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "sal_list_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function search(kind)
{
	if (kind == "1")
	{
		if ((document.frmreport.txtname.value != "") || (document.frmreport.lbdepartment.options[document.frmreport.lbdepartment.selectedIndex].value != "0"))
		{
			document.frmreport.action = "sal_list_staff.asp?act=vra" + kind;
			document.frmreport.submit();	
		}
	}	
	else
	{
		document.frmreport.action = "sal_list_staff.asp?act=vra" + kind;
		document.frmreport.submit();
	}	
}

//function setchecked(val) 
//{
//	if (val == 1)
//		window.document.frmreport.action = "sal_list_staff.asp?act=vca1"
//	else
//		window.document.frmreport.action = "sal_list_staff.asp?act=vca2"
//	
//	window.document.frmreport.submit();		
//}

function setchecked(val) 
{
	with (document.frmreport) 
	{
		len = elements.length;
		for(var ii=0; ii<len; ii++) 
			if (elements[ii].name == "chkget") 
				elements[ii].checked = val
	}
}

function viewsal(varid)
{
	window.document.frmreport.txthidden.value = varid;
//	window.document.frmreport.action = "sal_staff_tms.asp";
//	window.document.frmreport.action = "sal_type_staff.asp";
	window.document.frmreport.action = "sal_staff_tms.asp";
	window.document.frmreport.submit();
}

function printpage()
{
	window.document.frmreport.M.value = window.document.frmreport.lbmonth.options[window.document.frmreport.lbmonth.selectedIndex].value;
	window.document.frmreport.Y.value = window.document.frmreport.lbyear.options[window.document.frmreport.lbyear.selectedIndex].value

	window.document.frmreport.action = "sal_list_staff.asp?act=vpr";
	window.document.frmreport.submit();
}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<%	If strError1 = "" Then%>        
		  <tr> 
		    <td> 
		      <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%		If strError <> "" Then%>		      
		        <tr bgcolor="#E7EBF5">
				  <td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
		        </tr>
<%		End If%>		        
		        <tr align="center"> 
		          <td class="blue-normal" height="30" align="right" width="40">&nbsp;&nbsp;Name&nbsp;</td>
  				  <td class="blue" height="30" align="left" width="169"> 
				    <input type="text" name="txtname" value="<%=showvalue(strSName)%>" class="blue-normal" size="15" style=" width:150">
				  </td>
				  <td class="blue-normal" height="30" align="right" width="65">Department&nbsp;</td>
				  <td class="blue" height="30" align="left" width="79"> 
				    <select id="lbdepartment" size="1" name="lbdepartment" class="blue-normal">
					  <option value="0" selected>&nbsp;</option>
<%
		If intNum >= 0 Then
			For ii = 0 To intNum
%>                    
					  <option <%If CInt(intDepartmentID)=CInt(varDepartment(0,ii)) Then%> selected <%End If%> value="<%=varDepartment(0,ii)%>"><%=showlabel(varDepartment(1,ii))%></option>
<%
			Next
		End If	
%>					

				    </select>
				  </td>
				  <td class="blue-normal" height="30" align="left" width="255"> 
				    <table width="120" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
					  <tr> 
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
						  <div align="center" class="blue"><a href="javascript:search('1');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Search</a></div>
					    </td>
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
						  <a href="javascript:search('2');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Show All</a>
					    </td>
					  </tr>
				    </table>
				  </td>
			    </tr>
		        <tr> 
				  <td bgcolor="#8CA0D1" colspan="5"><img src="../../../IMAGES/DOT-01.GIF" width="1" height="1"></td>
				</tr>
				<tr align="center"> 
				  <td class="title" height="40" align="center" colspan="5">Generate Salary Sheet</td>
			    </tr>
			    <tr bgcolor="#FFFFFF" height="20" align="center">
				  <td colspan="5" height="30" valign="top">
		             <table width="30%" border="0" cellspacing="0" cellpadding="2" >
		               <tr>
		                 <td valign="bottom" width="10%" class="blue">&nbsp;&nbsp;</td>
			             <td valign="middle" class="blue-normal" width="15%">Month&nbsp;</td>
			             <td valign="bottom" class="blue-normal" width="30%"> 
				           <select name="lbmonth" size="1" class="blue-normal">
						<%For ii=1 To 12%>
								<option <%If CInt(intMonth)=ii Then%>selected<%End If%> value="<%=ii%>"><%=SayMonth(ii)%></option>
						<%Next%>
						   </select>
			             </td>
			             <td valign="middle" width="15%" class="blue-normal">Year&nbsp;</td>
			             <td valign="bottom" class="blue-normal" width="30%"> 
					       <select name="lbyear" size="1" class="blue-normal">
						<%For ii=Year(Date)-2 To Year(Date)%>
		 	        	     <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
						<%Next%>
			 		       </select>
			             </td>
			           </tr>
		            </table>
				  </td>
				</tr>  				
			  </table>
		    </td>
		  </tr>
		  <tr> 
		    <td height="100%"> 
		      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:'79%'" height="365">
		        <tr> 
		          <td bgcolor="#FFFFFF" valign="top"> 
		            <table width="100%" border="0" cellspacing="0" cellpadding="0">
		              <tr> 
		                <td bgcolor="#617DC0"> 
		                  <table width="100%" border="0" cellspacing="1" cellpadding="5">
		                    <tr bgcolor="#8CA0D1"> 
		                      <td class="blue" align="center" width="35%"><a href="javascript:sort('n')" onMouseOver="self.status='Sort by Full Name';return true" onMouseOut="self.status='';return true" class="c">Full Name</a></td>
		                      <td class="blue" align="center" width="30%"><a href="javascript:sort('t')" onMouseOver="self.status='Sort by Job Title';return true" onMouseOut="self.status='';return true" class="c">Job Title</a> </td>
		                      <td class="blue" align="center" width="30%"><a href="javascript:sort('d')" onMouseOver="self.status='Sort by Department';return true" onMouseOut="self.status='';return true" class="c">Department</a></td>
                              <td valign="top" width="5%" class="blue-normal" align="center">&nbsp;</td> 
		                    </tr>
<%
	Response.Write(strLast)
%>                            
		                  </table>
		                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;&nbsp;* Click 
                                on the exact name to view salary sheet.</td>
                              <td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">
		                        <a href="javascript:setchecked(1);" onMouseOver="self.status='';return true">Select All</a>&nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='';return true">Clear All</a> &nbsp;&nbsp;&nbsp;</td>
                            </tr>
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal" colspan="2">&nbsp;&nbsp;* Choose month, year before click 
                                on Print button to print salary sheet(s) of chose staff. 
                              </td>
                            </tr>
                            <tr> 
                              <td bgcolor="#FFFFFF" class="blue-normal" colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td bgcolor="#FFFFFF" height="40" class="blue-normal" colspan="2"> 
                                <table width="60" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
<!--                                  
                                    <td width="60" bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
                                      <a href="#" class="b">Export</a>
                                    </td>
-->                                    
                                    <td width="60" bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
                                      <a href="javascript:printpage()" onMouseOver="self.status='';return true"  onMouseOut="self.status='';return true" class="b">Print</a>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
		                  </table>
		                </td>
		              </tr>
		            </table>
		          </td>
		        </tr>
		      </table>
		    </td>
		  </tr>
          <tr> 
            <td> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
                <tr> 
                  <td align="right" bgcolor="#E7EBF5"> 
                    <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
                      <tr class="black-normal"> 
                        <td align="right" valign="middle" width="27%" class="blue-normal">Page</td>
                        <td align="center" valign="middle" width="13%" class="blue-normal"> 
                          <input type="text" name="txtpage" class="blue-normal" value="<%=intCurPage%>" size="2" style="width:50">
                        </td>
                        <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:viewpage(1);" onMouseOver="self.status='';return true"><font color="#990000">Go</font></a></td>
						<td align="right" valign="middle" width="25%" class="blue-normal"><%If CInt(intTotalPage) <> 0 Or intTotalPage <> "" Then%>Pages <%=intCurPage%>/<%=intTotalPage%><%End If%>&nbsp;&nbsp;</td>
						<td valign="middle" align="right" width="28%" class="blue-normal"><%If CInt(intCurPage) <> 1 Then%><a href="javascript:viewpage(2);" onMouseOver="self.status='Move Previous';return true" onMouseOut="self.status='';return true">Previous</a><%End If%><%If CInt(intCurPage) <> 1 And  CInt(intCurPage) <> CInt(intTotalPage) Then%>/<%End If%><%If CInt(intCurPage) <> CInt(intTotalPage) And (CInt(intTotalPage) <> 0 Or intTotalPage <> "") Then%><a href="javascript:viewpage(3);" onMouseOver="self.status='Move Next';return true" onMouseOut="self.status='';return true"> Next</a><%End If%>&nbsp;&nbsp;&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
<%	Else          
		If strError <> "" Then
%>               
				<tr>
				  <td class="red">&nbsp;<%=strError%></td>
				</tr>
<%		End If%>				

		  <tr>
         	<td class="red" align="center" valign="middle"><b><%=strError1%></b></td>
		  </tr>	          

<%	End If
	Set objDatabase = Nothing
%>
        </table>
      
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
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="txtstatus" value="<%=strStatus%>">
<input type="hidden" name="P" value="<%=intCurPage%>">
<input type="hidden" name="S" value="<%=fgSort%>">
<input type="hidden" name="name" value="<%=strSName%>">
<input type="hidden" name="depart" value="<%=intDepart%>">
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">

</form>
<%
	If Request.QueryString("act") = "vpr" Then
		If fgCheckbox = 1 Then
%>
<SCRIPT language="javascript">
var objNewWindow;

	window.status = "";
	
	strFeatures = "top=20,left=70,width=800,height=600,toolbar=no," 
              + "menubar=yes,location=no,directories=no,resizable=yes,scrollbars=yes";

	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('sal_print_all.asp?m=' + document.frmreport.lbmonth.value + '&y=' + document.frmreport.lbyear.value, "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  

</SCRIPT>
<%
		End If
	End If%>

</body>
</html>