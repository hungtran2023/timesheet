<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' function: appendTree
' Description: draw tree of subtask
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function AppendTree (ByVal strsName, ByVal intLevel, ByVal blnShow, ByVal intValue)
Dim strTmp, i, strColor
	strTmp = ""
	If intLevel > 0 Then		
		For i = 1 to intLevel
			strTmp = strTmp & "<IMG alt='' border='0' height='18' src='../../images/t_dot.gif' width='36'>"
		Next
		strTmp = strTmp & "<IMG alt='' border='0' src='../../images/dot1.gif'>"
		strTmp = strTmp & "<IMG alt='' border='0' height='10' width='12' src='../../images/nosign.gif'>"
	End If
	LineOnPage = LineOnPage + 1
	strColor = "#FFF2F2"
	if blnShow = true and intLevel>0 then
		AppendTree = "<tr bgcolor='" & strColor & "'><td valign='top' class='blue'>" & strTmp & Showlabel(strsName) & "</td>" &_
					"<td valign='top' width='5%' align='center' class='blue'><input type='checkbox' name='chkrem' value='" & intValue & "@" & showlabel(strsName) & "'>" &_
		            "</td></tr>" & chr(13)
	else
		AppendTree = "<tr bgcolor='" & strColor & "'><td valign='top' class='black'>" & strTmp & Showlabel(strsName) & "</td>" &_
					"<td valign='top' width='5%' align='center' class='black'>&nbsp;" &_
		            "</td></tr>" & chr(13)
	end if
End Function
'****************************************
' function: FetchChild
' Description: this is a recursive function.
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub FetchChild(ByRef rsGet, ByRef strTree, ByVal intLevel, ByRef rsAll, ByVal ParentOwner)
Dim strName, intContinue
	Do Until rsGet.EOF
		blnOwner = false
		If rsGet("Owner")<>0 or ParentOwner=true then
			blnOwner = true
		end if
		ParOwner = blnOwner or (rsGet("Owner")<>0)
		
	    strTree = strTree & AppendTree(rsGet("sName"), intLevel, blnOwner, rsGet("sID"))
		rsAll.Filter = "sParentID = " & rsGet("sID")
		intContinue = 0
		If rsAll.RecordCount > 0 then
		  intContinue = 1
		  Call CopyData(rsAll, arrRs(intLevel + 1))
		  arrRs(intLevel + 1).MoveFirst
		End If
		rsAll.Filter = ""
		If intContinue = 1 Then
		  FetchChild arrRs(intLevel + 1), strTree, intLevel + 1, rsAll, ParOwner
		End If
		rsGet.MoveNext
	Loop
End Sub
'****************************************
' function: makeSelect
' Description: generate a listbox of project type
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function makeSelect(ByRef arrItems, ByVal varEdit, ByVal strSelected)
	strOut = "<select name='lsttype' MULTIPLE class='blue-normal' size='3' "
	if not varEdit then
		strOut = strOut & "onChange='restorelist(this, &quot;" & strarrType & "&quot;);'>"
	else
		strOut = strOut & ">"
	end if
	if isArray(arrItems) then
		For i = 0 to Ubound(arrItems, 2)
			if InStr(strSelected, "@" & CStr(arrItems(0, i)) & "@")>0 then
				strOut = strOut & "<option value='" & arrItems(0, i) & "' selected>" & showlabel(arrItems(1, i)) & "</option>"
		    else
				strOut = strOut & "<option value='" & arrItems(0, i) & "'>" & showlabel(arrItems(1, i)) & "</option>"
		    end if
		Next
	end if
	strOut = strOut & "</select>"
	makeSelect = strOut
end function
'****************************************
' function: task_delete
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_delete
	strMode=Request.Form("txtmode")
	strProID = Request.Form("txtproID")
	proID = strProID
	strProName = Request.Form("txtproName")
	varInhouse = Request.Form("chkinhouse")
	varActivate = Request.Form("chkactivate")
	if varInhouse = "" then varInhouse = 0
	if varActivate = "" then varActivate = 0
	strOut = makeSelect(session("typeofproject"), true, session("selected"))
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	fgdel=false
	if ret then
		set myCmd = Server.CreateObject("ADODB.Command")
		set myCmd.ActiveConnection = objDb.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "sp_checkproDel"
		set myParam = myCmd.CreateParameter("result",adTinyInt,adParamReturnValue)
		myCmd.Parameters.Append myParam
		set myParam = myCmd.CreateParameter("proID",adVarChar,adParamInput, 20)
		myCmd.Parameters.Append myParam
		myCmd("proID") = proID
		myCmd.Execute , , adExecuteNoRecords
		if myCmd("result") = 1 then
			strQuery = "DELETE FROM ATC_Projects WHERE ProjectID = '" & proID & "'"
			ret = objDb.runActionQuery(strQuery)
			if not ret then 
				gMessage = objDb.strMessage
			else
				fgdel=true			
			end if			
		else
			strQuery = "UPDATE ATC_Projects SET fgDelete = 1 WHERE ProjectID = '" & proID & "'"
			ret = objDb.runActionQuery(strQuery)
			if not ret then 
				gMessage = objDb.strMessage
			else
				fgdel=true			
			end if
		end if
		set myCmd = nothing
		objDb.dbdisConnect
	else
		gMessage = objDb.strMessage		
	end if	
	set objDb = nothing
	if fgdel then Response.Redirect("listofproject.asp?act=REFRESH")
End Sub
'****************************************
' function: task_remove
' Description: removing assignment
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_remove
	countU = Request.Form("chkrem").Count
	if countU>0 then
	  Set objDb = New clsDatabase
	  strConnect = Application("g_strConnect")
	  ret = objDb.dbConnect(strConnect)
	  if ret then
	    strDonot = ""
	    for ii = 1 to countU
	  		strTaskname = Request.Form("chkrem")(ii)
	  		strTask = Mid(strTaskname, 1, Instr(strTaskname, "@") - 1)
	  		strTaskname = Mid(strTaskname, Instr(strTaskname, "@") + 1, len(strTaskname))			
		
			set myCmd = Server.CreateObject("ADODB.Command")
			set myCmd.ActiveConnection = objDb.cnDatabase
			myCmd.CommandType = adCmdStoredProc
			myCmd.CommandText = "sp_checkanddeltask"
			set myParam = myCmd.CreateParameter("result", adTinyInt, adParamReturnValue)
			myCmd.Parameters.Append myParam
			set myParam = myCmd.CreateParameter("taskID", adInteger, adParamInput)
			myCmd.Parameters.Append myParam
			myCmd("taskID") = CInt(strTask)
			myCmd.Execute , , adExecuteNoRecords
			if myCmd("result") = 0 then
				strDonot = strDonot & " " & strTaskname & ","
			end if
			set myCmd = nothing
	    next
	    if strDonot<>"" then
	      strDonot = Mid(strDonot, 1, len(strDonot)- 1)
	      gMessage = "Can not remove '" & strDonot & "'."
	      fgRefresh = "0"
	    else
	  	  gMessage = "Removed successfully."
	  	  fgRefresh = "1"
	    end if
	    objDb.dbDisConnect
	  else
	    gMessage =  objDb.strMessage
	  end if
	  Set objDb = Nothing
	end if
End Sub
'****************************************
' function: task_add
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function task_add
Dim fgChange
  Set objDb = New clsDatabase
  strConnect = Application("g_strConnect")
  ret = objDb.dbConnect(strConnect)
  if ret then
'---------------------------
' Check duplicate project id
'---------------------------	
	strQuery = "Select count(*) as mysum From ATC_Projects WHERE projectID = '" & strProIDi & "'"
	ret = objDb.runQuery(strQuery)
	if objDb.rsElement("mysum") > 0 then
		gMessage = "This ProjectID has already been inputted."
	end if
	if gMessage="" then '---starting
	  objDb.cnDatabase.BeginTrans
	  strQuery = "INSERT INTO ATC_Projects(ProjectID, ProjectName, fgActivate, CompanyID) " &_
	  		"VALUES('" & strProIDi & "', '" & strProNamei & "', " & varActivate & ", " & varCompany & ")"
	  ret = objDb.runActionQuery(strQuery)
	  if ret then
	  '------insert table atc_tasks
	    strQuery = "INSERT INTO ATC_Tasks(ProjectID, SubTaskName, ownerID) VALUES('" & strProIDi &_
	  			"', '" & strProIDi & " _ " & strProNamei & "', " & session("USERID") & ")"
	    ret = objDb.runActionQuery(strQuery)
	    if not ret then 
		  gMessage = objDb.strMessage
		else
		'------insert table ATC_ProjectPrjType
		  for ii = 1 to countU
			  strQuery = "INSERT INTO ATC_ProjectPrjType(ProjectID, ProjectTypeID) VALUES('" & strProIDi &_
						"', " & int(Request.Form("lsttype")(ii)) & ")"
			  ret = objDb.runActionQuery(strQuery)
			  if not ret then 
				gMessage = objDb.strMessage
				exit for
			  end if
		  next
		end if
	  else
	    gMessage = objDb.strMessage
	  end if
	  if gMessage<>"" then 
	  	objDb.cnDatabase.RollbackTrans
	  	fgChange = false
	  else
	  	objDb.cnDatabase.CommitTrans
	  	gMessage = "Added successfully."
	  	fgRightonPro = true
	  	fgChange = true
	  	proID = strProIDi
	  end if
	end if'-----------ending
	objDb.dbdisConnect
  else
    gMessage = objDb.strMessage
  end if  
  set objDb = nothing
  task_add = fgChange
End Function
'****************************************
' function: task_update
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function task_update
Dim fgChange
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then '---------------connection	  
	  if proID=strProIDi then 'normal			  
	    strQuery = "UPDATE ATC_Projects SET ProjectName = '" & strProNamei & "', " &_
	  				"CompanyID = " & varCompany & ", fgActivate = " & varActivate &_
	  				"WHERE ProjectID = '" & proID & "'"
	  else ' special, ie ProjectID is changed	
		strQuery = "Select count(*) as mysum From ATC_Projects WHERE projectID = '" & strProIDi & "'"
		ret = objDb.runQuery(strQuery)
		if objDb.rsElement("mysum") > 0 then
		  gMessage = "This ProjectID has already been inputted."
		  strQuery = ""
		else
		  strQuery = "UPDATE ATC_Projects SET ProjectID = '" & strProIDi & "', ProjectName = '" & strProNamei & "', " &_
	  					"CompanyID = " & varCompany & ", fgActivate = " & varActivate &_
	  					"WHERE ProjectID = '" & proID & "'"
	  	end if
	  end if
	  if strQuery<>"" then '-----------starting
	    objDb.cnDatabase.BeginTrans
		ret = objDb.runActionQuery(strQuery)
	    if not ret then 
			gMessage = objDb.strMessage
	    else
	  		if proID<>strProIDi then
	  		  ''update in table ATC_Tasks
	  		  strQuery = "UPDATE ATC_Tasks SET ProjectID = '" & strProIDi & "' WHERE ProjectID = '" & proID & "' AND TaskID is not null"
	  		  strQuery = strQuery & " UPDATE ATC_Tasks SET ProjectID = '" & strProIDi & "', SubTaskName ='" & strProIDi & " _ " & strProNamei & "' WHERE ProjectID = '" & proID & "' AND TaskID is NULL"
	  		  strQuery = strQuery & " UPDATE ATC_CostStatus SET ProjectID = '" & strProIDi & "' WHERE ProjectID = '" & proID & "'"
	  		  strQuery = strQuery & " UPDATE ATC_Progress SET ProjectID = '" & strProIDi & "' WHERE ProjectID = '" & proID & "'"
	  		  strQuery = strQuery & " UPDATE ATC_ProjectContacts SET ProjectID = '" & strProIDi & "' WHERE ProjectID = '" & proID & "'"
	  		  strQuery = strQuery & " UPDATE ATC_Invoices SET ProjectID = '" & strProIDi & "' WHERE ProjectID = '" & proID & "'"
	  		  ret = objDb.runActionQuery(strQuery)
	  		  if not ret then gMessage = objDb.strMessage
	  		end if
	  		'update type of project
	  		if gMessage = "" then
	  			strQuery = "DELETE ATC_ProjectPrjType WHERE ProjectID = '" & proID & "'"
	  			ret = objDb.runActionQuery(strQuery)
	  			if not ret then 
	  			  gMessage = objDb.strMessage
	  			else
	  				for ii = 1 to countU
	  					strQuery = "INSERT INTO ATC_ProjectPrjType(ProjectID, ProjectTypeID) VALUES('" & strProIDi &_
	  								"', " & int(Request.Form("lsttype")(ii)) & ")"
	  					ret = objDb.runActionQuery(strQuery)
	  					if not ret then 
	  						gMessage = objDb.strMessage
	  						exit for
	  					end if
	  				next
	  			end if			
			end if
		end if
		if gMessage <> "" then
			objDb.cnDatabase.RollbackTrans
			fgChange = false
		else
			objDb.cnDatabase.CommitTrans
			gMessage = "Updated successfully."
			proID = strProIDi
			fgRefresh = "1"
			fgChange = true
		end if
	  end if
	  objDb.dbdisConnect
	else '-----------error connection
	  gMessage = objDb.strMessage
	end if	
	set objDb = nothing
	task_update = fgChange
End function
'****************************************
' function: task_prepare
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_prepare
  'get list of project type
  Set objDb = New clsDatabase
  strConnect = Application("g_strConnect")
  ret = objDb.dbConnect(strConnect)
  if ret then
	ret = objDb.runQuery("SELECT * FROM ATC_ProjectTypes")
	if not objDb.noRecord then	
	  session("typeofproject") = objDb.rsElement.GetRows
	else
	'nothing to be done
	  session("typeofproject") = ""
	end if
	objDb.dbdisConnect
  else 'error
 	gMessage = objDb.strMessage
  end if
  set objDb = nothing	

  proID = Request.Form("txthiddenstrproID")
  if proID="" then 'addnew
	strProID = ""
	strProName = ""
	varInhouse = 0
	varActivate = 1
	strMode = "EDIT"
	strOut = makeSelect(session("typeofproject"), true, "")
	if fgUpdate then fgRightonPro = true
  else 'view or edit
  	strMode="VIEW"
	'get info of specified project
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		strQuery = "select distinct a.ProjectID, a.ProjectName, a.fgActivate, ISNULL(b.CompanyID, 0) Inhouse, ISNULL(c.ProjectTypeID, 0) ProjectType, isnull(d.OwnerID,0) OwnerID " &_
					"FROM ATC_Projects a " &_
					"LEFT JOIN (select CompanyID from ATC_Companies where CompanyID=" & session("InHouse")& ") b ON a.CompanyID = b.CompanyID " &_
					"LEFT JOIN ATC_projectPrjType c ON c.ProjectID = a.ProjectID " &_
					"LEFT JOIN ATC_Tasks d ON d.ProjectID = a.ProjectID WHERE a.ProjectID='" & proID & "' AND d.taskID is null"
		ret = objDb.runQuery(strQuery)
		if not objDb.noRecord then
			strarrType = ""
			Do Until objDb.rsElement.EOF
				strarrType = strarrType & "@" & objDb.rsElement("ProjectType")
				objDb.MoveNext
			Loop
			objDb.MoveFirst
			if Cint(objDb.rsElement("OwnerID"))= session("USERID") then 
				fgRightonPro = true
			end if
			strarrType = strarrType & "@"
			session("selected") = strarrType
			strProID = proID
			strProName = objDb.rsElement("ProjectName")
			if objDb.rsElement("Inhouse") = 0 then 'not in house
				varInhouse = 0
			else
				varInhouse = 1
			end if

			if objDb.rsElement("fgActivate") = true then
				varActivate = 1
			else
				varActivate = 0
			end if
			objDb.CloseRec
		end if
		objDb.dbdisConnect
	else
		'error
		gMessage = objDb.strMessage
	end if
	set objDb = nothing
	strOut = makeSelect(session("typeofproject"), false, strarrType)
  end if
End Sub
'--------------------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	Dim objUser, gMessage, strMode, strProID 'value that is shown in text box
	Dim strProName, varInhouse, varActivate, proID 'value that is got from query string
	Dim strOut, fgRefresh, strarrType, fgUpdate, fgRightonPro 'only people who make this project can delete or update this project
	Dim fgRightProject 'person have right 

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check ACCESS right
'-----------------------------------
'	tmp = Request.ServerVariables("URL") 
'	while Instr(tmp, "/")<>0
'		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
'	Wend
	tmp = Request.Form("txtpreviouspage")
	strFilename = tmp
	tmp = "listofproject.asp"
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if

'-----------------------------------
'Check UPDATE PROJECT right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRightonPro = false
	else
		getRight = session("Righton")
		fgRightonPro = false
		for ii = 0 to Ubound(getRight, 2)
			if lcase(getRight(0, ii)) = lcase("Update project") then
				fgRightonPro=true
				exit for
			end if
		next
		set getRight = nothing
	end if
'---------------------------	
' Get Full Name
'---------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
	
	'Make list of menu
	If isEmpty(session("Menu")) then 
		getRes = getarrMenu(session("USERID"))
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End if	
	
	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")
'------------------------------
' main procedure
'------------------------------
	if Request.QueryString("fgMenu") <> "" then
		fgExecute = false
	else
		fgExecute = true
	end if

	gMessage = ""
	if fgExecute then
		strAct = Request.QueryString("act")
		if Request.QueryString("addsub") <> "" then 'after add subtask
			strAct = "RESTORE"
			fgRefresh = "1"
		end if
	
		if Request.QueryString("addprotype") <> "" then strAct = "RESTORE"
		
		fgOutside = Request.QueryString("outside")		
		if fgOutside = "1" then
			Call freeproInfo
		end if
		
		Call freeListpro
	else
		strAct = "RESTORE"
	end if
	
	if Request.Form("txtrightonpro") = "False" then
		fgRightonPro = false
	elseif Request.Form("txtrightonpro") = "True" then
		fgRightonPro = true
	end if

	select case strAct
	case "RESTORE"
		strMode=Request.Form("txtmode")
		strProID = Request.Form("txtproID")
		proID = strProID
		strProName = Request.Form("txtproName")
		varInhouse = Request.Form("chkinhouse")
		varActivate = Request.Form("chkactivate")
		if varInhouse = "" then varInhouse = 0
		if varActivate = "" then varActivate = 0 else varActivate = 1 end if
		strOut = makeSelect(session("typeofproject"), true, session("selected"))
	case "EDIT" 
		strMode="EDIT"
		strProID = Request.Form("txtproID")
		proID = strProID
		strProName = Request.Form("txtproName")
		varInhouse = Request.Form("chkinhouse")
		varActivate = Request.Form("chkactivate")
		if varInhouse = "" then varInhouse = 0
		if varActivate = "" then varActivate = 0 else varActivate = 1 end if
		strOut = makeSelect(session("typeofproject"), true, session("selected"))
	case "ADD"
		if fgUpdate then fgRightonPro = true
		strProID = ""
		strProName = ""
		varInhouse = 0
		varActivate = 1
		strOut = makeSelect(session("typeofproject"), true, "")
		strMode = "EDIT"
		session("READYPRO") = false
	case "SAVE"
		strMode = "EDIT"
		proID = Lcase(Request.Form("txthiddenproID"))	'Request.Querystring("proID")
		strProID = Lcase(Request.Form("txtproID"))
		strProIDi = Replace(strProID, "'", "''")
		strProIDi = Replace(strProIDi, chr(34), "''")
		strProName = Request.Form("txtproName")
		strProNamei = Replace(strProName, "'", "''")
		strProNamei = Replace(strProNamei, chr(34), "''")
		varInhouse = Request.Form("chkinhouse")
		varActivate = Request.Form("chkactivate")
		if varInhouse = "" then varInhouse = 0
		if varActivate = "" then varActivate = 0 else varActivate = 1 end if
		if varInhouse = 0 then
			varCompany = "NULL"
		else
		    varCompany = session("InHouse")
		end if
		countU = Request.Form("lsttype").Count
		tmp = ""
		for ii = 1 to countU
			tmp = tmp & "@" & Request.Form("lsttype")(ii)
		next
		tmp = tmp & "@"
		strOut = makeSelect(session("typeofproject"), true, tmp)
		session("selected") = tmp
		if proID = "" then 'add new
		  'insert into
		  ret = task_add
		else 'update------------------------------------
		  ret = task_update
		end if
		if ret then
			strMode = "VIEW"
		end if
	case "" 'is called from list of project
		Call task_prepare
		if Request.QueryString("addprotype") <> "" then
			strProID = Request.Form("txtproID")
			strProName = Request.Form("txtproName")
			varInhouse = Request.Form("chkinhouse")
			varActivate = Request.Form("chkactivate")
			if varInhouse = "" then varInhouse = 0
			if varActivate = "" then varActivate = 0 else varActivate = 1 end if
		end if		
	case "REMOVE"
		Call task_remove
		if fgRefresh="1" then
			strMode=Request.Form("txtmode")
			strProID = Request.Form("txtproID")
			proID = strProID
			strProName = Request.Form("txtproName")
			varInhouse = Request.Form("chkinhouse")
			varActivate = Request.Form("chkactivate")
			if varInhouse = "" then varInhouse = 0
			if varActivate = "" then varActivate = 0 else varActivate = 1 end if
			strOut = makeSelect(session("typeofproject"), true, session("selected"))
		else
			strMode=Request.Form("txtmode")
			strProID = Request.Form("txtproID")
			proID = strProID
			strProName = Request.Form("txtproName")
			varInhouse = Request.Form("chkinhouse")
			varActivate = Request.Form("chkactivate")
			if varInhouse = "" then varInhouse = 0
			if varActivate = "" then varActivate = 0 else varActivate = 1 end if
			strOut = makeSelect(session("typeofproject"), true, session("selected"))
		end if
	case "DELETE"
		Call task_delete
	end select

	if proID <>"" then 'draw sub task
	  fgExec = false
	  If isEmpty(session("READYPRO")) or session("READYPRO")<> True or fgRefresh = "1" then
		strConnect = Application("g_strConnect")
		Set objDb = New clsDatabase
		objDb.recConnect(strConnect)

		strQuery = "SELECT DISTINCT a.SubTaskID as sID, a.SubTaskName as sName, ISNULL(a.taskID, 0) as sParentID, ISNULL(a.ChainID, '') as ChainID, " &_
					"ISNULL(b.OwnerID, 0) as Owner, ISNULL(c.StaffID, 0) as RightOn, ISNULL(d.TaskID,'') as Leaf " &_
					"FROM ATC_Tasks a LEFT JOIN (Select OwnerID, SubTaskID From ATC_Tasks Where OwnerID = " & session("USERID") & ") b " &_
					"ON a.SubTaskID = b.SubTaskID " &_
					"LEFT JOIN (Select StaffID, SubTaskID From ATC_RightOnTasks Where StaffID = " & session("USERID") & ") c " &_
					"ON a.SubTaskID = c.SubTaskID LEFT JOIN (Select TaskID from ATC_Tasks WHERE ProjectID = '" & proID & "') d " &_
					"ON a.SubTaskID = d.TaskID WHERE a.ProjectID = '" & proID & "' ORDER BY sID"
		If objDb.openRec(strQuery) Then
		  objDb.recDisConnect
		  if not objDb.noRecord then
			set rsTask = objDb.rsElement.Clone
			session("READYPRO") = true
			rsTask.MoveFirst
			set session("rsTaskCache") = rsTask
			fgExec = true
		  else
			gMessage = "No data."
		  end if
		  objDb.CloseRec
		Else
		  gMessage = objDb.strMessage
		End if
		Set objDb = Nothing
	  else
		set rsTask = session("rsTaskCache")
		fgExec = true
	  end if
	  
	  if fgExec then
		set objUser = Server.CreateObject("ADODB.Recordset")
		Dim arrRs(4)
		'-- Create the ADO Objects
		For i = 0 to 4
		  set arrRs(i) = Server.CreateObject("ADODB.Recordset")
		  Call SetAttRs(arrRs(i))
		Next
		on error resume next
  		rsTask.Filter = "sParentID = 0"
  		if Err.number>0 then
  			Response.Write Err.description & ", " & Err.number
			Err.Clear  	  
  		end if
		arrRs(0).AddNew Array("sID", "sName", "sParentID", "ChainID", "Owner", "RightOn", "Leaf"),_
						 Array(rsTask(0), rsTask(1), rsTask(2), rsTask(3), rsTask(4), rsTask(5), rsTask(6))
		rsTask.Filter = ""
		strTree = ""
		strLast = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & chr(13) & _
					  "  <tr><td bgcolor='#DDDDDD'>" &_
					  "<table width='100%' border='0' cellspacing='1' cellpadding='5'>"
		arrRs(0).MoveFirst
		rsTask.MoveFirst
		if arrRs(0)("Owner") <> 0 or arrRs(0)("RightOn") <> 0 then parOwner = true else parOwner = false end if
		FetchChild arrRs(0), strTree, 0, rsTask, parOwner
		strLast = strLast & strTree & chr(13)
		strLast = strLast & "</table></td></tr></table>"
	  
		For i = 0 to 4
		  arrRs(i).Close
			Set arrRs(i) = Nothing
		Next
	  end if
	end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>
<script LANGUAGE="JavaScript">
var objNewWindow;

function addsub() { //v2.0
var strproid = "<%=proID%>";
  window.status = "";
  strFeatures = "top="+(screen.height/2-92)+",left="+(screen.width/2-132)+",width=265,height=184,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no";
  if((objNewWindow) && (!objNewWindow.closed))
	objNewWindow.focus();	
  else {
	objNewWindow = window.open("addsubtask.asp?proid=" + strproid, "MyNewWindow", strFeatures);
  }
  window.status = "Opened a new browser window.";  
}

function addprotype() { //v2.0
  window.status = "";
  strFeatures = "top="+(screen.height/2-78)+",left="+(screen.width/2-132)+",width=265,height=158,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no";
  if((objNewWindow) && (!objNewWindow.closed))
	objNewWindow.focus();	
  else {
	objNewWindow = window.open("addprotype.asp", "MyNewWindow", strFeatures);
  }
  window.status = "Opened a new browser window.";  
}

function window_onunload() {
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.close();
}

function CheckMode(field){
var varMode="<%=strMode%>";
    if (varMode!="EDIT"){
        field.blur();
	}
}

function restorelist(obj, strSelected) {
	for (var i = 0; i < obj.options.length; i++) {
		if (strSelected.indexOf("@" + obj.options[i].value + "@")!=-1)
			obj.options[i].selected = true
		else
			obj.options[i].selected = false
	}
}

function CheckData() {
	if (isnull(document.proinfo.txtproID.value)==true) {
		alert("Please enter value for this field.");
		document.proinfo.txtproID.focus();
		return false ;
	}
	else {
		var spec = String.fromCharCode(34);
		var spec2 = String.fromCharCode(39);
		var tmp = document.proinfo.txtproID.value;
		if (tmp.substring(0,3).toUpperCase()=="TEMP")
		{
			alert("Project ID cannot begin with 'TEMP'");
			document.proinfo.txtproID.focus();
			return false;
		}
		else if(tmp.indexOf(spec)!=-1) {
			alert(spec + " can not be present in ProjectID.");
			document.proinfo.txtproID.focus();
			return false;
		}
		else if(tmp.indexOf(spec2)!=-1) {
			alert(spec2 + " can not be present in ProjectID.");
			document.proinfo.txtproID.focus();
			return false;
		}
	}
	
	if (isnull(document.proinfo.txtproName.value)==true) {
		alert("Please enter value for this field.");
		document.proinfo.txtproName.focus();
		return false ;
	}
	
	if (document.proinfo.lsttype.selectedIndex == -1) {
		alert("Please choose at least one type");
		document.proinfo.lsttype.focus();
		return false ;
	}
	return true;
}

function savedata() {
	if (CheckData()==true) {
		document.proinfo.action = "projectinfo.asp?act=SAVE";	//&proID=" + proid;
		document.proinfo.target = "_self";
		document.proinfo.submit();
	}
}

function edit() {
	document.proinfo.action = "projectinfo.asp?act=EDIT";	//&proID=" + proid;
	document.proinfo.target = "_self";
	document.proinfo.submit();
}

function add() {
	document.proinfo.action = "projectinfo.asp?act=ADD";
	document.proinfo.target = "_self";
	document.proinfo.submit();
}

/*function window_onload() {
var tmpConfirm = "<%=gConfirm%>";
	if (tmpConfirm != "") 
		if (confirm(tmpConfirm)) {
			var proid = "<%=strproID%>"
			document.proinfo.action = "projectinfo.asp?act=DELETE&confirmed=1&proID=" + proid;
			document.proinfo.target = "_self";
			document.proinfo.submit();
		}			
}
onLoad="return window_onload();"
*/

function setchecked(val) {
  with (document.proinfo) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if (elements[ii].name == "chkrem") {
			elements[ii].checked = val;
		}
	}
  }
}

function chkremove() {
  fg = false;
  with (document.proinfo) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chkrem") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 if (fg == false) alert("No task selected.")
 return(fg)
}

function remove() {
	if(chkremove()==true) {
		document.proinfo.action = "projectinfo.asp?act=REMOVE";	
		document.proinfo.target = "_self";
		document.proinfo.submit();
	}
}

function mydelete() {
  if(document.proinfo.txthiddenstrproID.value!="") {
	if(confirm("Are you sure you want to delete this project?")) {
		document.proinfo.action = "projectinfo.asp?act=DELETE";	
		document.proinfo.target = "_self";
		document.proinfo.submit();
	}
  }
}

function assign() {
	document.proinfo.action = "assignment.asp?outside=1";	//&proid=" + proid + "&proname=" + proname;
	document.proinfo.target = "_self";
	document.proinfo.submit();
}

function assignright() {
	document.proinfo.action = "assignright.asp?outside=1";	//&proid=" + proid + "&proname=" + proname;
	document.proinfo.target = "_self";
	document.proinfo.submit();
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onUnload="return window_onunload();">
<form name="proinfo" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>

			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))

			%>		
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
          <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
        </tr>
        <tr> 
          <td class="blue" align="left" width="23%">&nbsp;&nbsp;
		<a href="listofproject.asp" onMouseOver="self.status='Show the list of projects'; return true;" onMouseOut="self.status=''">
		Project List</a> </td>
          <td class="blue" width="77%">&nbsp;</td>
        </tr>
        <tr align="center"> 
          <td class="title" height="50" align="center" colspan="2"> Projects Information</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="blue-normal" width="33%">&nbsp;</td>
                <td class="blue-normal" width="15%">Project ID</td>
                <td colspan="2"> 
                  <input type="text" maxlength="20" name="txtproID" class="blue-normal" value="<%=Showvalue(strProID)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>>
                </td>
              </tr>
              <tr> 
                <td class="blue-normal" width="33%">&nbsp;</td>
                <td class="blue-normal" width="15%">Project Name</td>
                <td colspan="2"> 
                  <input type="text" maxlength="50" name="txtproName" class="blue-normal" value="<%=Showvalue(strProName)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>>
                </td>
              </tr>
              <tr> 
                <td class="blue-normal" width="33%">&nbsp;</td>
                <td class="blue-normal" width="15%">Type</td>
                <td colspan="2" class="blue-normal">
<%				Response.Write strOut
%>                                  
                  </td>
              </tr>
              <tr> 
                <td width="33%">&nbsp;</td>
                <td width="15%" class="blue-normal" valign="middle">In House Project </td>
                <td valign="middle" colspan="2"> 
                  <input type="checkbox" name="chkinhouse" value="1" <%if varInhouse <> 0 then%>checked <%end if%> <%if strMode<>"EDIT" then%>onClick="return false;" <%end if%>>
                </td>
              </tr>
              <tr> 
                <td width="33%">&nbsp;</td>
                <td width="15%" class="blue-normal" valign="middle">Activate </td>
                <td colspan="2" class="blue-normal" valign="middle"> 
                  <input type="checkbox" name="chkactivate" value="1" <%if varActivate <> 0 then%>checked <%end if%> <%if strMode<>"EDIT" then%>onClick="return false;" <%end if%>>
                </td>
              </tr>
              <tr> 
                <td colspan="4" height="40"> 
                  <table width="360px" border="0" cellspacing="5" cellpadding="0" align="center" height="20">
                    <tr> 
<% if fgRightonPro and fgUpdate then %>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center"> 
                          <a href="javascript:add();" class="b" onMouseOver="self.status='Add a new project'; return true;" onMouseOut="self.status=''">Add</a></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if strMode<>"EDIT" and proID<>"" then%><a href="javascript:edit();" class="b" onMouseOver="self.status='Edit project'; return true;" onMouseOut="self.status=''">Edit</a>
<%else %>Edit<%end if%></td>
					  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if strMode<>"VIEW" then%><a href="javascript:savedata();" class="b" onMouseOver="self.status='Save changes'; return true;" onMouseOut="self.status=''">Save</a>
<%else%>Save<%end if%></td>
					  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if proID<>"" then%><a href="javascript:mydelete();" class="b" onMouseOver="self.status='Delete project'; return true;" onMouseOut="self.status=''">Delete</a>
<%else%>Delete<%end if%></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if proID<>"" then%><a href="javascript:assign();" class="b" onMouseOver="self.status='Assignment'; return true;" onMouseOut="self.status=''">Assign</a>
<%else%>Assign<%end if%></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
<%if proID<>"" then%><a href="javascript:assignright();" class="b" onMouseOver="self.status='Right on tasks'; return true;" onMouseOut="self.status=''">Right</a>
<%else%>Right<%end if%></td>
<% else	'don't have right%>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
<%if fgUpdate then%><a href="javascript:add();" class="b" onMouseOver="self.status='Add a new project'; return true;" onMouseOut="self.status=''">Add</a><%else%>Add<%end if%></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">Edit</td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">Save</td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">Delete</td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
                      <a href="javascript:assign();" class="b" onMouseOver="self.status='Assignment'; return true;" onMouseOut="self.status=''">Assign</a></td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
                      <a href="javascript:assignright();" class="b" onMouseOver="self.status='Right on tasks'; return true;" onMouseOut="self.status=''">Right</a></td>
<%end if%>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
<%If proID<>"" then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="blue" align="left" height="20">&nbsp;&nbsp;
	<%if fgUpdate then%><a href="javascript:addsub();" onMouseOver="self.status='Add a sub-task'; return true;" onMouseOut="self.status=''">Add Sub-Task</a>
	<%else%>Add Sub-Task<%end if%>
                  &nbsp;</td>
              </tr>
            </table>
	<%Response.Write strLast%>            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <tr>
	<%if fgUpdate then%>
                 <td class="blue-normal" align="left" height="20" width="69%">&nbsp;&nbsp;*Choose 
                   the checkbox, then click &quot;remove&quot; to remove sub-task.</td>
                 <td class="blue" align="right" height="20" width="31%">&nbsp;<a href="javascript:setchecked(1);" onMouseOver="self.status='Check all'; return true;" onMouseOut="self.status=''">Check 
                   All</a>&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='Clear all'; return true;" onMouseOut="self.status=''">Clear All</a>&nbsp;&nbsp;&nbsp; 
                   <a href="javascript: remove();" onMouseOver="self.status='Remove sub-task'; return true;" onMouseOut="self.status=''"> Remove</a> &nbsp;</td>
	<%else%>
                 <td class="blue-normal" align="left" height="20" width="69%">&nbsp;&nbsp;*Choose 
                   the checkbox, then click &quot;remove&quot; to remove sub-task.</td>
                 <td class="blue" align="right" height="20" width="31%">&nbsp;Check 
                   All&nbsp;&nbsp;&nbsp; Clear All&nbsp;&nbsp;&nbsp; 
                    Remove &nbsp;</td>
	<%end if%>
               </tr>
             </table>
<%End if%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
      </table>
    </td>
  </tr>
</table>
			<%
			Response.Write(arrTmp(1))
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))
			%>
<input type="hidden" value="<%=strMode%>" name="txtmode">
<input type="hidden" name="txtrightonpro" value="<%=fgRightonPro%>">
<input type="hidden" name="txthiddenproID" value="<%=proID%>">
<input type="hidden" name="txthiddenstrproID" value="<%=strproID%>">
<input type="hidden" name="txthiddenstrproName" value="<%=strproName%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>