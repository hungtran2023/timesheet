<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%
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
	if strleavedue = "" or strleavedue = "0" then
		strleavedue_I = "Null"
	else
		strleavedue_I = strleavedue
	end if
	
	strfirst_I = "'"& Replace(strfirst, "'", "''") &"'"

	if strmiddle = "" then 
		strmiddle_I = "Null"
	else 
		strmiddle_I = "'"& Replace(strmiddle, "'", "''") & "'"
	end if

	strsurname_I = "'"& Replace(strsurname, "'", "''") &"'"
	strtitleU_I = "'"& strtitleU &"'"

	if straddress = "" then 
		straddress_I = "Null"
	else 
		straddress_I = "'"& Replace(straddress, "'", "''") &"'"
	end if

	if strcity = "" then 
		strcity_I = "Null"
	else 
		strcity_I = "'"& Replace(strcity, "'", "''") &"'"
	end if

	if strstate = "" then 
		strstate_I = "Null"
	else 
		strstate_I = "'"& Replace(strstate, "'", "''") &"'"
	end if

	if strpostcode = "" then 
		strpostcode_I = "Null"
	else 
		strpostcode_I = "'"& Replace(strpostcode, "'", "''") &"'"
	end if

	if stremail = "" then 
		stremail_I = "Null"
	else 
		stremail_I = "'"& Replace(stremail, "'", "''") &"'"
	end if

	if strExemail = "" then 
		strExemail_I = "Null"
	else 
		strExemail_I = "'"& Replace(strExemail, "'", "''") &"'"
	end if

	if strphone = "" then 
		strphone_I = "Null"
	else 
		strphone_I = "'"& Replace(strphone, "'", "''") &"'"
	end if

	if strmobile = "" then 
		strmobile_I = "Null"
	else 
		strmobile_I = "'"& Replace(strmobile, "'", "''") &"'"
	end if
			
	if stridnum = "" then 
		stridnum_I = "Null" 
	else 
		stridnum_I = "'"& stridnum &"'"
	end if
	
	if strTelExt = "" then 
		strTelExt_I = "Null" 
	else 
		strTelExt_I = "'"& strTelExt &"'"
	end if
			
	strusernameU_I = "'"& Ucase(Replace(strusernameU, "'", "''")) &"'"

	varbirth_I = "'" & CStr(varbirth) & "'"
	varstart_I = "'" & CStr(varstart) & "'"
	if varleave<>"" then
		varleave_I = "'" & CStr(varleave) & "'"
	else
		varleave_I = "Null"
	end if
	
	Set objSHA1 = New clsSHA1
	strpass_I = "'"& ObjSHA1.SecureHash(strpass) &"'"
	Set ObjSHA1 = Nothing

	if intreport = 0 then 
		intreport_I = "Null"
	else
		intreport_I = intreport
	end if
			
	if varIndirect=true then 
		varIndirect_I = "1"
	else
		varIndirect_I = "0"
	end if

	stroldname = Request.Form("txtnameold")
	stroldname = "'"& Ucase(Replace(stroldname, "'", "''")) &"'"
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then '---------------connection
	  if ucase(stroldname)<>ucase(strusernameU_I) then
		strQuery = "Select count(*) as mysum From ATC_Users WHERE UserName = " & strusernameU_I
		ret = objDb.runQuery(strQuery)
		if ret then
			if objDb.rsElement("mysum") > 0 then
			  gMessage = "This UserName has already been inputted."
			else 
			  strQuery = "UPDATE ATC_Users SET UserName = " & strusernameU_I & " WHERE UserID = " & struserID
	  		end if
	  	else
	  		gMessage = objDb.strMessage
	  	end if
	  else
		strQuery = ""
	  end if	  

	  if gMessage="" then '-----------starting
	    objDb.cnDatabase.BeginTrans
	    if strQuery<>"" then	    
		  ret = objDb.runActionQuery(strQuery)
	      if not ret then gMessage = objDb.strMessage
	    end if
	    if gMessage="" then
	  		'update table ATC_PersonalInfo
			strQuery = "UPDATE ATC_PersonalInfo SET Title=" & strtitleU_I & ", Lastname=" & strsurname_I & ", Middlename=" &_
					strmiddle_I & ", Firstname=" & strfirst_I & ", Gender=" & vargender & ", Birthday=" & varbirth_I &_
					", NationalityID=" & intnat & ", Address=" & straddress_I & ", City=" & strcity_I & ", State=" & strstate_I &_
					", PostalCode=" & strpostcode_I & ", CountryID=" & intCountry & ", EmailAddress=" & stremail_I & ", EmailAddress_Ex=" & strExemail_I & ", Phone=" &_
					strphone_I & ", MobilePhone=" & strmobile_I & ", IDNumber=" & stridnum_I & " WHERE PersonID = " & struserID
		
	  		ret = objDb.runActionQuery(strQuery)
	  		if not ret then 
	  		  gMessage = objDb.strMessage 
	  		else
	  		  strQuery = "UPDATE ATC_Employees SET JobTitleID=" & intjobtitle & ", DepartmentID=" & intdepartment &_
	  					", DirectLeaderID=" & intreport_I & ", JoinDate=" & varstart_I & ", Leavedate=" & varleave_I &_
	  					", fgIndirect=" & varIndirect_I & ", Leavedue = " & strleavedue_I & ",ExtPhone=" & strTelExt_I & " WHERE StaffID = " & struserID
	  		  ret = objDb.runActionQuery(strQuery)
	  		  if not ret then gMessage = objDb.strMessage
			end if
		end if
		fgChanged = false	
		if gMessage <> "" then
	 		objDb.cnDatabase.RollbackTrans
		else
			objDb.cnDatabase.CommitTrans
			gMessage = "Updated successfully."
			fgChanged = true
		end if
		objDb.dbdisConnect
	  end if
	else '-----------error connection
	  gMessage = objDb.strMessage
	end if	
	set objDb = nothing
	task_update = fgChanged
End Function
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
function task_add
	if strleavedue = "" or strleavedue = "0" then
		strleavedue_I = "Null"
	else
		strleavedue_I = strleavedue
	end if

	strfirst_I = "'"& Replace(strfirst, "'", "''") &"'"

	if strmiddle = "" then 
		strmiddle_I = "Null"
	else 
		strmiddle_I = "'"& Replace(strmiddle, "'", "''") & "'"
	end if

	strsurname_I = "'"& Replace(strsurname, "'", "''") &"'"
	strtitleU_I = "'"& strtitleU &"'"

	if straddress = "" then 
		straddress_I = "Null"
	else 
		straddress_I = "'"& Replace(straddress, "'", "''") &"'"
	end if

	if strcity = "" then 
		strcity_I = "Null"
	else 
		strcity_I = "'"& Replace(strcity, "'", "''") &"'"
	end if

	if strstate = "" then 
		strstate_I = "Null"
	else 
		strstate_I = "'"& Replace(strstate, "'", "''") &"'"
	end if

	if strpostcode = "" then 
		strpostcode_I = "Null"
	else 
		strpostcode_I = "'"& Replace(strpostcode, "'", "''") &"'"
	end if

	if stremail = "" then 
		stremail_I = "Null"
	else 
		stremail_I = "'"& Replace(stremail, "'", "''") &"'"
	end if
	
	if strExemail = "" then 
		strExemail_I = "Null"
	else 
		strExemail_I = "'"& Replace(strExemail, "'", "''") &"'"
	end if
	if strphone = "" then 
		strphone_I = "Null"
	else 
		strphone_I = "'"& Replace(strphone, "'", "''") &"'"
	end if

	if strmobile = "" then 
		strmobile_I = "Null"
	else 
		strmobile_I = "'"& Replace(strmobile, "'", "''") &"'"
	end if
			
	if stridnum = "" then 
		stridnum_I = "Null" 
	else 
		stridnum_I = "'"& stridnum &"'"
	end if
	
	if strTelExt = "" then 
		strTelExt_I = "Null" 
	else 
		strTelExt_I = "'"& strTelExt &"'"
	end if
			
	strusernameU_I = "'"& Ucase(Replace(strusernameU, "'", "''")) &"'"

	varbirth_I = "'" & CStr(varbirth) & "'"
	varstart_I = "'" & CStr(varstart) & "'"
	if varleave<>"" then
		varleave_I = "'" & CStr(varleave) & "'"
	else
		varleave_I = "Null"
	end if
	
	Set objSHA1 = New clsSHA1
	strpass_I = "'"& ObjSHA1.SecureHash(strpass) &"'"
	Set ObjSHA1 = Nothing

	if intreport = 0 then 
		intreport_I = "Null"
	else
		intreport_I = intreport
	end if
			
	if varIndirect=true then 
		varIndirect_I = "1"
	else
		varIndirect_I = "0"
	end if
	

  Set objDb = New clsDatabase
  strConnect = Application("g_strConnect")
  ret = objDb.dbConnect(strConnect)
  if ret then
'---------------------------
' Check duplicate username
'---------------------------	
	strQuery = "Select count(*) as mysum From ATC_Users WHERE Username = " & strusernameU_I

	ret = objDb.runQuery(strQuery)
	if ret then
		if objDb.rsElement("mysum") > 0 then gMessage = "This UserName has already been inputted."
	else
		gMessage = objDb.strMessage
	end if
	if gMessage="" then '---starting
	  objDb.cnDatabase.BeginTrans	  	  
	  strQuery = "INSERT INTO ATC_PersonalInfo(Title, Lastname, Middlename, Firstname, Gender, Birthday, NationalityID, " &_
			"Address, City, State, PostalCode, CountryID, EmailAddress,EmailAddress_Ex, Phone, MobilePhone, IDNumber, CompanyID, UserType) " &_
	  		"VALUES(" & strtitleU_I & ", " & strsurname_I & ", " & strmiddle_I & ", " & strfirst_I & ", " & vargender & ", " &_
	  		varbirth_I & ", " & intnat & ", " & straddress_I & ", " & strcity_I & ", " & strstate_I & ", " & strpostcode_I & ", " &_
	  		intCountry & ", " & stremail_I & ", " & strExemail_I & ", " & strphone_I & ", " & strmobile_I & ", " & stridnum_I & ", " & session("Inhouse") & ", 1)"
	  		
	  ret = objDb.runActionQuery(strQuery)
	  if ret then
	  '------insert table atc_Users
		strQuery = "SELECT @@IDENTITY as ID"
		ret = objDb.runQuery(strQuery)
		if ret then
			intpersonID = objDb.rsElement("ID")

			strQuery = "INSERT INTO ATC_Users(UserID, userName, Password) VALUES(" & intpersonID &_
	  				", " & strusernameU_I & ", " & strpass_I & ")"
			ret = objDb.runActionQuery(strQuery)
			if not ret then 
			  gMessage = objDb.strMessage
			else
			'------insert table ATC_Employees
				strQuery = "INSERT INTO ATC_Employees(StaffID, JobTitleID, DepartmentID, DirectLeaderID, JoinDate, Leavedate, " &_
					"fgIndirect, LeaveDue,extphone) VALUES(" & intpersonID & ", " & intjobtitle & ", " & intdepartment & ", " & intreport_I &_
					", " & varstart_I & ", " & varleave_I & ", " & varIndirect_I & ", " & strleavedue_I & ", " & strTelExt_I & ")"
				ret = objDb.runActionQuery(strQuery)
		
				if ret then
				'---------------insert table ATC_SalarySatus
				  strQuery = "INSERT INTO ATC_SalaryStatus(StaffID, Salarydate, WorkingHourID, Salary, SalaryTax) VALUES(" &_
						intpersonID & ", " & varstart_I & ", " & intWorkingHour & ", '" & encode(0, 128) & "', '" & encode(0, 128) & "')"
				  ret = objDb.runActionQuery(strQuery)
				  if not ret then gMessage = objDb.strMessage
				else
				  gMessage = objDb.strMessage
				end if
			end if
		else
		  gMessage = objDb.strMessage
		end if
	  else
	    gMessage = objDb.strMessage
	  end if
	  fgChanged = false
	  if gMessage<>"" then 
	  	objDb.cnDatabase.RollbackTrans
	  else
	  	objDb.cnDatabase.CommitTrans
	  	gMessage = "Added successfully."
	  	fgChanged = true
	  	struserID = intpersonID
	  end if
	end if'-----------ending
	objDb.dbdisConnect
  else
	gMessage = objDb.strMessage
  end if	
  set objDb = nothing
  task_add = fgChanged
end function
'****************************************
' function: task_show
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_show
	strQuery = "SELECT isnull(a.CountryID, 0) as CountryID, isnull(a.NationalityID, 0) as NationalityID, a.Firstname, " &_
	"isnull(a.Lastname, '') as Lastname, isnull(a.Middlename, '') as Middlename, isnull(a.IDNumber, '') as IDNumber, " &_
	"isnull(a.Birthday, '') as Birthday, a.Gender, isnull(a.Title, '')as Title, isnull(a.Address, '') as Address, " &_
	"isnull(a.City, '') as City, isnull(a.State, '') as State, isnull(a.PostalCode, '') as PostalCode, " &_
	"isnull(a.Phone, '') as Phone, isnull(a.MobilePhone, '') as Mobile, isnull(a.EmailAddress, '') as EmailAddress,isnull(a.EmailAddress_ex, '') as EmailAddress_Ex, " &_
	"b.JoinDate, b.LeaveDate, b.fgIndirect, isnull(b.DirectLeaderID, 0) as DirectLeaderID, b.JobTitleID, b.DepartmentID, " &_
	"c.UserName, c.Password, isnull(b.Leavedue, 0) leavedue, isnull(b.ExtPhone, '') ExtPhone,isnull(c.Photo,c.UserName) photo " &_
	"FROM ATC_PersonalInfo a LEFT JOIN ATC_Employees b ON a.PersonID = b.StaffID " &_
	"LEFT JOIN ATC_Users c ON a.PersonID = c.UserID WHERE PersonID = " & struserID
'Response.Write strQuery
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		ret = objDb.runQuery(strQuery)
		if ret then
			strleavedue = objDb.rsElement("leavedue")
			if strleavedue = "0" then strleavedue = ""
			strfirst = objDb.rsElement("Firstname")
			strmiddle = objDb.rsElement("Middlename")
			strsurname = objDb.rsElement("Lastname")
			strtitleU = objDb.rsElement("Title")
			vargender = objDb.rsElement("Gender")
			varbirth = objDb.rsElement("Birthday")
			intnat = objDb.rsElement("NationalityID")
			straddress = objDb.rsElement("Address")
			strcity = objDb.rsElement("City")
			strstate = objDb.rsElement("State")
			strpostcode = objDb.rsElement("PostalCode")
			intCountry = objDb.rsElement("CountryID")
			stremail = objDb.rsElement("EmailAddress")
			strExemail=objDb.rsElement("EmailAddress_Ex")
			strphone = objDb.rsElement("Phone")
			strmobile = objDb.rsElement("Mobile")
			stridnum = objDb.rsElement("IDNumber")
			strusernameU = objDb.rsElement("UserName")
			strpass = objDb.rsElement("Password")
			varstart = objDb.rsElement("JoinDate")
			varleave = objDb.rsElement("leavedate")
			intjobtitle = objDb.rsElement("JobTitleID")
			intdepartment = objDb.rsElement("DepartmentID")
			intreport = objDb.rsElement("DirectleaderID")
			varIndirect = objDb.rsElement("fgIndirect")
			strTelExt=objDb.rsElement("extPhone")
			strPhoto=objDb.rsElement("photo")
			
			strQuery = "select WorkingHourID from atc_salarystatus where staffid = " & struserID & " and SalaryDate in " &_
					"(select max(SalaryDate) from atc_SalaryStatus where staffid = " & struserID & ")"
			ret = objDb.runQuery(strQuery)

			if ret then
				if not objDb.noRecord then 
				  intWorkingHour = objDb.rsElement("WorkingHourID")
				else
				  intWorkingHour = 0
				end if
			else 
				gMessage = objDb.strMessage
			end if
		else
		  gMessage = objDb.strMessage
		end if
	else
		gMessage = objDb.strMessage
	end if
	objDb.dbdisConnect
	set objDb = nothing
end sub
'****************************************
' function: task_initvar
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_initvar
	strMode = "EDIT"
	strleavedue = ""
	strfirst = ""
	strmiddle = ""
	strsurname = ""
	strtitleU = "Mr"
	vargender = true
	varbirth = ""
	intnat = 0
	straddress = ""
	strcity = ""
	strstate = ""
	strpostcode = ""
	intCountry = 0
	stremail = ""
	strExemail=""
	strphone = ""
	strmobile = ""
	stridnum = ""
	strusernameU = ""
	strpass = ""
	varstart = ""
	varleave = ""
	intjobtitle = 0
	intdepartment = 0
	intreport = 0
	varIndirect = false
	intWorkingHour = 0
	strTelExt=""
	strphoto=""
end sub
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
Function task_delete
	fgdel = false
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		strQuery = "UPDATE ATC_PersonalInfo SET fgDelete=1 WHERE PersonID = " & struserID
		ret = objDb.runActionQuery(strQuery)
		if ret then
			'gMessage = "Deleted successfully."
			fgdel = true
		else
			gMessage = objDb.strMessage
		end if
	else
		gMessage = objDb.strMessage
	end if
	objDb.dbdisConnect
	set objDb = nothing
	task_delete = fgdel
End function
'-------------------------------------------
	Dim varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	dim arrlstBirth(2), arrlstStart(2), arrlstLeave(2)
	Dim strfirst, strmiddle, strsurname, strtitleU, vargender, varbirth, intnat, straddress, strcity
	Dim strstate, strpostcode, intCountry, stremail, strphone, strmobile, stridnum, strusernameU, strleavedue
	Dim strpass, varstart, varleave, intjobtitle, intdepartment, intreport, varIndirect, intWorkingHour, struserID,strTelExt,strPhoto
	Dim strExemail
	Dim gMessage, fgChanged

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

'----------------------------------
' Get Full Name and Job Title
'----------------------------------
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
'----------------------------------	
' Make list of menu
'----------------------------------
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
	if strChoseMenu = "" then strChoseMenu = "AB"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'----------------------------------------
' analyse query string
'----------------------------------------
	gMessage = ""
	if Request.QueryString("fgMenu") <> "" then
		fgExecute = false
	else
		fgExecute = true
	end if
	
	Call freeListemp
	strAct = Request.QueryString("act")
	if Request.QueryString("addmore") <> "" then strAct = "ADDMORE"
	
	if not fgExecute then 
		strAct = "RESTORE"
	else
		if not fgUpdate and (strAct="EDIT" or strAct="ADD" or strAct="SAVE" or strAct="DELETE") then
			strAct = "RESTORE"
		end if
	end if

	fgChanged = false

	select case strAct
	case "RESTORE"
		strMode = Request.Form("txtmode")
		struserID = Request.Form("txtuserid")
		if struserID = "" then
			Call task_initvar
		else
			Call task_show
		end if
	case "VIEW" 'was called from list of emp
		strMode = "VIEW"
		struserID =	Request.Form("txthidden")
		Call task_show
	case "ADD"
		strMode = "EDIT"
		Call task_initvar
	case "EDIT"
		strMode = "EDIT"
		struserID = Request.Form("txtuserid")
		Call task_show
	case "DELETE"
		struserID = Request.Form("txtuserid")
		ret = task_delete
		if ret then Response.Redirect "listofemployee.asp?act=REFRESH"
	case "SAVE"
		strleavedue = trim(Request.Form("txtleavedue"))
		struserID = Request.Form("txtuserid")
		strfirst = trim(Request.Form("txtfirst"))
		strmiddle = trim(Request.Form("txtmiddle"))
		strsurname = trim(Request.Form("txtsurname"))
		strtitleU = Request.Form("lsttitle")
		vargender = Request.Form("optgender")
		varbirth = Cdate(Request.Form("lstmonbirth")&"/"&Request.Form("lstdaybirth")&"/"&Request.Form("lstyearbirth"))
		intnat = Request.Form("lstnationality")
		straddress = trim(Request.Form("txtaddress"))
		strcity = trim(Request.Form("txtcity"))
		strstate = trim(Request.Form("txtstate"))
		strpostcode = trim(Request.Form("txtpost"))
		intCountry = Request.Form("lstcountry")
		stremail = trim(Request.Form("txtemail"))
		strExemail=trim(Request.Form("txtExemail"))
		strphone = trim(Request.Form("txtphone"))
		strmobile = trim(Request.Form("txtmobile"))
		stridnum = trim(Request.Form("txtidnum"))
		strusernameU = trim(Request.Form("txtusername"))
		strpass =  Request.Form("txtpass")
		varstart = Cdate(Request.Form("lstmonstart")&"/"&Request.Form("lstdaystart")&"/"&Request.Form("lstyearstart"))
		
		if Request.Form("lstmonleave") = "" then 
			varleave = ""
		else 
			varleave = Cdate(Request.Form("lstmonleave")&"/"&Request.Form("lstdayleave")&"/"&Request.Form("lstyearleave"))
		End if

		intjobtitle = trim(Request.Form("lstjobtitle"))
		intdepartment = trim(Request.Form("lstdepartment"))
		intreport = trim(Request.Form("lstreportto"))
		if intreport="" then intreport=0

		varIndirect = Request.Form("chkindirect")
		if varIndirect="" then 
			varIndirect = false
		else
			varIndirect = true
		end if
		intWorkingHour = trim(Request.Form("lstworkinghour"))
		strTelExt=trim(Request.Form("txtTelExt"))
		if intWorkingHour = "" then intWorkingHour = 0

		if struserID="" then
			ret = task_add
		else
			ret = task_update
		end if
		
		if vargender = "1" then
			vargender=true
		else
			vargender=false
		end if
				
		if ret then 
			strMode="VIEW"
		else
			strMode="EDIT"
		end if
	case "ADDMORE" 'After add more department, job title, ...
		struserID = Request.Form("txtuserid")
		strfirst = trim(Request.Form("txtfirst"))
		strmiddle = trim(Request.Form("txtmiddle"))
		strsurname = trim(Request.Form("txtsurname"))
		strtitleU = Request.Form("lsttitle")
		vargender = Request.Form("optgender")
		varbirth = Cdate(Request.Form("lstmonbirth")&"/"&Request.Form("lstdaybirth")&"/"&Request.Form("lstyearbirth"))
		intnat = Request.Form("lstnationality")
		straddress = trim(Request.Form("txtaddress"))
		strcity = trim(Request.Form("txtcity"))
		strstate = trim(Request.Form("txtstate"))
		strpostcode = trim(Request.Form("txtpost"))
		intCountry = Request.Form("lstcountry")
		stremail = trim(Request.Form("txtemail"))
		strExemail = trim(Request.Form("txtExemail"))		
		strphone = trim(Request.Form("txtphone"))
		strmobile = trim(Request.Form("txtmobile"))
		stridnum = trim(Request.Form("txtidnum"))
		strusernameU = trim(Request.Form("txtusername"))
		strpass =  Request.Form("txtpass")
		varstart = Cdate(Request.Form("lstmonstart")&"/"&Request.Form("lstdaystart")&"/"&Request.Form("lstyearstart"))
		
		if Request.Form("lstmonleave") = "" then 
			varleave = ""
		else 
			varleave = Cdate(Request.Form("lstmonleave")&"/"&Request.Form("lstdayleave")&"/"&Request.Form("lstyearleave"))
		End if

		intjobtitle = trim(Request.Form("lstjobtitle"))
		intdepartment = trim(Request.Form("lstdepartment"))
		intreport = trim(Request.Form("lstreportto"))
		strTelExt=trim(Request.Form("txtTelExt"))
		if intreport="" then intreport=0

		varIndirect = Request.Form("chkindirect")
		if varIndirect="" then 
			varIndirect = false
		else
			varIndirect = true
		end if
		intWorkingHour = trim(Request.Form("lstworkinghour"))
		if intWorkingHour = "" then intWorkingHour = 0
		if vargender = "1" then
			vargender=true
		else
			vargender=false
		end if				
		strMode="EDIT"
	end select

'----------------------------------------
' Prepare form
'----------------------------------------
	intTmp = year(now())
	if varbirth<>"" then
		arrlstBirth(0) = selectmonth("lstmonbirth", month(varbirth), 0)	
		arrlstBirth(1) = selectday("lstdaybirth", day(varbirth), 0)
		arrlstBirth(2) = selectyear("lstyearbirth", year(varbirth), intTmp - 70, intTmp - 20, 0)
	else
		arrlstBirth(0) = selectmonth("lstmonbirth", 1, 0)
		arrlstBirth(1) = selectday("lstdaybirth", 1, 0)
		arrlstBirth(2) = selectyear("lstyearbirth", intTmp, intTmp - 70, intTmp - 20, 0)
	end if
	if varstart<>"" then
		arrlstStart(0) = selectmonth("lstmonstart", month(varstart), 0)
		arrlstStart(1) = selectday("lstdaystart", day(varstart), 0)
		arrlstStart(2) = selectyear("lstyearstart", year(varstart), 1999, intTmp + 10, 0)
	else
		arrlstStart(0) = selectmonth("lstmonstart", month(now()), 0)
		arrlstStart(1) = selectday("lstdaystart", day(now()), 0)
		arrlstStart(2) = selectyear("lstyearstart", intTmp, 1999, intTmp + 10, 0)
	end if	
	if varleave<>"" then

		arrlstLeave(0) = selectmonth("lstmonleave", month(varleave), 1)
		arrlstLeave(1) = selectday("lstdayleave", day(varleave), 1)
		arrlstLeave(2) = selectyear("lstyearleave", year(varleave), 1999 , intTmp + 1, 1)

	else	
		arrlstLeave(0) = selectmonth("lstmonleave", 0, 1)
		arrlstLeave(1) = selectday("lstdayleave", 0, 1)
		arrlstLeave(2) = selectyear("lstyearleave", 0, 1999, intTmp + 10, 1)
	end if

	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		ret = objDb.runQuery("SELECT * FROM ATC_Countries WHERE fgActivate=1 ORDER BY CountryName ")
		strOut1 = ""
		strOut2 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut1 = "<select id='lstnationality' name='lstnationality' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' onFocus='CheckMode(this)'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intnat) then strSel=" selected " else strSel="" end if
			    strOut1 = strOut1 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut1 = strOut1 & "</select>"
		end if
		
		ret = objDb.runQuery("SELECT * FROM ATC_Countries WHERE fgActivate=1 ORDER BY Nationality ")
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut2 = "<select name='lstcountry' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' onFocus='CheckMode(this)'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
			    if objDb.rsElement(0)=int(intcountry) then strSel=" selected " else strSel="" end if
			    strOut2 = strOut2 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(3)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut2 = strOut2 & "</select>"
		end if
		
		ret = objDb.runQuery("SELECT * FROM ATC_JobTitle WHERE fgActivate=1 ORDER BY JobTitle")
		strOut3 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else		
			strOut3 = "<select name='lstjobtitle' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' onFocus='CheckMode(this)'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intjobtitle) then strSel=" selected " else strSel="" end if
			    strOut3 = strOut3 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut3 = strOut3 & "</select>"
		end if

		ret = objDb.runQuery("SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department")
		strOut4 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut4 = "<select name='lstdepartment' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' onFocus='CheckMode(this)'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intdepartment) then strSel=" selected" else strSel="" end if
			    strOut4 = strOut4 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut4 = strOut4 & "</select>"
		end if
		
		strQuery = "SELECT DISTINCT a.UserID, e.Firstname + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') as Fullname " &_
					"FROM ATC_UserGroup a LEFT JOIN ATC_Group b ON a.GroupID = b.GroupID " &_
					"LEFT JOIN ATC_Permissions c ON b.GroupID = c.GroupID " &_
					"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID " &_
					"LEFT JOIN ATC_PersonalInfo e ON a.UserID = e.PersonID " &_
					"WHERE d.Description = 'Receive Report' AND e.fgDelete = 0 ORDER BY Fullname"
		ret = objDb.runQuery(strQuery)

		strOut5 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut5 = "<select name='lstreportto' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' onFocus='CheckMode(this)'>"
			if intreport="" then strSel=" selected" else strSel="" end if
			strOut5 = strOut5 & "<option value=''" & strSel & ">None</option>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intreport) then strSel=" selected" else strSel="" end if
			    strOut5 = strOut5 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut5 = strOut5 & "</select>"
		end if

		strQuery = "SELECT * from ATC_WorkingHours ORDER BY Hours"
		ret = objDb.runQuery(strQuery)

		strOut6 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut6 = "<select name='lstworkinghour' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' " &_
						"onFocus='CheckModeforwork(this)'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intWorkingHour) then strSel=" selected" else strSel="" end if
				if strMode="EDIT" and strSel<>"" then
					strOut6 = "<span class='blue-normal'>" & showlabel(objDb.rsElement(2)) & "</span>"
					strOut6 = strOut6 & "<input type='hidden' name='lstworkinghour' value='" & intWorkingHour & "'>"
					exit do
				end if
			    strOut6 = strOut6 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(2)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			if Instr(strOut6, "lstworkinghour")>0 then
				strOut6 = strOut6 & "</select>"
			end if
			objDb.CloseRec
		end if
	else
		'error in connection
		gMessage = objDb.strMessage
	end if
	objDb.dbdisConnect
	set objDb = nothing
	
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

<style type="text/css">

fieldset{
    margin: 			0;  
	padding: 			0;
    display:            inline;
    border:             none;
    vertical-align:     top;
}
  
fieldset legend{
	padding-bottom: 	20px;
}

fieldset ol {
    padding: 			0; 
    margin:             0;
	list-style: 		none;
}

fieldset li {  
    padding:0;
	padding-bottom: 	0.5em;
}

fieldset li label 
{	
	width:			    80px;
	clear:				left;
	float:				left;
}

#submit ul
{
      list-style:none;
      text-align:center;
      padding:0;
}

#submit ul li
{    
    margin-left:1px;
    display:inline;
}

#submit ul li a
{
    padding-top:3px;
    display:inline-block;
    width:60px;
    height:22px;
    background-color:#8CA0D1;
    text-align:center;
    font-weight: bold;
    text-decoration:none;
}

#submit ul li a:hover
{
    background-color:#7791D1;
    color:white;
}

</style>

<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
var objNewWindow;
function CheckMode(field){
var varMode="<%=strMode%>";
var fggender=<%=int(vargender)%>;
    if (varMode!="EDIT"){
	  field.blur();
      if (field.type!="text") document.frmdetail.txtfirst.focus();
      if (field.type=="radio") {
		if(fggender==-1) document.frmdetail.optgender(0).checked = true;
		else document.frmdetail.optgender(1).checked = true;
	  }
    }
}

function CheckModeforwork(field){
var varMode = "<%=strMode%>";
var varid = "<%=struserID%>";
    if((varMode!="EDIT")||(varid!="")){
      document.frmdetail.txtfirst.focus();
    }
}

function checkday(m, d, y) {
	varM = "a04a06a09a11a";
	varmonth = eval("document.frmdetail." + m + ".options[document.frmdetail." + m + ".selectedIndex].value");
	varday = eval("document.frmdetail." + d + ".options[document.frmdetail." + d + ".selectedIndex].value");
	varyear = eval("document.frmdetail." + y + ".options[document.frmdetail." + y + ".selectedIndex].value");
	tmp = "a"+varmonth+"a";
	if ((varyear=="")&&(varmonth!="")&&(varday!="")) return false
	if ((varyear!="")&&(varmonth=="")&&(varday!="")) return false
	if ((varyear!="")&&(varmonth!="")&&(varday=="")) return false
	if ((varM.indexOf(tmp)!=-1)&&(varday>30)) return false;
	if (checkYear(varyear)==0)
	{
		if ((varmonth==2)&&(varday>28)) return false;}
	else
	{
		if ((varmonth==2)&&(varday>29)) return false;}
	return true;
}

function checkYear(year) 
{ 
	return (((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0)) ? 1 : 0;
}

function checkstartleave(m1, d1, y1, m2, d2, y2) {
	varmonth1 = eval("document.frmdetail." + m1 + ".options[document.frmdetail." + m1 + ".selectedIndex].value");
	varday1 = eval("document.frmdetail." + d1 + ".options[document.frmdetail." + d1 + ".selectedIndex].value");
	varyear1 = eval("document.frmdetail." + y1 + ".options[document.frmdetail." + y1 + ".selectedIndex].value");
	varmonth2 = eval("document.frmdetail." + m2 + ".options[document.frmdetail." + m2 + ".selectedIndex].value");
	varday2 = eval("document.frmdetail." + d2 + ".options[document.frmdetail." + d2 + ".selectedIndex].value");
	varyear2 = eval("document.frmdetail." + y2 + ".options[document.frmdetail." + y2 + ".selectedIndex].value");
	if((varmonth2!="")&&(varday2!="")&&(varyear2!="")) {
	  if (varyear1>varyear2) return false;
	  if (((varyear1==varyear2)&&(varmonth1>varmonth2))||((varyear1==varyear2)&&(varmonth1==varmonth2)&&(varday1>varday2))) return false;
	}
	return true;
}
	  
function checkdata() {
	var tmp1 = alltrim(document.frmdetail.txtsurname.value);
	document.frmdetail.txtsurname.value = tmp1;
	if(tmp1=="") {
		alert("Please enter the surname.");
		document.frmdetail.txtsurname.focus();
		return false;
	}
	tmp1 = alltrim(document.frmdetail.txtfirst.value);
	document.frmdetail.txtfirst.value = tmp1;
	if(tmp1=="") {
		alert("Please enter the first name.");
		document.frmdetail.txtfirst.focus();
		return false;
	}
	if(checkday("lstmonbirth", "lstdaybirth", "lstyearbirth")==false) {
		alert("Invalid date!");
		document.frmdetail.lstmonbirth.focus();
		return false;
	}
	if(checkday("lstmonstart", "lstdaystart", "lstyearstart")==false) {
		alert("Invalid date!");
		document.frmdetail.lstmonstart.focus();
		return false;
	}
	if(checkday("lstmonleave", "lstdayleave", "lstyearleave")==false) {
		alert("Invalid date!");
		document.frmdetail.lstdayleave.focus();
		return false;
	}
	else {
		if(checkstartleave("lstmonstart", "lstdaystart", "lstyearstart", "lstmonleave", "lstdayleave", "lstyearleave")==false) {
			alert("Leave date after than Join date!");
			document.frmdetail.lstdayleave.focus();
			return false;
		}
	}
	
	tmp1 = alltrim(document.frmdetail.txtusername.value);
	document.frmdetail.txtusername.value = tmp1;
	if(tmp1=="") {
		alert("Please enter the user name.");
		document.frmdetail.txtusername.focus();
		return false;
	}
	tmp1 = alltrim(document.frmdetail.txtemail.value);
	document.frmdetail.txtemail.value = tmp1;
	if (tmp1=="") 
	{
		alert("Please enter a value for email.");
		document.frmdetail.txtemail.focus();
		return false;
	}
	else
	{
	  if(!isemail(tmp1)) {
		alert("Invalid value email address \nValid format is: 'NickName@domain.com'");
		document.frmdetail.txtemail.focus();
		return false;
	  }
	}
	
	tmp1 = alltrim(document.frmdetail.txtExemail.value);
	document.frmdetail.txtExemail.value = tmp1;
	
	if (tmp1!="") 
	{
		if(!isemail(tmp1)) {
			alert("Invalid value external email address \nValid format is: 'itsupport@atlasindustries.com'");
			document.frmdetail.txtExemail.focus();
			return false;
		}
	}
	
	tmp1 = document.frmdetail.lstreportto.options[document.frmdetail.lstreportto.selectedIndex].value;
	if(tmp1==""){
		if(confirm("You have not yet selected a 'Reports to' value.\n Do you want to continue to input data?"))
			return false
	}
	return true;
}

function _act(kind) {
var varMode="<%=strMode%>";
var stract="<%=strAct%>";
	act=0;

	if(kind=="EDIT") {
		if(varMode!="EDIT") act=1;
		}
	else {
		if(kind=="DELETE"){
			if(stract!="ADD"){
			  if(confirm("Are you sure you want to delete this Employee?")) act=1;
			  }
			}
		else {
			if(kind=="SAVE") {
				if(checkdata()==true) act=1;
			}
			else act=1;
		}
	}
	if(act==1) {
		document.frmdetail.action = "employeedetail.asp?act=" + kind;
		document.frmdetail.target = "_self";
		document.frmdetail.submit();
	}
}

function addmore(vtype) { //v2.0
  window.status = "";
  if((vtype==1)||(vtype==4))
	strFeatures = "top="+(screen.height/2-78)+",left="+(screen.width/2-132)+",width=265,height=158,toolbar=no," 
	            + "menubar=no,location=no,directories=no,resizable=no";
  else
  	strFeatures = "top="+(screen.height/2-92)+",left="+(screen.width/2-132)+",width=265,height=184,toolbar=no," 
	            + "menubar=no,location=no,directories=no,resizable=no";
  if((objNewWindow) && (!objNewWindow.closed))
	objNewWindow.focus();	
  else {
	objNewWindow = window.open("addmore.asp?type=" + vtype, "MyNewWindow", strFeatures);
  }
  window.status = "Opened a new browser window.";  
}


function window_onunload() {
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.close();
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onUnload="return window_onunload();">
<form method="post" name="frmdetail">
<input type="hidden" name="txtleavedue" class="blue-normal" maxlength="20" style="width:160px" value="<%=showlabel(strleavedue)%>">
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
            <td style="padding:20px 20px 0 20px;"> 
              <%if gMessage<>"" then%>
               <div style="font-weight:bold; height:20px; background-color:#E7EBF5;" class="red"><%=gMessage%></div>
              <%end if%>
               <a class="blue" href="../../Management/staff/listofemployee.asp" onMouseOver="self.status='Show the list of employees'; return true;" onMouseOut="self.status=''">Employee List</a>
               <div class="title" style="padding:10px; text-align:center;">Employee Details</div>
            </td>
          </tr>
          
          <tr> 
            <td height="100%" valign="top" style="padding:0 0 0 20px">            
               <fieldset >
                    <legend class="blue">Personal Information</legend>
                    <ol>                                              
						<li>
                            <label for="lsttitle" class="blue-normal">Title</label>
                                <select class="blue-normal" size="1" id="lsttitle" name="lsttitle" onFocus="CheckMode(this)">
						              <option value="Mr" <%if ucase(trim(strtitleU))=ucase("Mr") then%>selected<%end if%>>Mr</option>
						              <option value="Mrs" <%if ucase(trim(strtitleU))=ucase("Mrs") then%>selected<%end if%>>Mrs</option>
						              <option value="Ms" <%if ucase(trim(strtitleU))=ucase("Ms") then%>selected<%end if%>>Ms.</option>
						            
					            </select></li>
					    <li>
					        <label for="txtsurname" class="blue-normal">Surname*</label>
					        <input type="text" id="txtsurname" name="txtsurname" maxlength="25" class="blue-normal" style="width:160px" value="<%=showlabel(strsurname)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
					    <li>
					        <label for="txtmiddle" class="blue-normal">Middle Name</label>
					        <input type="text" id="txtmiddle" name="txtmiddle" class="blue-normal" maxlength="15" style="width:160px" value="<%=showlabel(strmiddle)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                        
                        <li>
                            <label for="txtfirst" class="blue-normal">First Name*</label>                        
                            <input type="text" id="txtfirst" name="txtfirst" class="blue-normal" maxlength="20" style="width:160px" value="<%=showlabel(strfirst)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
					    <li>
					        <label class="blue-normal" for="optgender">Gender</label>
					         <%strChkm=""
                                     strChkf=""
                                     if vargender=true then strChkm = " checked" else strChkf = " checked"
                                     if strMode<>"EDIT" then
                                     strChkm = strChkm & " onClick='CheckMode(this);'"
                                     strChkf = strChkf & " onClick='CheckMode(this);'"
                                     end if%>
                                  <input type="radio" id="optMale" name="optgender" value="1" <%=strChkm%>><span class="blue-normal"> Male</span>
                                  <input type="radio" id="optFemale" name="optgender" value="0" <%=strChkf%>> <span class="blue-normal">Female</span></li>
                            <li><label class="blue-normal">Birthdate*</label>
                                    <%	Response.Write arrlstBirth(1)
	                                    Response.Write arrlstBirth(0)
	                                    Response.Write arrlstBirth(2)
                                    %> </li>
                            <li><label class="blue-normal" for="lstnationality">Nationality</label>
                                    <%Response.Write strOut2%></li>
                            <li><label class="blue-normal" for="txtaddress">Address</label>
                                <input type="text" id="txtaddress" name="txtaddress" class="blue-normal" maxlength="50" style="width:160px" value="<%=showlabel(straddress)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal" for="txtcity">City</label>
                            <input type="text" id="txtcity" name="txtcity" class="blue-normal" maxlength="20" style="width:160px" value="<%=showlabel(strcity)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal" for="txtstate">State</label>
                                <input type="text" id="txtstate" name="txtstate" class="blue-normal" maxlength="20" style="width:160px" value="<%=showlabel(strstate)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal" for="txtpostcode">Postal code</label>
                                <input type="text" id="txtpostcode" name="txtpostcode" class="blue-normal" maxlength="10" style="width:160px" value="<%=showlabel(strpostcode)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal">Country</label>
                                    <%Response.Write strOut1%></li>
                            <li><label class="blue-normal" for="txtemail">Local Email *</label>
                                <input type="text" id="txtemail" name="txtemail" class="blue-normal" maxlength="60" style="width:160px" value="<%=showlabel(stremail)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal" for="txtExemail">Ex. Email</label> 
                                <input type="text" id="txtExemail" name="txtExemail" class="blue-normal" maxlength="60" style="width:160px" value="<%=showlabel(strExemail)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal" for="txtphone">Phone</label>                         
                                <input type="text" id="txtphone" name="txtphone" class="blue-normal" maxlength="50" style="width:160px" value="<%=showlabel(strphone)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                            <li><label class="blue-normal" for="txtmobile">Mobile Phone</label>  
                                <input type="text" id="txtmobile" name="txtmobile" class="blue-normal" maxlength="50" style="width:160px" value="<%=showlabel(strmobile)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                    </ol>
               </fieldset>                
               <fieldset style="padding-left:50px;">
                    <legend class="blue">Working Information</legend>
                    <ol>
<%if struserID<>"" then %>                    
                        <li><label class="blue-normal">Person ID</label>
                            <span class="blue-normal"> <%Response.Write struserID%></span> </li> 
<%end if %>                            
                        <li><label class="blue-normal" for="txtidnum">Staff ID</label>
                            <input type="text" id="txtidnum" name="txtidnum" maxlength="15" class="blue-normal" style="width:160px" value="<%=showlabel(stridnum)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                        <li><label class="blue-normal" for="txtidnum">User Name*</label>
                            <input type="text" id="txtusername" name="txtusername" maxlength="20" class="blue-normal"  style="width:160px" value="<%=showlabel(strusernameU)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                    <%if strAct="ADD" then%>
                        <li><label class="blue-normal" for="txtpass">Password</label>
                            <input type="password" id="txtpass" name="txtpass" maxlength="15" class="blue-normal"  style="width:160px">
                    <%end if%></li>
                        <li><label class="blue-normal">Start Date*</label>
                            <%	Response.Write arrlstStart(1)
	                            Response.Write arrlstStart(0)
	                            Response.Write arrlstStart(2)%></li>
                        <li><label class="blue-normal">Last Date</label>
                            <%	Response.Write arrlstLeave(1)
	                            Response.Write arrlstLeave(0)
	                            Response.Write arrlstLeave(2)%></li>      
                        <li><label class="blue-normal">Job Title</label>
                            <%Response.Write strOut3%></li>
                        <li><label class="blue-normal">Department</label>
                            <%Response.Write strOut4%></li>
                        <li><label class="blue-normal">Reports To</label>
                                <%Response.Write strOut5%> </li>
                        <li><label class="blue-normal">Working Hours</label>
                                <%Response.Write strOut6%> </li> 
                        <li><label class="blue-normal" for="chkindirect">Indirect</label>                            
                                <input type="checkbox" id="chkindirect" name="chkindirect" value="1" <%if varIndirect=true then%>checked <%end if%> <%if strMode<>"EDIT" then%>onClick="return false;" <%end if%>></li>
                        <li><label class="blue-normal" for="txtTelExt">Telephone Ext.</label>
                            <input type="text" id="txtTelExt" name="txtTelExt" class="blue-normal" maxlength="20" style="width:160px" value="<%=showlabel(strTelExt)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>></li>
                        <%if strPhoto<>"" Then%>
						    <li style="padding-left:80px"><img src="../../../../staff/images/<%=strPhoto%>.jpg"></li>
                    <%end if%>
                    </ol>
                </fieldset>
<%if fgUpdate then%>                
               <div id="submit">
                    <ul>
                        <li><a href="javascript:_act('ADD');">Add</a></li>
                        <%if struserID<>"" then%><li><a href="javascript:_act('EDIT')">Edit</a></li> <%end if%>
                        <%if strMode<>"VIEW" then%><li><a href="javascript:_act('SAVE')">Save</a></li><%end if%>
                        <%if struserID<>"" then%><li><a href="javascript:_act('DELETE');">Delete</a></li><%end if%>
                    </ul> 

               </div>
<%end if %>               
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
<input type="hidden" name="txtuserid" value="<%=struserid%>">
<input type="hidden" name="txtnameold" value="<%=strusernameU%>">
<input type="hidden" name="txtmode" value="<%=strMode%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>