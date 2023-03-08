<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
--> 
<%
'--------------------------------------------------------------------------------
'check Invoive 
'--------------------------------------------------------------------------------
Function InvoieAlready(byval strProIDkey )
	dim blnReturn, strSql
	
	blnReturn=false
		
	strSql="SELECT COUNT(*) as numInv FROM ATC_ProjectInvoices WHERE ProjectID='"& strProIDkey(0)&_
			strProIDkey(1) &_
			strProIDkey(2) &_
			strProIDkey(3) &_
			strProIDkey(4) &_
			strProIDkey(5) &_
			strProIDkey(6) &"'"
	Call GetRecordset(strSql,rsProInv)
	
	InvoieAlready=cint(rsProInv("numInv"))>0
End Function
'--------------------------------------------------------------------------------
'Get projectKey in HTML
'--------------------------------------------------------------------------------
Function TableField(byval strValue)
	dim strOut
	if trim(strValue)="" then strValue="&nbsp;"
	strOut="<table width='100%' border='0' bordercolor='#333333'>"
	strOut=strOut & "<tr><td bgcolor='#333333'>"
	strOut=strOut & "<table width='100%' border='0' cellspacing='1'>"
    strOut=strOut & "<tr><td class='blue-normal' bgcolor='#FFFFFF'>" & strValue
    strOut=strOut & "</td></tr></table></td></tr></table>"
    TableField=strOut
End Function
'--------------------------------------------------------------------------------
'Get projectKey in HTML
'fgMode:
'      - <empty> : new project
'      - New     : Approval
'      - Issued  : Modify project
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'Get projectKey in HTML
'fgMode:
'      - <empty> : new project
'      - New     : Approval
'      - Issued  : Modify project
'--------------------------------------------------------------------------------
Function ProjectKeyHTMLForNew(byval fgMode,byval fgRightApproval,byval strProIDkey)
	
	dim strOut
	dim rsProType,rsCountry,rsSector,rsServiceType
	
	strOut=ReadTemplate("../../templates/template1/","main/ats_APK.htm")
	Call GetRecordset("SELECT ProjectType,Typevalue FROM ATC_ProjectTypes WHERE fgActivate=1 ORDER BY Typevalue",rsProType)
	Call GetRecordset("SELECT CountryName,CountryCode,Currency FROM ATC_Countries WHERE fgActivate=1 AND fgIsMarket=1 ORDER BY CountryCode",rsCountry)	
	'Call GetRecordset("SELECT SectorDescription,SectorValue FROM ATC_ProjectSector ORDER BY SectorValue",rsSector)
	
	Call GetRecordset("SELECT ServiceName,ServiceCode,fgActivate FROM ATC_ProjectServiceType ORDER BY ServiceCode",rsServiceType)
	Call GetRecordset("SELECT SectorDescription,SectorCode FROM ATC_ProjectSector ORDER BY SectorCode",rsSector)
	'For New project
	if fgMode = "" then
	
			'For client code
			strOut= Replace(strOut,"@@ClientCode","<input name='txtClientCode' size='10' onblur='checkclient()' style='WIDTH: 72px; HEIGHT: 20px' class='blue-normal' maxlength='3' value='" & trim(strProIDkey(0)) & "'>")
			'Project Number 
			strOut= Replace(strOut,"@@ProjectNumber",TableField(strProIDkey(1)))
			'@@Variation
			strOut= Replace(strOut,"@@Variation",TableField(strProIDkey(2)))	
			strOut= Replace(strOut,"@@ProjectType",PopulateDataToList("lstProType",rsProType,"Typevalue","ProjectType",strProIDkey(3)))	
			
			strOut= Replace(strOut,"@@Country",PopulateDataToList("lstCountry",rsCountry,"CountryCode","CountryName",strProIDkey(4)))			
			'strOut= Replace(strOut,"@@Sector",PopulateDataToList("lstSector",rsSector,"SectorValue","SectorDescription",strProIDkey(5)))
			'rsServiceType.Filter="fgActivate=1"
			
			strOut= Replace(strOut,"@@ServiceType",PopulateDataToList("lstServiceType",rsServiceType,"ServiceCode","ServiceName",strProIDkey(5)))		
			strOut= Replace(strOut,"@@Sector",PopulateDataToList("lstSector",rsSector,"SectorCode","SectorDescription",strProIDkey(6)))
			
	else
		'For Approval project
		'For client code
		strOut= Replace(strOut,"@@ClientCode",TableField(strProIDkey(0)))
		if fgMode = "New" then		
			'Project Number 
			strOut= Replace(strOut,"@@ProjectNumber","<input name='txtProjectNumber' size='10' style='WIDTH: 72px; HEIGHT: 20px' class='blue-normal' maxlength='4' value='" & trim(strProIDkey(1)) & "'>")
			'@@Variation
			strOut= Replace(strOut,"@@Variation","<input name='txtVariation' size='10' style='WIDTH: 72px; HEIGHT: 20px' class='blue-normal' maxlength='3' value=''>")
		else
		
			strOut= Replace(strOut,"@@ProjectNumber",TableField(strProIDkey(1)))
			strOut= Replace(strOut,"@@Variation",TableField(strProIDkey(2)))
			
			
		end if
		
		strOut= Replace(strOut,"@@ProjectType",PopulateDataToList("lstProType",rsProType,"Typevalue","ProjectType", strProIDkey(3)))
			
		if InvoieAlready(strProIDkey) then
			rsCountry.Filter = "CountryCode='" & strProIDkey(4) & "'"
			strOut= Replace(strOut,"@@Country",TableField(IIF(not rsCountry.EOF,rsCountry(1) & " - " & rsCountry(0),"")))
			rsServiceType.Filter="ServiceCode=" & strProIDkey(5) 			
			strOut= Replace(strOut,"@@ServiceType",TableField(IIF(not rsServiceType.EOF,rsServiceType(1) & " - " & rsServiceType(0),"")))		
			rsSector.Filter="SectorCode='" & strProIDkey(6) & "'"		
			strOut= Replace(strOut,"@@Sector",TableField(IIF(not rsSector.EOF,rsSector(1) & " - " & rsSector(0),"")))	
		else				
	
			strOut= Replace(strOut,"@@Country",PopulateDataToList("lstCountry",rsCountry,"CountryCode","CountryName",strProIDkey(4)))		
			strOut= Replace(strOut,"@@ServiceType",PopulateDataToList("lstServiceType",rsServiceType,"ServiceCode","ServiceName",cdbl(strProIDkey(5))))
			strOut= Replace(strOut,"@@Sector",PopulateDataToList("lstSector",rsSector,"SectorCode","SectorDescription",strProIDkey(6)))	
		end if
	end if
	
	ProjectKeyHTMLForNew=strOut
End Function


'--------------------------------------------------------------------------------
' PopulateDataToList
'--------------------------------------------------------------------------------
Function PopulateDataToList(byval strName, byval rs,byval strValueField, byval strDisplayField, byval strValue)
	Dim strOut
	strOut="<select name='" & strname & "' class='blue-normal' style='HEIGHT: 22px; WIDTH: 228px'><option value=''></option>"
	
	if not rs.Eof then
	  Do Until rs.EOF
		
		strOut = strOut & "<option value='" & rs(strValueField) & "'"		
		if rs(strValueField)=strValue then strOut = strOut & " selected "
		strOut = strOut & ">" & rs(strValueField) & " - " & showlabel(rs(strDisplayField)) & "</option>"
	    rs.MoveNext
	  Loop       
	end if
	
	PopulateDataToList=strOut & "</select>"
End function
'--------------------------------------------------------------------------------
' NavigationHTML
'--------------------------------------------------------------------------------
Function NavigationHTML(byval fgRightApproval,byval fgMode, byval fgRegistry, byval fgUpdate)
	dim strOut,strSave,strDelete,strAdd
	strOut=""
	'for Save button
	if (fgRegistry=true and fgMode="") or (fgUpdate=true and fgMode="Issued") or (fgRightApproval=true and fgMode="New") then
		strOut=strOut & "<td bgcolor='#8CA0D1' width='60' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;' height='20' valign='middle'>"
		strOut=strOut & "<a class='b' href='javascript:actpro(&quot;save&quot;);' onMouseOver='self.status=&quot;Save project&quot;;return true;' onMouseOut='self.status=&quot;&quot;;return true;'>Save</a></td>"
	end if
	'
	if (fgUpdate=true and fgMode="Issued") or (fgRightApproval=true and fgMode="New") then
		strOut=strOut & "<td bgcolor='#8CA0D1' width='60' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;'height='20' valign='middle'>"
		strOut=strOut & "<a class='b' href='javascript:actpro(&quot;del&quot;);' onMouseOver='self.status=&quot;Delete this project&quot;;return true;' onMouseOut='self.status=&quot;&quot;;return true;'><span>Delete<span></a></td>"
	end if
    if strOut<>"" then
		strOut="<table border='0' cellspacing='5' cellpadding='0' align='center' height='20' name='aa'><tr>" & strOut
		strOut=strOut & "</tr></table>" 		
	end if   
    NavigationHTML=strOut
End Function
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub GetInforProjectID(byval strHidden, byref strProjectID,byref strSattus, byref strDate,byref managerID)
	dim arrTemp	
	if strHidden<>"" then	
		arrTemp=split(strHidden,";")	
		strProjectID=arrTemp(0)
		strSattus=arrTemp(1)
		strDate=cDate(arrTemp(2))
		managerID=arrTemp(3)
	else
		'strProjectID="____"
		strSattus=""
		strDate=Date()
		managerID=0
	end if	
end sub
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub LoadProject(byval strProjectID,byval strProjectDate, byref strErr, byref rsPro)
	dim objDatabase
	dim strCnn,strSql
	dim strProKey
	
	strErr=""
	set rsPro=nothing	
	strProKey=split(strProjectID,"_")
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	

	If objDatabase.dbConnect(strConnect) Then
			Set rsPro = Server.CreateObject("ADODB.Recordset")
			Set rsPro.ActiveConnection = objDatabase.cnDatabase
			rsPro.CursorLocation = adUseClient			' Set the Cursor Location to Client
			
			strSql = "SELECT ATC_Projects.DepartmentID,DepartmentCode,ProjectKey2, NumCode as ProjectKey3, ATC_Projects.ProjectID, ProjectName, " & _
                      "DateTransfer, HourTransfer,ManagerID, CompanyName,[Value], [Description],CSOMainHours,clientCountryCode, " & _
                      "SignContract,DailyRate,ExchangeRate,EstRemaining,CurrencyCode,CSOGrossProfit,CSOFilename,isBIM, isDesign, SeverPath,Contingency, ATC_Projects.BDMID " & _
				"FROM ATC_Projects INNER JOIN ATC_ProjectStage ON ATC_Projects.ProjectID = ATC_ProjectStage.ProjectID " & _
					"INNER JOIN ATC_COMPANIES ON LEFT(ATC_Projects.PRojectID,3)=ATC_COMPANIES.CharCode " & _
					"INNER JOIN ATC_Department ON ATC_Projects.DepartmentID = ATC_Department.DepartmentID " & _
					"LEFT JOIN HR_BDM ON ATC_Projects.BDMID = HR_BDM.BDMID " & _
					"WHERE (ATC_Projects.ProjectID = '" & strProjectID & "') AND DateTransfer='" & strProjectDate & "'"		
			
			rsPro.Open strSQL
'Response.Write 	strSql		
			
			If Err.number =>0 then	
				strErr = Err.Description
			else
				set rsPro.ActiveConnection=nothing
			end if
	Else
		strErr = objDatabase.strMessage
	End If
	Set objDatabase = Nothing

End Sub
'--------------------------------------------------------------------------------
'Get company name
'--------------------------------------------------------------------------------
Function GetCompanyName(byval strCharCode,byref strPath, byref intID)	
	dim objDatabase
	dim strCnn,strSql
	dim strCompanyName
	
	strCompanyName=""
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	If objDatabase.dbConnect(strCnn) then
		strSQL = "SELECT CompanyID,CompanyName,SeverPath FROM ATC_Companies WHERE CharCode='" & replace(strCharCode,"'","''") & "'"
		if (objDatabase.runQuery(strSql)) then
			if not objDatabase.noRecord then
				strCompanyName=objDatabase.getColumn_by_name("CompanyName")
				intID=objDatabase.getColumn_by_name("CompanyID")
				strPath=objDatabase.getColumn_by_name("SeverPath")
			end if
		end if
	End if
	
	Set objDatabase = Nothing
	GetCompanyName=strCompanyName
End function
'--------------------------------------------------------------------------------
'Registry project ID
'--------------------------------------------------------------------------------
Function ProRegister(byval strKeys, byval strProjectName,byval dateTransfer,byval intDepartID ,byval intManagerID,byval strDescription,byval strUtilised,byval intOwnerID,byval blnisBIM,byval blnDesigned,byval dblBDM,byref strError)
	Dim myCmd,objDatabase,strCnn
	dim strProjectID

	strProjectID=strkeys(3)&strkeys(4)&strkeys(5)& strkeys(6)

	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strConnect) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "GetProjectID"
		
		Set myParam1 = myCmd.CreateParameter("ProjectID", adVarChar,adParamOutput,20)
		myCmd.Parameters.Append myParam1		
		Set myParam2 = myCmd.CreateParameter("clientCode", adVarChar,adParamInput,5)
		myCmd.Parameters.Append myParam2
		Set myParam3 = myCmd.CreateParameter("projectname", adVarChar,adParamInput,120)
		myCmd.Parameters.Append myParam3
		Set myParam4 = myCmd.CreateParameter("datetransfer", adVarChar,adParamInput,10)
		myCmd.Parameters.Append myParam4		
		Set myParam5 = myCmd.CreateParameter("departmentid",adInteger,adParamInput)
		myCmd.Parameters.Append myParam5
		Set myParam6 = myCmd.CreateParameter("managerid",adInteger,adParamInput)
		myCmd.Parameters.Append myParam6
		Set myParam7 = myCmd.CreateParameter("description",adVarChar,adParamInput,5000)
		myCmd.Parameters.Append myParam7
		Set myParam8 = myCmd.CreateParameter("staffid", adInteger,adParamInput)
		myCmd.Parameters.Append myParam8
		Set myParam9 = myCmd.CreateParameter("key2", adInteger,adParamInput)
		myCmd.Parameters.Append myParam9
		Set myParam10 = myCmd.CreateParameter("isBIM",adBoolean,adParamInput)
		myCmd.Parameters.Append myParam10
		Set myParam10 = myCmd.CreateParameter("isDesigned",adBoolean,adParamInput)
		myCmd.Parameters.Append myParam10
		
		Set myParam10 = myCmd.CreateParameter("BDMID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam10
		
		Set myParam11 = myCmd.CreateParameter("temPro", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam11

		myCmd("temPro")			= strProjectID
		myCmd("clientCode")		= strKeys(0)
		myCmd("departmentid")	= intDepartID
		myCmd("projectname")	= strProjectName
		myCmd("description")	= strDescription
		myCmd("managerid")		= intManagerID
		myCmd("StaffID")		= intOwnerID
		myCmd("datetransfer")	= dateTransfer
		myCmd("key2")			= strUtilised
		myCmd("isBIM")          = blnisBIM
		myCmd("isDesigned")          = blnDesigned
		myCmd("BDMID")          = dblBDM

		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			strError = ""
			ProRegister = myCmd("ProjectID")
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing
End Function
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub AddnewProject(byval strKeys, byval strProjectName,byval dateTransfer,byval intDepartID ,byval intManagerID,byval strDescription,byval strUtilised,byval intOwnerID,byval blnisBIM,byval blnDesigned,byval dblBDM,byref strError)
	
	dim ProjectID
	dim strErr
	dim objEmployee,varFullName

	ProjectID=ProRegister(strKeys, strProjectName,dateTransfer,intDepartID ,intManagerID,strDescription,strUtilised,intOwnerID,blnisBIM,blnDesigned,dblBDM,strError)

	if strErr<>"" then
		strError=strErr
	else
		Set objEmployee = New clsEmployee	
		
		objEmployee.SetFullName(intOwnerID)
		varFullName = split(objEmployee.GetFullName,";")
		'call SendEmailRequestApproval(ProjectID,varFullName(0),varFullName(3),strErr)
		
		Set objEmployee = nothing
	end if
	
End Sub
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub UpdateProject(byval strFullProjectID,byval strProjectIDkey,byval strOldDate,byval intDepartID,byval strProjectkey2,byval strProjectName,byval strDescription,byval intManagerID,_
	byval strDateTranfer,byval dblHourTransfer,byval fgApp, byval dblCSOMainHours,byval intSignContract,byval projectValue,byval dailyRate, byval exRate,byval estRemain,_
	 byval strClientCountryCode,byval strCurrency, byval strCSOFilename,byval blnisBIM,byval blnDesigned, byval dblContingency,byval dblBDM,_
	 byval dblCSOGrossProfit, byref strError)
	dim objDatabase
	dim strCnn,strSql
	dim strOld,strTemp,strReturn
	dim strNewID

	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strConnect) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "UpdateProject"
		
		
'@oldID nvarchar(20), @newID nvarchar(20), @strOldDate datetime, @intDepartID int, @strProjectkey2 int, @strProjectName nvarchar(50),
'@strDescription nvarchar(4000), @intManagerID int, @strDateTranfer datetime, @dblHourTransfer decimal(15,2), @blnCSOCompleted bit, 
'@dblCSOMainHours decimal(12,2), @strCSOComment nvarchar(200), @blnCSOApproval bit, @intSignContract int, @projectValue decimal(12,2), 
'@dailyRate decimal(18,2), @exRate decimal(18,2), @estRemain int	
		Set myParam1 = myCmd.CreateParameter("oldID", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam1		
		Set myParam2 = myCmd.CreateParameter("newID", adVarChar,adParamInput,20)
		myCmd.Parameters.Append myParam2
		Set myParam3 = myCmd.CreateParameter("strOldDate", adVarChar,adParamInput,10)
		myCmd.Parameters.Append myParam3
		Set myParam4 = myCmd.CreateParameter("intDepartID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam4		
		Set myParam5 = myCmd.CreateParameter("strProjectkey2",adInteger,adParamInput)
		myCmd.Parameters.Append myParam5
		Set myParam6 = myCmd.CreateParameter("strProjectName",adVarChar,adParamInput,120)
		myCmd.Parameters.Append myParam6
		Set myParam7 = myCmd.CreateParameter("strDescription",adVarChar,adParamInput,4000)
		myCmd.Parameters.Append myParam7
		Set myParam8 = myCmd.CreateParameter("intManagerID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam8
		Set myParam9 = myCmd.CreateParameter("strDateTranfer", adVarChar,adParamInput,10)
		myCmd.Parameters.Append myParam9		
		
		Set myParam10 = myCmd.CreateParameter("dblHourTransfer", adDecimal,adParamInput,17)
		myParam10.Precision=15
		myParam10.NumericScale=2
		myCmd.Parameters.Append myParam10
		
		Set myParam11 = myCmd.CreateParameter("dblCSOMainHours",adDecimal,adParamInput,14)
		myParam11.Precision=12
		myParam11.NumericScale=2		
		myCmd.Parameters.Append myParam11
		
		Set myParam15 = myCmd.CreateParameter("intSignContract", adInteger,adParamInput)
		myCmd.Parameters.Append myParam15
		
		Set myParam16 = myCmd.CreateParameter("projectValue",adDecimal,adParamInput,14)
		myParam16.Precision=12
		myParam16.NumericScale=2		
		myCmd.Parameters.Append myParam16
		
		Set myParam17 = myCmd.CreateParameter("dailyRate",adDecimal,adParamInput,20)
		myParam17.Precision=18
		myParam17.NumericScale=4		
		myCmd.Parameters.Append myParam17
		
		Set myParam18 = myCmd.CreateParameter("exRate",adDecimal,adParamInput,22)
		myParam18.Precision=18
		myParam18.NumericScale=4		
		myCmd.Parameters.Append myParam18
		Set myParam19 = myCmd.CreateParameter("estRemain", adInteger,adParamInput)
		myCmd.Parameters.Append myParam19
				
		Set myParam22 = myCmd.CreateParameter("strClientCountryCode", adVarChar,adParamInput,3)
		myCmd.Parameters.Append myParam22
		
		Set myParam23 = myCmd.CreateParameter("strCurrency", adVarChar,adParamInput,3)
		myCmd.Parameters.Append myParam23
			
		Set myParam13 = myCmd.CreateParameter("strCSOFilename",adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam13
		
		Set myParam21 = myCmd.CreateParameter("isBIM",adBoolean,adParamInput)
		myCmd.Parameters.Append myParam21
		
		Set myParam21 = myCmd.CreateParameter("isDesigned",adBoolean,adParamInput)
		myCmd.Parameters.Append myParam21
		Set myParam18 = myCmd.CreateParameter("Contingency",adDecimal,adParamInput,20)
		myParam18.Precision=18
		myParam18.NumericScale=2		
		myCmd.Parameters.Append myParam18

		Set myParam18 = myCmd.CreateParameter("CSOGrossProfit",adDecimal,adParamInput,20)
		myParam18.Precision=18
		myParam18.NumericScale=2		
		myCmd.Parameters.Append myParam18
		
		Set myParam8 = myCmd.CreateParameter("BDMID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam8
		
		myCmd("oldID")				=strFullProjectID
		myCmd("newID")				=strProjectIDkey(0) & strProjectIDkey(1) & strProjectIDkey(2) & strProjectIDkey(3)& strProjectIDkey(4)& strProjectIDkey(5)& strProjectIDkey(6)
		myCmd("strOldDate")			=strOldDate
		myCmd("intDepartID")		=intDepartID
		myCmd("strProjectkey2")		=strProjectkey2
		myCmd("strProjectName")		=strProjectName
		myCmd("strDescription")		=strDescription
	    
		myCmd("intManagerID")		=intManagerID
		myCmd("strDateTranfer")		=strDateTranfer
		myCmd("dblHourTransfer")	=dblHourTransfer
		myCmd("dblCSOMainHours")	=dblCSOMainHours
		
		myCmd("intSignContract")	=intSignContract
		myCmd("projectValue")		=projectValue
		myCmd("dailyRate")			=dailyRate
		myCmd("exRate")				=exRate
		myCmd("estRemain")			=estRemain
		myCmd("strClientCountryCode")=strClientCountryCode
		myCmd("strCurrency")		=strCurrency
		myCmd("strCSOFilename")		=strCSOFilename
		
        myCmd("isBIM")              = blnisBIM
        myCmd("isDesigned")         = blnDesigned
        
        myCmd("Contingency")		=dblContingency
        myCmd("CSOGrossProfit")		=dblCSOGrossProfit
		myCmd("BDMID")				= dblBDM
		myCmd.Execute
'response.write strProjectIDkey(0) & strProjectIDkey(1) & strProjectIDkey(2) & strProjectIDkey(3)& strProjectIDkey(4)& strProjectIDkey(5)& strProjectIDkey(6)
		If Err.number > 0 Then
			strError= Err.Description
		Else
			strError = ""
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	
	set objDatabase=nothing
	
End Sub
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Function CheckApproval(strProjectIDkey,byref strError)
	dim objDatabase
	dim strCnn,strSql
	dim strOld
	
	strError=""
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 		
	If objDatabase.dbConnect(strCnn) then		
		strSql="SELECT ProjectID FROM ATC_Projects WHERE ProjectID='" & strProjectIDkey(0) & strProjectIDkey(1) & strProjectIDkey(2) & strProjectIDkey(3)& strProjectIDkey(4)& strProjectIDkey(5) & "'"
		if objDatabase.runQuery(strSql) then
			if not objDatabase.noRecord then 
				strError="Project '" & strProjectIDkey(0) & strProjectIDkey(1) & strProjectIDkey(2) & strProjectIDkey(3)& strProjectIDkey(4)& strProjectIDkey(5) & "' currently exists. Please check and re-input it."
			end if
		else
			strError=objDatabase.strMessage
		end if
	else
		strError=objDatabase.strMessage
	end if
	Set objDatabase = nothing
	CheckApproval=(strError="")
End function

'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub ApprovalProject(byval strFullProjectID,byval strProjectIDkey,byval strProjectName,byval strDescription,byval intManager,byval strDateTranfer,byval strUtilised,byval blnisBIM,byval blnisDesigned,byval dblBDM, byref strError)
	dim objDatabase
	dim strCnn,strSql
	dim strOld,strTemp,strReturn

	
	'strTemp=split(strFullProjectID,"_")
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 		
	if CheckApproval(strProjectIDkey,strError) then
		If objDatabase.dbConnect(strCnn) then		
			Set myCmd = Server.CreateObject("ADODB.Command")
			Set myCmd.ActiveConnection = objDatabase.cnDatabase
			myCmd.CommandType = adCmdStoredProc
			myCmd.CommandText = "ApprovalProject"
			
			set myParam2 =  myCmd.CreateParameter("newID", adVarChar,adParamInput,20)
			myCmd.Parameters.Append myParam2
			set myParam3 =  myCmd.CreateParameter("projectname", adVarChar,adParamInput,120)
			myCmd.Parameters.Append myParam3
			set myParam4 =  myCmd.CreateParameter("description", adVarChar,adParamInput,5000)
			myCmd.Parameters.Append myParam4
			set myParam8 =  myCmd.CreateParameter("managerID", adInteger,adParamInput)
			myCmd.Parameters.Append myParam8
			Set myParam5 = myCmd.CreateParameter("dateTransfer", adVarChar,adParamInput,10)
			myCmd.Parameters.Append myParam5				
			Set myParam6 = myCmd.CreateParameter("oldID", adVarChar,adParamInput,20)
			myCmd.Parameters.Append myParam6
			Set myParam7 = myCmd.CreateParameter("key2", adInteger,adParamInput)
			myCmd.Parameters.Append myParam7
			set myParam9 = myCmd.CreateParameter("isBIM", adBoolean,adParamInput)
			myCmd.Parameters.Append myParam9
		    set myParam9 = myCmd.CreateParameter("isDesigned", adBoolean,adParamInput)
			myCmd.Parameters.Append myParam9
			set myParam8 =  myCmd.CreateParameter("BDMID", adInteger,adParamInput)
			myCmd.Parameters.Append myParam8
			Set myParam8 = myCmd.CreateParameter("Return", adVarChar,adParamOutput,100)
			myCmd.Parameters.Append myParam8
	
			myCmd("newID")= strProjectIDkey(0) & strProjectIDkey(1) & strProjectIDkey(2) & strProjectIDkey(3)& strProjectIDkey(4)& strProjectIDkey(5)& strProjectIDkey(6)
			myCmd("projectname")=strProjectName
			myCmd("description")=strDescription
			myCmd("dateTransfer")=strDateTranfer
			myCmd("managerID")=cint(intManager)
			myCmd("oldID")=strFullProjectID
			myCmd("key2")= strUtilised
			myCmd("isBIM")= blnisBIM
			myCmd("isDesigned")= blnisDesigned
			myCmd("BDMID")= dblBDM
			
			myCmd.Execute
					
			If Err.number > 0 Then
				strError= Err.Description
			Else
				strError = ""
				strReturn=split(myCmd("Return"),"#")
				call SendEmailActivateProject(strProjectIDkey(0) & "_" & strProjectIDkey(1) & "_" & strProjectIDkey(2),strProjectName,strDescription,strReturn(0), strReturn(1))
			End If
			Err.Clear
		
			set myCmd=nothing
		else
			strError=objDatabase.strMessage
		end if
	end if
	Set objDatabase = nothing	
End Sub
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Function DeleteProject(byval fullProjectID)
	dim objDatabase
	dim strCnn,strSql
	dim arrProKey,blnReturn
	
	blnReturn=false
	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	strSql="DELETE FROM ATC_ProjectStage WHERE ProjectID='" &_
			fullProjectID & "'"
	If objDatabase.dbConnect(strCnn) then		
		blnReturn= (objDatabase.runActionQuery(strSql))	
	else
		strError=objDatabase.strMessage
	end if
	
	Set objDatabase = nothing
	DeleteProject=blnReturn
End function
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub SendEmailRequestApproval(byval ProjectID, byval strOwnerName, byval strEmailAdd, byref strErr)
	dim objDatabase
	dim strCnn,strSql
	dim strContent,strSubject,rsApproval
	Dim MyCDONTSMail
	
	strErr=""
	
	strContent=ReadTemplate("../../templates/template1/","main/emailApproval.inc")
	if instr(1,strOwnerName,"(")>0 then 
		strContent=Replace(strContent,"@@ownerProject",left(strOwnerName,instr(1,strOwnerName,"(")-2))
	else
		strContent=Replace(strContent,"@@ownerProject",strOwnerName)
	end if
	strContent=Replace(strContent,"@@Projectkey",right(ProjectID,len(ProjectID)))
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	If objDatabase.dbConnect(strCnn) then
		strSql = "SELECT Fullname,EmailAddress_Ex as EmailAddress" &_
				 "FROM ATC_Users INNER JOIN ATC_UserGroup ON ATC_Users.UserID = ATC_UserGroup.UserID " &_
					"INNER JOIN ATC_Group ON ATC_UserGroup.GroupID = ATC_Group.GroupID " &_
					"INNER JOIN ATC_Permissions ON ATC_Group.GroupID = ATC_Permissions.GroupID " &_
					"INNER JOIN ATC_Functions ON ATC_Permissions.FunctionID = ATC_Functions.FunctionID " &_
					"INNER JOIN HR_Employee ON ATC_Users.UserID=HR_Employee.PersonID " &_
				 "WHERE (ATC_Functions.Description = 'approving project') " & _
				 "GROUP BY FirstName, MiddleName, LastName, EmailAddress"
		if (objDatabase.runQuery(strSql)) then
			if not objDatabase.noRecord then
				objDatabase.rsElement.MoveFirst
				
				do while not objDatabase.rsElement.EOF
					Set cdoMessage = CreateObject("CDO.Message")  
					With cdoMessage 
						Set .Configuration = getCDOConfiguration()  
						.From = strEmailAdd
						.To =  objDatabase.rsElement("EmailAddress")  
						'.Bcc="uyenchi.nguyentai@atlasindustries.com"
						.Subject = ucase(trim(replace(ProjectID,"_"," "))) & " - Approval required"
						if instr(1,objDatabase.rsElement("Fullname"),"(")>0 then
							.TextBody = replace(strContent,"@@Name",left(objDatabase.rsElement("Fullname"),instr(1,objDatabase.rsElement("Fullname"),"(")-2))
						else
							.TextBody = replace(strContent,"@@Name",objDatabase.rsElement("Fullname"))
						end if
						.Send 
					End With

					Set cdoMessage = Nothing  
					Set cdoConfig = Nothing 
					objDatabase.rsElement.MoveNext
				loop				
			end if
		else
			strErr=objDatabase.strMessage
		end if
	else
		strErr=objDatabase.strMessage
	End if
	
	Set objDatabase = Nothing
End Sub

'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub SendEmailActivateProject(byval ProjectID, byval projectname, byval Description,byval strOwnerName, byval strOwnerEmail)

	dim strContent,strSubject,varFullName
	Dim MyCDONTSMail,objEmployee
	
	strErr=""
	
	Set objEmployee = New clsEmployee	
		
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	
	strContent=ReadTemplate("../../templates/template1/","main/emailActivated.inc")
	if instr(1,strOwnerName,"(")>0 then 
		strContent=Replace(strContent,"@@Name",left(strOwnerName,instr(1,strOwnerName,"(")-2))
	else
		strContent=Replace(strContent,"@@Name",strOwnerName)
	end if
	
	strContent=Replace(strContent,"@@Projectkey",ProjectID)
	strContent=Replace(strContent,"@@Projectname",projectname)
	strContent=Replace(strContent,"@@Projectdesc",Description)
	
	Set cdoMessage = CreateObject("CDO.Message")  
	With cdoMessage 
		Set .Configuration = getCDOConfiguration()  
		.From = varFullName(3)
		.To =  strOwnerEmail

		.Subject = ucase(trim(replace(ProjectID,"_"," "))) & " - Timesheet project activation confirmation"
		if instr(1,varFullName(0),"(")>0 then
			.TextBody = replace(strContent,"@@Approval",left(varFullName(0),instr(1,varFullName(0),"(")-2))
		else
			.TextBody = replace(strContent,"@@Approval",varFullName(0))
		end if
		.Send 
	End With

	Set cdoMessage = Nothing  
	Set cdoConfig = Nothing 

End Sub
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Sub SendEmailForManager(byval strProjectIDkey, byval projectname, byval managerID)

	dim strContent,strSubject,varSendFullName,varReceiFullName
	Dim MyCDONTSMail,objEmployee,ProjectID
	
	ProjectID=strProjectIDkey(0) & strProjectIDkey(1) & strProjectIDkey(2) & strProjectIDkey(3)& strProjectIDkey(4)& strProjectIDkey(5)
	
	strErr=""
	
	Set objEmployee = New clsEmployee	
		
	objEmployee.SetFullName(session("USERID"))
	varSendFullName = split(objEmployee.GetFullName,";")
	
	
	objEmployee.SetFullName(managerID)
	varReceiFullName = split(objEmployee.GetFullName,";")	
	
	strContent=ReadTemplate("../../templates/template1/","main/emailAssignment.inc")
	if instr(1,varReceiFullName(0),"(")>0 then 
		strContent=Replace(strContent,"@@name",left(varReceiFullName(0),instr(1,varReceiFullName(0),"(")-2))
	else
		strContent=Replace(strContent,"@@name",varReceiFullName(0))
	end if
	
	strContent=Replace(strContent,"@@ProjectID",ProjectID)
	strContent=Replace(strContent,"@@Projectname",projectname)
	
	if instr(1,varSendFullName(0),"(")>0 then 
		strContent=Replace(strContent,"@@sender",left(varSendFullName(0),instr(1,varSendFullName(0),"(")-2))
	else
		strContent=Replace(strContent,"@@sender",varSendFullName(0))
	end if	

	Set cdoMessage = CreateObject("CDO.Message")  
	With cdoMessage 
		Set .Configuration = getCDOConfiguration()  
		.From = varSendFullName(3)
		.To =  varReceiFullName(3) 
		.Bcc=varSendFullName(3)
		.Subject = ucase(trim(replace(ProjectID,"_"," "))) & " - Timesheet project assignment"
		.TextBody = strContent
		.Send 
	End With

	Set cdoMessage = Nothing  
	Set cdoConfig = Nothing 
		
End Sub
'--------------------------------------------------------------------------------
'The project won't be deleted if there are any hours in Timesheet table.
'--------------------------------------------------------------------------------
Function beAbleToDelete(byval strfullID, byval dateTransfer)
	dim objDatabase
	dim strCnn,strSqlTemp,strSql
	dim arrProKey,i
	dim blnReturn
	blnReturn=true
	arrProKey=Split(strfullID,"_")
	
	strSqlTemp="SELECT * FROM @@ATC_Table WHERE AssignmentID in " & _
				"(SELECT AssignmentID FROM ATC_ProjectStage a INNER JOIN ATC_Tasks b ON a.ProjectID=b.ProjectID " & _
							"INNER JOIN ATC_Assignments c ON b.SubTaskID=c.SubTaskID " & _
							"WHERE a.ProjectID='" & strfullID & "') " & _
				"AND TDate>=(SELECT DateTransfer FROM ATC_ProjectStage WHERE ProjectID='" & _
						strfullID & "')"

	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	If objDatabase.dbConnect(strCnn) then
		for i=year(dateTransfer) to year(Date)
			if i<year(Date) then
				if i<=1999 then i=2000
				strSql=Replace(strSqlTemp,"@@ATC_Table","ATC_Timesheet" & i)
			else
				strSql=Replace(strSqlTemp,"@@ATC_Table","ATC_Timesheet")
			end if
			if (objDatabase.runQuery(strSql)) then
				if not objDatabase.noRecord then
					blnReturn=false
					exit for
				end if
			end if
		next
	End if
	Set objDatabase = Nothing
	beAbleToDelete=blnReturn
End Function

'--------------------------------------------------------------------------------
'The project won't be deleted if there are any hours in Timesheet table.
'--------------------------------------------------------------------------------
Function beAbleToUpdateDateTransfer(byval strfullID, byval dateTransfer,byval dateTransferNew)
	dim objDatabase
	dim strCnn,strSqlTemp,strSql
	dim arrProKey,i
	dim blnReturn
	blnReturn=true
	arrProKey=Split(strfullID,"_")

	if cdate(dateTransferNew)>cdate(dateTransfer) then
		strSqlTemp="SELECT * FROM @@ATC_Table WHERE AssignmentID in " & _
					"(SELECT AssignmentID FROM ATC_ProjectStage a INNER JOIN ATC_Tasks b ON a.ProjectID=b.ProjectID " & _
								"INNER JOIN ATC_Assignments c ON b.SubTaskID=c.SubTaskID " & _
								"WHERE a.ProjectID='" & arrProKey(2) & "_" & arrProKey(3) & "_" & arrProKey(4) & "') " & _
					"AND TDate>=CONVERT(datetime,'" & year(dateTransfer) & "-" & month(dateTransfer) & "-" & day(dateTransfer) & "',101) " & _
					"AND TDate<=CONVERT(datetime,'" & year(dateTransferNew) & "-" & month(dateTransferNew) & "-" & day(dateTransferNew) & "',101)"
					
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	If objDatabase.dbConnect(strCnn) then
		for i=year(dateTransfer) to year(dateTransferNew)
			if i<year(Date) then
				if i<=1999 then i=2000
				strSql=Replace(strSqlTemp,"@@ATC_Table","ATC_Timesheet" & i)
			else
				strSql=Replace(strSqlTemp,"@@ATC_Table","ATC_Timesheet")
			end if
			if (objDatabase.runQuery(strSql)) then
				if not objDatabase.noRecord then
					blnReturn=false
					exit for
				end if
			end if
		next
	End if
	end if
	beAbleToUpdateDateTransfer=blnReturn
End Function

%>
