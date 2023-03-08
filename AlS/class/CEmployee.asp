<!-- #include file = "CDatabase.asp"-->

<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

'**************************************************
' Copyright (C) by Atlas Industries Limited
' E-mail: info@atlasindustries.com
'
' CLASS NAME:
'
'		clsEmployee
'
' DESCRIPTION:
'
'
' METHODS:
'
'   Function GetFullName()
'   Sub SetFullName(ByVal intUserID)
'	Function GetDetailedInfo()
'	Sub SetDetailedInfo(ByVal intUserID)
'
' AUTHOR:
' DATE:
'
' NOTE:
'**************************************************

Class clsEmployee
	Private strFullName, strJobTitle, strSQL, strConnect, objDatabase
	Private dblOffHour, strJoinDate, strSalDate, dblSalary, strCurrency, fgOvertime,intCountryID
	Public strError

'**************************************************
' Sub: SetFullName
' Description: set username's fullname and jobtitle
' Parameters: - intUserID: integer
' Return value: None
' Author: 
' Date: 25/06/2001
' Note:
'**************************************************
	
	Public Sub SetFullName(ByVal intUserID)
		dim strEmailAddress
		
		strConnect = Application("g_strConnect")	
			
' Connect to SQL database 
		Set objDatabase = New clsDatabase 

		If objDatabase.dbConnect(strConnect) Then   
			strSQL = "SELECT (isnull(a.FirstName,'') + ' ' + isnull(a.LastName,'') + ' ' + isnull(a.MiddleName,'')) AS FullName, d.JobTitle, c.DepartmentID, a.EmailAddress, a.CountryID " & _
						" FROM ATC_PersonalInfo a INNER JOIN ATC_Users b ON a.PersonID=b.UserID INNER JOIN ATC_Employees c ON a.PersonID=c.StaffID LEFT JOIN ATC_JobTitle d ON c.JobTitleID=d.JobTitleID" & _ 
						" WHERE b.UserID=" & intUserID
			strSQL="SELECT Fullname, JobTitle,DepartmentID, EmailAddress_Ex as EmailAddress, NationalityID as CountryID  from HR_Employee WHERE PersonID=" & intUserID
			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					strFullName = objDatabase.getColumn_by_name("FullName")
					strJobTitle = objDatabase.getColumn_by_name("JobTitle")
					intDepartmentID_ = objDatabase.getColumn_by_name("DepartmentID")
					strEmailAddress=objDatabase.getColumn_by_name("EmailAddress")
					
					intCountryID=objDatabase.getColumn_by_name("CountryID")
				End If
			Else
				strError = objDatabase.strMessage
			End If
		Else
			strError = objDatabase.strMessage
		End If
															
		strFullName = strFullName & ";" & strJobTitle & ";" & intDepartmentID_ & ";" & strEmailAddress & ";" & intCountryID
		objDatabase.dbDisConnect()
	End Sub
		
'**************************************************
' Sub: GetFullName
' Description: Get username's fullname and jobtitle
' Parameters: None
' Return value: Username's fullName and Jobtitle
' Author: 
' Date: 25/06/2001
' Note:
'**************************************************

	Public Function GetFullName()
		GetFullName = strFullName
	End Function
	
	
'**************************************************
' Sub: GetCountryID
' Description: Get CountryID
' Parameters: None
' Return value: CountryID
' Author: 
' Date: 6/01/2010
' Note:
'**************************************************

	Public Function GetCountryID()
		GetCountryID = intCountryID
	End Function	
	
'**************************************************
' Sub: SetDetailedInfo
' Description: set username's detailed information such as Office Hour, JoinDate, etc..
' Parameters: - strUserName: String 
' Return value: None
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************

	Public Sub SetDetailedInfo(ByVal intUserID, ByVal strCheckDate)
	
		strConnect = Application("g_strConnect")	
			
' Connect to SQL database 
		Set objDatabase = New clsDatabase 

		If objDatabase.dbConnect(strConnect) Then   
			strSQL = "SELECT Hours, JoinDate, Salary, CurrencyCode, OverTimePay, SalaryDate  FROM ATC_WorkingHours a INNER JOIN ATC_SalaryStatus b ON a.WorkingHourID=b.WorkingHourID INNER JOIN ATC_Employees c ON b.StaffID=c.StaffID" & _
							" WHERE c.StaffID=" & intUserID & " AND b.SalaryDate IN (SELECT max(SalaryDate) as SalaryDate FROM ATC_SalaryStatus d WHERE b.StaffID=d.StaffID AND SalaryDate <= '" & strCheckDate & "' GROUP BY d.StaffID)"
		
			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					dblOffHour = objDatabase.getColumn_by_name("Hours")
					strJoinDate = objDatabase.getColumn_by_name("JoinDate")
					strSalDate = objDatabase.getColumn_by_name("SalaryDate")
					dblSalary = objDatabase.getColumn_by_name("Salary")
					strCurrency = objDatabase.getColumn_by_name("CurrencyCode")
					fgOvertime = objDatabase.getColumn_by_name("OverTimePay")
				End If
			Else
				strError = objDatabase.strMessage
			End If
		Else
			strError = objDatabase.strMessage
		End If
															
		strJoinDate = strJoinDate & "/" & strSalDate & "/" & strCurrency & "/" & dblOffHour & "/" & dblSalary & "/" & fgOvertime

		objDatabase.dbDisConnect()
		
	End Sub
	
'**************************************************
' Sub: GetDetailedInfo
' Description: Get username's Joindate, Salary, etc...
' Parameters: None
' Return value: Username's fullName and Jobtitle
' Author: 
' Date: 28/06/2001
' Note:
'**************************************************

	Public Function GetDetailedInfo()
		GetDetailedInfo = strJoinDate
	End Function
	
End Class

</SCRIPT>