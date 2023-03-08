<%@ Language=VBScript ENABLESESSIONSTATE="False" %>
<%
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then
		strUsername = Request("username")
		strSQL = "SELECT UserName FROM ATC_Users c INNER JOIN ATC_Timesheet a ON c.UserID=a.StaffID INNER JOIN ATC_Assignments b ON a.AssignmentID = b.AssignmentID" & _
				" WHERE Tdate = '" & date & "' AND UserName='" & strUserName & "'"
		If (objDatabase.runQuery(strSQL)) Then
			If Not objDatabase.noRecord Then
				Response.AddHeader "Timesheet","Yes"
			Else	
				Response.AddHeader "Timesheet","No"
			End If	
		End If			
	End If
	Set objDatabase = Nothing
%>
<!-- #include file = "../class/CEmployee.asp"-->
