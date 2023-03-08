<!-- #include file = "../../../class/CDatabase.asp"-->
<!-- #include file = "../../../inc/library.asp"-->

<%
	Response.Buffer = True

	If Request.Form("M") = "" Then
		intMonth = Month(Date)
	Else	
		intMonth = Request.Form("M")
	End If
	
	If Request.Form("Y") = "" Then
		intYear = Year(Date)
	Else	
		intYear = Request.Form("Y")
	End If		
	
	intDayNum	= GetDay(intMonth,intYear)				' Numbers of days in a month
	intStaffID  = Request.Form("txthidden")
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	strCheckDate = CDate(intMonth & "/" & intDayNum & "/" & intYear)
	
	If objDatabase.dbConnect(strConnect) Then			
		strSQL = "SELECT MAX(SalaryDate) AS SalaryDate FROM ATC_SalaryStatus WHERE StaffID=" & intStaffID & " AND SalaryDate <= '" & strCheckDate & "'"
		If (objDatabase.runQuery(strSQL)) Then
			If Not objDatabase.noRecord Then
				If Month(objDatabase.getColumn_by_name("SalaryDate")) = CInt(intMonth) And Year(objDatabase.getColumn_by_name("SalaryDate")) = CInt(intYear) Then
					If Day(objDatabase.getColumn_by_name("SalaryDate")) <> 1 Then
						blnSalType = True
						strCheckDate = objDatabase.getColumn_by_name("SalaryDate")
					Else
						blnSalType = False	
					End If	
				Else
					blnSalType = False	
				End If
			End If
		End If
	End If	
	Set objDatabase = Nothing

%>

<html>
<body>
<form name="frmtemp" method="post">

<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="checkdate" value="<%=strCheckDate%>">
<input type="hidden" name="txthidden" value="<%=intStaffID%>">

</form>
<%
	If blnSalType Then
		strURL = "sal_staff_tms1.asp?act=" & Request.QueryString("act") 
	Else
		strURL = "sal_staff_tms.asp?act=" & Request.QueryString("act") 
	End If
%>

<script language="javascript">
<!--
	window.document.frmtemp.action = "<%=strURL%>";
	window.document.frmtemp.target = "_self";
	window.document.frmtemp.submit();
//-->
</script>

</body>
</html>