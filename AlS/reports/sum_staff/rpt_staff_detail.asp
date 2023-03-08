<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/getmenu.asp"-->


<%

	Dim ii, intRow, intCompanyID, intStaffID, intUserID, intMonth, intYear
	Dim strStaff, strProjectName, strView, strType, strSQL, strConnect, strFrom, strFrom1, strTo, strTo1, strTitle, strTitle2, strFunction
	Dim objDatabase, varUser, varFrom, varTo, varFullName,rsHours
	
Function ATSTable(strF,strT)
	dim strTable
	
	strTable=""
	For i=year(strF) to Year(strT)
		if strTable<>"" then strTable=strTable & "  UNION ALL "
		strTable=strTable & " SELECT AssignmentID, Hours, OverTime FROM ATC_Timesheet"
		if i<>year(date()) then strTable=strTable & i
	next
	
	ATSTable=strTable & " WHERE Tdate BETWEEN '" & cdate(strF) & "' AND '" & cdate(strT) & "'"
End Function

Function ViewDetail()
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	If objDatabase.dbConnect(strConnect) Then
		
		strSQL="SELECT b.ProjectID, ProjectName,Sum(Hours) as Hours, Sum(OverTime) as OT FROM " &_
				"(" & ATSTable(strFrom,strTo) & " AND StaffID=" & intStaffID & " AND AssignmentID<>1) ATS " & _
					"INNER JOIN ATC_Assignments a ON a.AssignmentID=ATS.AssignmentID " & _
					"INNER JOIN ATC_Tasks b ON a.SubtaskID=b.SubtaskID " & _
					"INNER JOIN ATC_Projects c ON c.ProjectID=b.projectID "
        Select Case cint(strView)
          Case 0
            strSQL=strSQL & " WHERE LEFT(b.ProjectID,3)<> 'ATL' AND fgBillable<>0"
          Case 1
            strSQL=strSQL & " WHERE LEFT(b.ProjectID,3)<> 'ATL' AND fgBillable=0"
          Case 2
            strSQL=strSQL & " WHERE LEFT(b.ProjectID,3)<> 'ATL'"
          Case 3
            strSQL=strSQL & " WHERE LEFT(b.ProjectID,3)='ATL' AND SUBSTRING(b.ProjectID,8,1)<>'Z'"
          Case 4
            strSQL=strSQL & " WHERE SUBSTRING(b.ProjectID,8,1)='Z'"
        End Select
					
									
		strSQL=strSQL & " GROUP BY b.ProjectID,ProjectName"

		If (objDatabase.runQuery(strSQL)) Then
			If objDatabase.noRecord = False Then
				set rsHours = objDatabase.rsElement
			End If
		Else
			ViewDetail = objDatabase.strMessage
		End If
	Else
		ViewDetail = objDatabase.strMessage
	End If
		
End Function

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------

	intStaffID = Request.Form("txthidden")
	strView	= Request.QueryString("t")
	
	strFrom=Request.Form("F") 
	strTo=Request.Form("T")

	intCompanyID = session("InHouse")

	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID	= session("USERID")

'--------------------------------------------------
' Analyse query and prepare report
'--------------------------------------------------

	intRow = -1
	If ViewDetail() <> "" Then
		Response.Write "<table width='780' border='0' cellspacing='0' cellpadding='0' align='center'><tr><td><font face='Arial' size='2' color='#FF0000'><b>" & GenerateReport & "</b></font></td></tr></table>"
	End If
		
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strFunction = "<a class='c' href='../../welcome.asp?choose_menu=B'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:window.history.back();'>Back</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:gopage();' onMouseOver='self.status=&quot;Preferences&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Preferences</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage();'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' title='Log Out'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "<b>" & varFullName(0) & " - " & varFullName(1) & "</b>"

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

	Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../../templates/template1/")

'--------------------------------------------------
' Free variables
'--------------------------------------------------	
	Set objDatabase = Nothing
	Set objEmployee = Nothing
	
%>

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/library.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	var url;
	url = "../../logout.asp";
	if (ns)
		document.location = url;
	else
	{
		window.document.frmreport.action = url;
		window.document.frmreport.target = "_self";
		window.document.frmreport.submit();
	}	
}

function gopage()
{
	document.frmreport.action = "../../tools/preferences.asp";
	document.frmreport.submit();
}

function printpage()
{
	window.print();
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
%>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6" background="../../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
    <td valign="top" height="100%" width="772">
      <table width="100%" border="0" cellspacing="1" cellpadding="0" align="center" style="height:79%" height="365">
<%If strError <> "" Then%>
		<tr>
          <td height="80"> 
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr bgcolor="#E7EBF5"> 
				<td class="blue" align="center"><%=strError%></td>
			  </tr>
			</table>
		  </td>	
		</tr>
<%End If%>			    
	    <tr> 
          <td height="80"> 
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr> 
				<td class="title" height="50" align="center"><%=strTitle2%></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<tr> 
		  <td valign="top"> 
			<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" bgcolor="#003399">
			  <tr bgcolor="#FFFFFF">
			    <td align="right" class="blue"><%=strTitle1%></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td align="right" class="blue">&nbsp;</td>
			  </tr>
			  <tr> 
           		<td bgcolor="#8FA4D3"> 
			      <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
                    <tr bgcolor="#617DC0" height="25"> 
                      <td class="white" align="center" width="35%"><b>Project ID</b></td>
                      <td class="white" align="center" width="35%"><b>Project Name</b></td>
                      <td class="white" align="center" width="15%"><b>Normal Hour</b></td>
                      <td class="white" align="center" width="15%"><b>OverTime Hour</b></td>
					</tr>
<%
		if not rsHours.EOF then
			
			do while not rsHours.EOF
%>					
					<tr bgcolor="#E7EBF5" height="20"> 
					  <td valign="middle" width="35%" class="blue-normal">&nbsp;&nbsp;<%=showlabel(rsHours("ProjectID"))%></td>
	                  <td valign="middle" width="35%" class="blue-normal">&nbsp;&nbsp;<%=showlabel(rsHours("ProjectName"))%></td>
		              <td valign="middle" width="15%" class="blue-normal" align="right"><%=Formatnumber(rsHours("Hours"), 2)%>&nbsp;&nbsp;</td>
  		              <td valign="middle" width="15%" class="blue-normal" align="right"><%=FormatNumber(rsHours("OT"), 2)%>&nbsp;&nbsp;</td>
			        </tr>
<%
				rsHours.MoveNext
				
			loop
		end if
%>			
       
				  </table>
				</td>
			  </tr>
			<tr bgcolor="#FFFFFF">
			    <td align="right" class="blue"><a href="rpt_sum_staff.asp?act=vpa1">Back to report</a></td>
			  </tr> 			  
			</table>
		  </td>
		</tr>
      </table>
    </td>
    <td width="2" background="../../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
  </tr>
</table>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>
</form>
</body>
</html>

