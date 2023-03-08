<!-- #include file = "../class/CDatabase.asp"-->
<!-- #include file = "../inc/library.asp"-->

<%
	Dim intAssignmentID, intMonth, intYear, intRow
	Dim varTimesheet
	
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	
	intMonth = Request.QueryString("m")
	intYear	 = Request.QueryString("y")
	
	intAssignmentID = Request.Form("assign")
	intRow = Request.Form("row")

'--------------------------------------------------
' End of initializing variables
'--------------------------------------------------
	
	If Not checkSession(session("varTimesheet")) Then
		Response.Redirect("../message.htm")
	End If					

	strError = tmsRemoveSubtask(intRow, intMonth, intYear)
%>

<form name="frmtemp" method="post">
	<input type="hidden" value="" name="xx">
</form>
<script language="javascript" src="../library/library.js"></script>
<script language="javascript">
	
	if ("<%=strError%>" == "")
		window.document.frmtemp.action = "timesheet.asp?act=vpa";
	else
		window.document.frmtemp.action = "timesheet.asp?act=vpae";
	
	window.document.frmtemp.submit();	
</script>		
