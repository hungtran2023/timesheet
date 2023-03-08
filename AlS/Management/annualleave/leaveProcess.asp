<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim intUserID, intMonth, intYear, dblCurLeave
	Dim strConnect, objDatabase, strError
	Dim rsDuration,rsIndividualRule,dblBalance,dblApplication,dblLeaveDue
	Dim dateTo
%>
<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="Styles.css">

</head>

<script language="javascript" src="../../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--

//-->
</script>

</head>
<body>
<div id="mainBody"></div>

<div id="navigation">
	<span class="title"></span>	
</div>
<div id="centerDoc">
	
</div>



</body>
</html>
