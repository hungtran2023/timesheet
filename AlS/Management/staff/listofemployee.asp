<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strError, PageSize, fgRight 'view all or Not
	Dim intApproval

'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
	
'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 
	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
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

'-----------------------------------
'Check VIEWALL right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = "view all" then
				fgRight=true
				exit for
			end if
		next
		set getRight = nothing
	end if

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
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
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
	If strChoseMenu = "" Then strChoseMenu = "AB"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">
<head>
<meta http-equiv="Content-type" content="text/html;charset=UTF-8"/>
<meta http-equiv="Content-Language" content="en"/>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Atlas Industries Timesheet System</title>

<link href="../../bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../bootstrap/css/dataTables.bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../css/timesheet.css" rel="stylesheet" >
<link href="../../css/style.css" rel="stylesheet" type="text/css">    

</head>
<body data-pinterest-extension-installed="cr1.39.1">

<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
<div class="container-fluid" >
<%If strError <> "" Then%>  
	<div class="row">	
			<div class="<%if strError="Update successfull." then %>alert alert-danger<%else%>alert alert-success<%end if%>">
				<strong>Error:</strong><%=strError%>
			</div>
		</div>
<% End If%>	
	<div class="row" style="padding:20px 0px 0px 20px;"><h3>List of Employees</h3> </div>
	<div class="row">
        <div class="col-sm-6 col-sm-offset-3">
            <div id="imaginary_container"> 
			<form name="searchform" method="post" action="tms_list_staff.asp">
				<div class="form-group">
					<div class="input-group">
						<input type="text" name="txtSearch" id="txtSearch" onkeyup="myFunction()" class="form-control" placeholder="Filter">
						<div class="input-group-btn">
							<button type="button" id="btnFilter" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
								<span id="filterLable">By Fullname</span>
								<span class="caret"></span>
							</button>
							<ul class="dropdown-menu">
								<li><a  href="#" id="filterBy">By Manager</a></li>
							</ul>
						</div>
					</div>
				</div>
			</form>
            </div>
        </div>
	</div>
	<div class="row">
		<form id="frmList" method="post" action="employeeProfile.asp">
			<div class="table-responsive">	
				<div class="form-group" style="padding-left:15px">	
					<button class="btn  btn-default btnNext" id="btnNew" type="button">Add New Employee</button>
				</div>
				<table class="table table-hover" id="tblList">
					<thead  class="thead-inverse tableheaderblue">
						<tr>							
							<th>Full Name</th>
							<th>Birthday</th>
							<th>Start Date</th>
							<th>Jobtitle</th>
							<th>Department</th>
							<th>Report To</th>		
							<th>StaffID</th>							
						
						</tr>
					</thead>					
					
				</table>	
				<input type="hidden" name="txtuserid" id="txtuserid" value=""/>
                <input type="hidden" name="txtpreviouspage" value="<%=strFilename%>"/>  
			</div>		
		</form>		
    </div> 
	

</div>
        
      
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

<script type="text/javascript" src="../../js/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/dataTables.bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>
<script src="../../bootstrap/js/bootstrap.min.js"></script>

<script language="javascript">
<!--

$.extend( true, $.fn.dataTable.defaults, {
     "paging": false,
		"searching": false,
		"processing": true,
		"ordering": false,
		"info":     false
} );

$(document).ready(function() {
    $('#tblList').DataTable( {       
        "serverSide": true,
        "ajax": "../../inc/JSON_listEmployeeDetails.asp",
        "columns": [
			{ "data": "Fullname" },
			{ "data": "Birthday" },
            { "data": "StartDate" },
            { "data": "JobTitle" },
            { "data": "Department" },
            { "data": "ReportTo" },
			{ "data": "StaffID" }
        ]
    });	
	

	$('#tblList tbody').on('click', 'tr td:not(:last-child) ', function (e) {

		e.preventDefault();
		var table = $('#tblList').DataTable();
        var dataid = table.row(this).id();
		
		$("#txtuserid").val(dataid);		
		$("#frmList" ).submit(); 
    });
	
	$("#btnNew").click(function() {
        $("#txtuserid").val('-1');		
		$("#frmList" ).submit();        
    });
	
	
	$('#filterBy').click(function(e){ 
		var cur;
		e.preventDefault();
		cur=$(this).text();
		if (cur=="By Manager" )
		{
			$("#filterLable").text("By Manager");
			$(this).text("By Fullname");
		}
		else
		{
			$("#filterLable").text("By Fullname");
			$(this).text("By Manager");
		}
		$("#btnFilter").dropdown("toggle");
		return false; 
	});
});

function myFunction() {
   var input, filter, table, tr, td, i, idx;
  input = document.getElementById("txtSearch");
  filter = input.value.toUpperCase();
  table = document.getElementById("tblList");
  tr = table.getElementsByTagName("tr");
	idx=0;
  if ($("#filterLable").text()=="By Manager")
	idx=5;
  // Loop through all table rows, and hide those who don't match the search query
  for (i = 0; i < tr.length; i++) {
    td = tr[i].getElementsByTagName("td")[idx];
    if (td) {
      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
        tr[i].style.display = "";
      } else {
        tr[i].style.display = "none";
      }
    } 
  }
}
//-->
</script>
</body>
</html>