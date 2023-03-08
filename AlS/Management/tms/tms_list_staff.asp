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
	
'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = 12'PageSizeDefault
		Set arrPre = Nothing
	Else
		intPageSize = 12'PageSizeDefault
	End If

'--------------------------------------------------
' Check ACCESS right
'--------------------------------------------------

	strTemp = Request.ServerVariables("URL") 
	While Instr(strTemp, "/")<>0
		strTemp = Mid(strTemp, Instr(strTemp, "/") + 1, Len(strTemp))
	Wend
	
	strFilename = strTemp
	
	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			
			If varGetRight(0, ii) = strTemp Then
				fgRight=True
				Exit For
			End If
		Next
		Set varGetRight = Nothing		
	End If	
	If fgRight = False Then		
		Response.Redirect("../../welcome.asp")
	End If
'--------------------------------------------------
' Check VIEWALL right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "view all" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------
' Check Enter Annual Leave right
'--------------------------------------------------
	fgHREnter=false
	If not isEmpty(session("RightOn")) Then
		varGetRight = session("RightOn")
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "Write Timesheet as HR control" OR intUserID=252 Then
				fgHREnter = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If



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
	If strChoseMenu = "" Then strChoseMenu = "AA"
	
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
	<div class="row" style="padding:20px 0px 0px 20px;"><h3>Write Timesheet For Employees</h3> </div>
	<div class="row">
        <div class="col-sm-6 col-sm-offset-3">
            <div id="imaginary_container"> 
			<form name="searchform" method="post" action="tms_list_staff.asp">
				<div class="form-group">
					<div class="input-group stylish-input-group">
						<input type="text" name="txtSearch" id="txtSearch" onkeyup="myFunction()" class="form-control"  placeholder="Search by fullname" >
						<span class="input-group-addon">
							<button type="submit" >
								<span class="glyphicon glyphicon-search"></span>
							</button>  
							
						</span>
					</div>
				</div>
			</form>
            </div>
        </div>
	</div>
	<div class="row">
		<form id="frmList" method="post" action="timesheet.asp">
			<div class="table-responsive">	
				<table class="table table-hover" id="tblList">
					<thead  class="thead-inverse tableheaderblue">
						<tr>
							<th>Full Name</th>
							<th>Jobtitle</th>
							<th>Department</th>
							<th>Report To</th>
<%if fgHREnter then%>							
							<th>Enter For Duration</th>
<%end if%>
						</tr>
					</thead>					
					
				</table>	
				<input type="hidden" id="txthidden" name="txthidden" value="">
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
        "ajax": "../../inc/JSON_listStaffsForWriteTimesheet.asp",
        "columns": [
            { "data": "Fullname" },
            { "data": "JobTitle" },
            { "data": "Department" },
            { "data": "ReportTo" }
<%if fgHREnter then%>							
			,{ "data": "Duration" }
<%end if%>			
        ]
    });	
	
<%if fgHREnter then%>	
	$('#tblList tbody').on('click', 'tr td:not(:last-child) ', function (e) {
<%else%>
	$('#tblList tbody').on('click', 'tr', function (e) {
<%end if%>
		e.preventDefault();
		var table = $('#tblList').DataTable();
        var dataid = table.row(this).id();
		
		$("#txthidden").val(dataid);		
		$("#frmList" ).submit(); 
    });
	
	
});

function myFunction() {
   var input, filter, table, tr, td, i;
  input = document.getElementById("txtSearch");
  filter = input.value.toUpperCase();
  table = document.getElementById("tblList");
  tr = table.getElementsByTagName("tr");

  // Loop through all table rows, and hide those who don't match the search query
  for (i = 0; i < tr.length; i++) {
    td = tr[i].getElementsByTagName("td")[0];
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