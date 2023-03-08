<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strError, PageSize, fgRight 'view all or Not
	Dim intApproval

'****************************************
' Function: Outbody
' Description: 
' Parameters: source recordset, number of lines on one page
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************

Function Outbody(ByRef rsSrc,intStatus)
	strOut = ""
	
	If Not rsSrc.EOF Then
		do while not rsSrc.EOF
			strApproval=""						
			dblBurn=0
			dblSale=0
			if cdbl(rsSrc("CSOHOurs"))>0 then dblBurn=(cdbl(rsSrc("ActualHours"))/cdbl(rsSrc("CSOHOurs")))*100
			if cdbl(rsSrc("CSOPayment"))>0 then dblSale=(cdbl(rsSrc("Sales"))/cdbl(rsSrc("CSOPayment")))*100
			if dblBurn=0 then
				strColor="#ffffff" 
			elseif dblBurn<101 then
				strColor="#c7edc4" 'green
			elseif dblBurn<=120 then
				strColor="#ffe4b3"
			elseif dblBurn>120 then
				strColor="#ffb3b3" 'red
			end if
			
			dblCompare=cint(rsSrc("EstValue"))-cint(dblBurn)

			if dblBurn=0 then
				strColor="#ffffff" 
			elseif dblCompare>=-15 then
				strColor1="#a2e19d" 'green
			elseif dblCompare>=-30  then
				strColor1="#ffd280" 
			else
				strColor1="#ff8080"
			end if
			
			if (cint(intStatus)<>2) OR (cint(intStatus)=2 AND dblBurn>120) then 
				strOut = strOut & "<tr>" &_
						 "<td >" & Showlabel(rsSrc("ProjectID")) & "</td>" &_
						 "<td>" & Showlabel(rsSrc("ProjectName")) & "</td>" &_
						 "<td>" & Showlabel(rsSrc("GM")) & "</td>" &_
						 "<td>" & Showlabel(rsSrc("Manager")) & "</td>" &_
						 "<td>" & formatnumber(rsSrc("CSOHOurs"),2) & "</td>" &_
						 "<td>" & formatnumber(rsSrc("ActualHours"),2) & "</td>" &_
						 "<td bgcolor='" & strColor & "'> "& formatnumber(dblBurn,0) & "%</td>" &_
						 "<td>" & "<button class='editbtn btn btn-default'>"& rsSrc("EstValue") &"%</button>" & "</td>" &_
						 "<td bgcolor='" & strColor1 & "'>" & formatnumber(dblCompare,0) & "%</td>" &_
						 "<td>" & formatnumber(rsSrc("CSOPayment"),2) & "</td>" &_
						 "<td>" & formatnumber(rsSrc("Sales"),2) & "</td>" &_
						 "<td>" & formatnumber(rsSrc("InvoiceValueUSD"),2) & "</td>" &_
						  "<td>" & formatnumber(dblSale,2) & "%</td>" &_
						 "</tr>" & chr(13)
			end if
			rsSrc.MoveNext
		loop		
		
	End If
	Outbody = strOut
End Function
    
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
' Initialize appoval timesheet records
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	strSearch=""
	
		
	intStatus=request.form("lbCriteria")
	
	if intStatus="" then intStatus=-1
'(b.ManagerID=1303 or b.ManagerID IN (SELECT StaffID FROM [dbo].[UserByReportTo] (1303) where sTAFFid IN (select UserID fROM HR_ReceiveReport)))	
	if cint(intStatus)=-1 then 
		strSearch=" (B.ManagerID=" & intUserID & " OR  b.ManagerID IN (SELECT StaffID FROM [dbo].[UserByReportTo] ("& intUserID &") where sTAFFid IN (select UserID fROM HR_ReceiveReport)))"
	elseif cint(intStatus)=1 then 
		strSearch=" A.CSOHOurs >=250 "
	elseif cint(intStatus)=2 then 
	
	end if
		
		
	 strSql = "SELECT A.*,B.ProjectName, B.ManagerID, (C.FirstName + ' ' + C.LastName) as Manager, (F.FirstNameLeader + ' ' + F.LastnameLeader) as GM, fgActivate, ISNULL(D.EstValue,0) as EstValue FROM " & _
				"(SELECT ProjectID,  SUM(Hours+OTHours) as ActualHours, SUM(InvoiceValue) as Sales, SUM(CSOHours) as CSOHOurs, SUM(CSOPayment) as CSOPayment, SUM(InvoiceValueUSD) as InvoiceValueUSD " & _
				"FROM rp_ProjectPerformanceByPeriod GROUP BY ProjectID) A INNER JOIN ATC_Projects B ON A.ProjectID=B.ProjectID LEFT JOIN ATC_PersonalInfo C ON B.ManagerID=C.PersonID " & _
				" LEFT JOIN HR_EmployeeAll F ON B.ManagerID=F.PersonID " & _
				" LEFT JOIN rp_GetTheLastEstimation D ON A.ProjectID=D.ProjectID  WHERE fgActivate=1 " & _
				" AND SUBSTRING(A.ProjectID,11,1) NOT IN ('T','V','R','M') AND ManagerID not IN (SELECT StaffID FROM ATC_Employees WHERE LeaveDate <=GetDate()) AND ActualHours>0 "'B.ManagerID=561"
	
	if strSearch<> "" then strSql=strSql & " AND " & strSearch
	
	strSql=strSql & " ORDER BY A.ProjectID" 
	
	'response.write strSql

	Call GetRecordset(strSQL,rsData)
	
	strLast=Outbody(rsData,intStatus)

'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	
	strFunction = "<div align='right'><a href='../../welcome.asp?choose_menu=B' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_

				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
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
	If strChoseMenu = "" Then strChoseMenu = "B"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../../templates/template1/")

'Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)

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
<link href="../../bootstrap/css/bootstrap-datepicker.css" rel="stylesheet" type="text/css">
<link href="../../css/timesheet.css" rel="stylesheet" >
<link href="../../css/style.css" rel="stylesheet" type="text/css">    
<style>
.modal-dialog{
   width: 60%;
   margin: auto;
}

.datepicker { font-size: 10px; }


.filterable {
    margin-top: 15px;
}
.filterable {
    margin-top: -20px;
}
.filterable .filters input[disabled] {
    background-color: transparent;
    border: none;
    cursor: auto;
    box-shadow: none;
    padding: 0;
    height: auto;
	color: white;
	font-weight: bold;
    font-size: 12px;
    font-family: arial;
	
}
.filterable .filters input[disabled]::-webkit-input-placeholder {
    color: white;

}
.filterable .filters input[disabled]::-moz-placeholder {
    color: white;

}
.filterable .filters input[disabled]:-ms-input-placeholder {
    color: white;

}
.filterable .filters input

/* basic positioning */
.legend {  list-style: none;}
.legend li { float: left; margin-right: 20px; }
.legend span { border: 1px solid #ccc; float: left; width: 12px; height: 12px; margin: 2px; }
/* your colors */
.legend .green { background-color: #bde9ba; }
.legend .yellow { background-color: #ffd78c; }
.legend .red { background-color: #FFB9B9; }
</style>	

</head>
<body data-pinterest-extension-installed="cr1.39.1">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))

%>
<table width="95%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td style="vertical-align: top;">
<div class="container-fluid" >

	<div class="row" style="padding:0px 0px 0px 20px;">
		<h3>Project Performance</h3> 
	</div>
	<div class="row">
		<form class="form-inline"  id="frmSearch" method="post" >
			<div class="form-group">
				<label for="week" class="control-label" style="padding-left:20px">Search </label>
				<select class="form-control" id="lbCriteria" name="lbCriteria">
					<option value="-1" <%if cint(intStatus)=-1 then%>selected<%end if%> >your projects</option>
					<option value="0" <%if cint(intStatus)=0 then%>selected<%end if%>>all active projects</option>
					<option value="1" <%if cint(intStatus)=1 then%>selected<%end if%>>greater than  250 man hours</option>
					<option value="2" <%if cint(intStatus)=2 then%>selected<%end if%>>a significant overburn from 20%</option>
			  </select>
			   <button type="submit" class="btn  btn-primary" id="btnSearch">Submit</button>
			</div>
			<div class="form-group">
				<div>
					<ul class="legend">
						<li><span class="green"></span>On track</li>
						<li><span class="yellow"></span> Slightly behind schedule</li>
						<li><span class="red"></span>Needs immediate attention</li>
					</ul>

				</div>
			</div>
		 
		</form>
		
			
	</div>
<%If strError <> ""  Then%>  
	<div class="row">	
			<div class="<%if strError="Update successfull." then %>alert alert-danger<%else%>alert alert-success<%end if%>">
				<strong>Error:</strong><%=strError%>
			</div>
	</div>
<%elseif strLast="" then%>	
	<div class="row">	
			<div class="alert alert-info">
				There are no matching APK
			</div>
	</div>
<% End If%>		
	<form class="form-inline"  id="frmAprove" method="post" >
		<div class="row">
			
			<div class="panel-heading">	
				
				<div class="pull-right" style="margin-bottom:5px;">
					<!--<button type="submit" class="btn btn-primary" id="btnApprove">Approve</button>	-->
					<button type="button" class="btn btn-default btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>				
				</div>
			</div>
			
			<input type="hidden" id="txthidden" name="txthidden" value="">
			<input type="hidden" id="txtstaffIDs" name="txtstaffIDs" value="">
			<div class="panel panel-primary filterable" >	
				
			<div class="table-responsive">	
				<div id='idWarning'></div>
					<table class="table table-hover" id="tblList">
						<thead  class="thead-inverse tableheaderblue" >	

							<tr class="filters"  style="background-color:#8ca0d1">	
								<th rowspan="3"><input type="text" class="form-control blue-normal" placeholder="APK" disabled></th>
								<th rowspan="3"><input type="text" class="form-control blue-normal" placeholder="ProjectName" disabled></th>
								<th rowspan="3"><input type="text" id="txtGManager" class="form-control blue-normal" placeholder="Group Manager" disabled></th>
								<th rowspan="3"><input type="text" id="txtmanager" class="form-control blue-normal" placeholder="Manager" disabled></th>
								<td colspan="5"><span style="color:white; font-weight: bold">Man Hour Comparison</span></td>							
								<td colspan="4"><span style="color:white; font-weight: bold">Value Comparison</span></td>								
							</tr>
							
							<tr>
								<th>CSO Hours</th>
								<th>Actual Hours</th>
								<th>% (1)</th>
								<th>Est. % complete (2)</th>
								<th> Variance <br> (2) - (1) </th>
								<th>CSO Value</th>
								<th>Invoice Value</th>
								<th>Invoice Value (USD)</th>
								<th>%</th>
							</tr>
						</thead>
<%if strLast<>"" then%>						
						<tbody>
						
<%=strLast%>
						</tbody>
<%end if%>
						
				</table>	
				
			</div>		
			
			</div>		
			
    </div> 
	</form>	
</div>

</div>      
</td>
</tr>
</table> 
      
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(1))
%>
<input type="hidden" name="txthidden" id="txthidden" value="">

<script type="text/javascript" src="../../js/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/dataTables.bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap-datepicker.js" charset="UTF-8"></script>

<script type="text/javascript" src="../../js/library.js"></script>
<script src="../../bootstrap/js/bootstrap.min.js"></script>

<script language="javascript">
<!--
var weekpicker, start_date, end_date;

$.extend( true, $.fn.dataTable.defaults, {
     "paging": false,
		"searching": false,
		"processing": true,
		"ordering": false,
		"info":     false
} );


$(document).ready(function() {
	
	weekpicker = $('.week-picker');
    
    weekpicker.datepicker({
		weekStart: 6, 
		autoclose: true,
        forceParse: false,
    }).on("changeDate", function(e) {
        set_week_picker(e.date);
    });
    $('.week-prev').on('click', function() {
        var prev = new Date(start_date.getTime());
        prev.setDate(prev.getDate() - 1);
        set_week_picker(prev);
    });
    $('.week-next').on('click', function() {
	
        var next = new Date(end_date.getTime());
        next.setDate(next.getDate() + 2);
		
        set_week_picker(next);
    });
	
    set_week_picker(new Date(<%=year(strWeekStart)%>,<%=month(strWeekStart)-1%>,<%=day(strWeekStart)+1%>));
	
    	
	 $('.btn-filter').click(function(){
        var $row = $(this).parents('.row'),
		$panel = $row.find('.filterable'),
        $filters = $panel.find('.filters input'),
        $tbody = $panel.find('.table tbody');
        if ($filters.prop('disabled') == true) {
            $filters.prop('disabled', false);
            $filters.first().focus();
        } else {
            $filters.val('').prop('disabled', true);
            $tbody.find('.no-result').remove();
            $tbody.find('tr').show();
        }
    });

    $('.filterable .filters input').keyup(function(e){
        /* Ignore tab key */
        var code = e.keyCode || e.which;
        if (code == '9') return;
		
        /* Useful DOM data and selectors */
        var $input = $(this),
        inputContent = $input.val().toLowerCase(),
		$rowss = $(this).parents('.row'),
		$panel = $rowss.find('.filterable'),
		//$panel = $input.parents('.filterable'),
        column = $panel.find('.filters th').index($input.parents('th')),
        $table = $panel.find('.table'),
        $rows = $table.find('tbody tr');
		
		 /*clear all input */
		 
        /* Dirtiest filter function ever ;) */
        var $filteredRows = $rows.filter(function(){
            var value = $(this).find('td').eq(column).text().toLowerCase();
            return value.indexOf(inputContent) === -1;
        });
        /* Clean previous no-result if exist */
        $table.find('tbody .no-result').remove();
        /* Show all rows, hide filtered ones (never do that outside of a demo ! xD) */
        $rows.show();
        $filteredRows.hide();
        /* Prepend no-result row if all rows are filtered */
        if ($filteredRows.length === $rows.length) {
            $table.find('tbody').prepend($('<tr class="no-result text-center"><td colspan="'+ $table.find('.filters th').length +'">No result found</td></tr>'));
        }
				
		var $inputClears=$panel.find('.filters input');		 
		$inputClears.val('');
		
		$input.val(inputContent);
    });	

	$('#btnSearch').click(function(e){
	
		e.preventDefault(); 
		
		$("#frmSearch").attr('action', 'ProjectPerformance.asp').submit();
				
	});
 
	 $('.editbtn').click(function(e){
	 
		e.preventDefault(); 
		$("#txthidden").val($(this).parent().siblings(":first").text());
		$("#frmAprove").attr('action', 'ProjectPerformanceDetails.asp').submit();
        //alert($(this).parent().siblings(":first").text());
    });
	
});

function showalert(message,alerttype) {

    $('#idWarning').append('<div id="alertdiv" class="alert ' +  alerttype + '"><a class="close" data-dismiss="alert">×</a><span>'+message+'</span></div>')
    setTimeout(function() { // this will automatically close the alert and remove this if the users doesnt close it in 5 secs
		$("#alertdiv").remove();
		}, 5000);
  }
function set_week_picker(date) {
    start_date = new Date(date.getFullYear(), date.getMonth(), date.getDate() - date.getDay()-1);
    end_date = new Date(date.getFullYear(), date.getMonth(), date.getDate() - date.getDay() + 5);
    weekpicker.datepicker('update', start_date);
    weekpicker.val((start_date.getDate()) + '/' + (start_date.getMonth() + 1) + '/' + start_date.getFullYear() + ' - ' + (end_date.getDate()) + '/' + (end_date.getMonth()+1) + '/' + end_date.getFullYear());
}
//-->
</script>
</body>
</html>