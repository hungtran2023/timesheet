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

Function Outbody(ByRef rsSrc)
	strOut = ""


	If Not rsSrc.EOF Then
		do while not rsSrc.EOF
			strApproval=""						
			if rsSrc("approved")>0 then strApproval="<img src='../../images/yes.gif'>"		
'Response.Write psize & "--" & Showlabel(rsSrc("Fullname")) & "<br>"			
			strOut = strOut & "<tr>" &_
			         "<td >" & Showlabel(rsSrc("Fullname")) & "</td>" &_
			         "<td>" & Showlabel(rsSrc("JobTitle")) & "</td>" &_
					 "<td>" & Showlabel(rsSrc("Department")) & "</td>" &_
			         "<td>" & Showlabel(rsSrc("ReportTo")) & "</td>" &_
			         "<td>" & strApproval & "</td>" &_
					 "<td><input type='checkbox' class='editor-active' value='"& rsSrc("PersonID") &"'></td>" &_
			         "</tr>" & chr(13)
			         			
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
	
	strAct = Request.QueryString("atc")

	if strAct<>"" then
		strWeekStart= cdate(Request.Form("txtstartDate"))
		intStatus=request.form("lbApproval")
	
		if intStatus="" then intStatus=0
		strSearch=""
		if cint(intStatus)=1 then 
			strSearch="ApprovalID IS NOT NULL AND"
		elseif cint(intStatus)=2 then 
			strSearch="ApprovalID IS NULL AND "
		end if
		
	else
		strWeekStart=Date - Weekday(Date, vbSaturday)+1 -7
	'		strWeekStart=Cdate("9-Dec-2017")
	end if
	
	strWeekEnd=strWeekStart + 6	
	
	 strSql = "SELECT PersonID, Fullname, JobTitle,Department, FirstNameLeader + ' ' + LastnameLeader as reportTo,ISNULL(b.ApprovalID,'0') as approved FROM HR_Employee a " & _
				" LEFT JOIN (SELECT staffID,ApprovalID FROM ATC_TimesheetApproval WHERE  DateFrom='" & strWeekStart & "' AND DateTo='"& strWeekEnd & "') b ON a.PersonID=b.StaffID" &_
				" WHERE " & strSearch & " FirstName<>'Managers' AND fgIndirect=0 ORDER BY a.FirstName"
				'response.write strSql
	Call GetRecordset(strSQL,rsData)
	strLast=Outbody(rsData)

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

.datepicker table tr td span.active{
    background: #04c!important;
    border-color: #04c!important;
}
.datepicker .datepicker-days tr td.active {
    background: #04c!important;
}
#week-picker-wrapper .datepicker .datepicker-days tr td.active~td, #week-picker-wrapper .datepicker .datepicker-days tr td.active {
    color: #fff;
    background-color: #04c;
    border-radius: 0;
}

#week-picker-wrapper .datepicker .datepicker-days tr:hover td, #week-picker-wrapper .datepicker table tr td.day:hover, #week-picker-wrapper .datepicker table tr td.focused {
    color: #000!important;
    background: #e5e2e3!important;
    border-radius: 0!important;
}
</style>	

</head>
<body data-pinterest-extension-installed="cr1.39.1">

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
	<div class="row" style="padding:20px 0px 0px 20px;">
		<h3>Approval register</h3> 
	</div>
	<form class="form-inline"  id="frmAprove" method="post" >
	<div class="row">
		
			<div class="form-group" id="week-picker-wrapper" >
				<label for="week" class="control-label" style="padding-left:20px">Select Week</label>
				<div class="input-group">
					<span class="input-group-btn">
						<button type="button" class="btn btn-rm week-prev">&laquo;</button>
					</span>
					<input type="text" class="form-control week-picker" placeholder="Select a Week">
					<span class="input-group-btn">
						<button type="button" class="btn btn-rm week-next">&raquo;</button>
					</span>
				</div>
			</div>
			<div class="form-group">
				<select class="form-control" id="lbApproval" name="lbApproval">
					<option value="0">All Staff</option>
					<option value="1">Approved</option>
					<option value="2">Not yet</option>
			  </select>
			</div>
		  <button type="submit" class="btn  btn-primary" id="btnSearch">Submit</button>
		  <input type="hidden" id="txtstartDate" name="txtstartDate" value="<%=strWeekStart%>">
		
	</div>
	<div class="row">
		
			<div class="panel-heading">			
				<div class="pull-right">
					<button type="submit" class="btn btn-primary" id="btnApprove">Approve</button>	
					<button type="button" class="btn btn-default btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>				
				</div>
			</div>
			<input type="hidden" id="txthidden" name="txthidden" value="">
			<input type="hidden" id="txtstaffIDs" name="txtstaffIDs" value="">
			<div class="panel panel-primary filterable">	
				
				<div class="table-responsive">	
				<div id='idWarning'></div>
					<table class="table table-hover" id="tblList">
						<thead  class="thead-inverse tableheaderblue">						
							<tr class="filters" style="background-color:#8ca0d1">								
								<th><input type="text" class="form-control blue-normal" placeholder="Fullname" disabled></th>
								<th><input type="text" class="form-control blue-normal" placeholder="Jobtitle" disabled></th>
								<th><input type="text" class="form-control blue-normal" placeholder="Department" disabled></th>
								<th><input type="text" class="form-control blue-normal" placeholder="Report To" disabled></th>
								<th style="width:5%"></th>
								<th style="width:5%"><input type="checkbox" class="editor-active" id="select_all_existent"></th>
								
							</tr>
						</thead>					
						<tbody>
<%=strLast%>						
						</tbody>
					</table>	
					
				</div>		
			
			</div>		
			
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

<<script type="text/javascript" src="../../js/jquery-3.2.1.min.js"></script>
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
	
    $('#tblList').DataTable( {       
        //"serverSide": true,
        //"ajax": "../../inc/JSON_listStaffs.asp",
        "columns": [
			{ "data": "Fullname" },
            { "data": "JobTitle" },
            { "data": "Department" },
            { "data": "ReportTo" },
			{ "data": "approved" },
			{ "data":   "active"            
			}
			
        ]
    });		
	
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
	
	$('#btnApprove').click(function(e){
	
		e.preventDefault(); 
		var dataArr = [];
       $('input:checked').not('#select_all_existent').each(function(){
            //dataArr.push($(this).closest('tr[id]').attr('id')); // insert rowid's to array
			dataArr.push($(this).val());
		});		
		if (dataArr!="")
			{
				
				$("#txtstaffIDs").val(dataArr);
				$("#frmAprove").attr('action', 'tms_approvalByTeam.asp').submit();
			}
		
		else
		{
			showalert("You must select at least one checkbox before a full name.","alert-danger");
		}
		
	});
	
	$('#btnSearch').click(function(e){
	
		e.preventDefault(); 
		//alert((start_date.getMonth() + 1)+ '/' + (start_date.getDate())  + '/' + start_date.getFullYear());
		$("#txtstartDate").val((start_date.getMonth() + 1)+ '/' + (start_date.getDate())  + '/' + start_date.getFullYear());
		$("#frmAprove").attr('action', 'tms_listfor_approval.asp?atc=1').submit();
				
	});
	 //// Align modal when user resize the window
	 $('#select_all_existent').change(function(){
		var $table = $(this).parents('.table'),
			$rows = $table.find('tbody tr');
		//var $table = $(this).parents('.row');
		var cells = $rows.filter(function(){
			return $(this).is(":visible"); 
		});
		
		$( cells).find(':checkbox').prop('checked', $(this).is(':checked'));
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