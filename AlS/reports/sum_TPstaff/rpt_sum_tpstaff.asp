<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
dim dblSumHours,dblSumOT
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

Function OutbodyByStaff(ByRef rsSrc,intStatus)
	strOut = ""
dblSumHours=0
dblSumOT=0
	If Not rsSrc.EOF Then
		strOut="<tbody id='myTable'>"
		do while not rsSrc.EOF
				
				strOut = strOut & "<tr>"
				strOut = strOut & "<td>" & rsSrc("Fullname") & "</td>"
				strOut = strOut & "<td>"&rsSrc("CompanyName") & "</td>"
				strOut = strOut & "<td>"&rsSrc("ProjectID") & "</td>"
				strOut = strOut & "<td>"&rsSrc("SubTaskName") & "</td>"
				strOut = strOut & "<td class='hours'>"&rsSrc("Hours") & "</td>"
				strOut = strOut & "<td class='OThours'>"&rsSrc("OverTime") & "</td>"
				strOut = strOut & "<td >"&cdbl(rsSrc("Hours"))+cdbl(rsSrc("OverTime")) & "</td>"
				strOut = strOut & "</tr>" & chr(13)
				dblSumHours=dblSumHours+cdbl(rsSrc("Hours"))
				dblSumOT=dblSumOT+cdbl(rsSrc("OverTime"))
			rsSrc.MoveNext
		loop	
		strOut=strOut& "<tfoot style='background-color:#8ca0d1'><tr><td colspan=4 style='text-align:right;padding-right:20px'><b>Sub Total</b></td>"
		strOut=strOut& "<td id='sumHours'>" & dblSumHours & "</td>"
		strOut=strOut& "<td  id='sumOTHours'>" & dblSumOT & "</td>"
		strOut=strOut& "<td id='sumTotal'>" & dblSumHours + dblSumOT & "</tr></tfoot>"
		strOut=strOut & "</tbody>"
	End If
	OutbodyByStaff = strOut
End Function
    
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

Function OutbodyByCompany(ByRef rsSrc,intStatus)
	strOut = ""
dblSumHours=0
dblSumOT=0
	If Not rsSrc.EOF Then
		strOut="<tbody id='myTable'>"
		do while not rsSrc.EOF
				
				strOut = strOut & "<tr>"
				strOut = strOut & "<td>"&rsSrc("CompanyName") & "</td>"
				strOut = strOut & "<td>"&rsSrc("ProjectID") & "</td>"
				strOut = strOut & "<td>"&rsSrc("SubTaskName") & "</td>"
				strOut = strOut & "<td class='hours'>"&rsSrc("Hours") & "</td>"
				strOut = strOut & "<td class='OThours'>"&rsSrc("OverTime") & "</td>"
				strOut = strOut & "<td >"&cdbl(rsSrc("Hours"))+cdbl(rsSrc("OverTime")) & "</td>"
				strOut = strOut & "</tr>" & chr(13)
				dblSumHours=dblSumHours+cdbl(rsSrc("Hours"))
				dblSumOT=dblSumOT+cdbl(rsSrc("OverTime"))
			rsSrc.MoveNext
		loop	
		strOut=strOut& "<tfoot style='background-color:#8ca0d1'><tr><td colspan=3 style='text-align:right;padding-right:20px'><b>Sub Total</b></td>"
		strOut=strOut& "<td id='sumHours'>" & dblSumHours & "</td>"
		strOut=strOut& "<td  id='sumOTHours'>" & dblSumOT & "</td>"
		strOut=strOut& "<td id='sumTotal'>" & dblSumHours + dblSumOT & "</tr></tfoot>"
		strOut=strOut & "</tbody>"
	End If
	OutbodyByCompany = strOut
End Function

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

Function OutbodyByAPK(ByRef rsSrc,intStatus)
	strOut = ""
dblSumHours=0
dblSumOT=0
	If Not rsSrc.EOF Then
		strOut="<tbody id='myTable'>"
		do while not rsSrc.EOF
				
				strOut = strOut & "<tr>"
				strOut = strOut & "<td>"&rsSrc("ProjectID") & "</td>"
				strOut = strOut & "<td>"&rsSrc("SubTaskName") & "</td>"
				strOut = strOut & "<td>"&rsSrc("CompanyName") & "</td>"
				strOut = strOut & "<td class='hours'>"&rsSrc("Hours") & "</td>"
				strOut = strOut & "<td class='OThours'>"&rsSrc("OverTime") & "</td>"
				strOut = strOut & "<td >"&cdbl(rsSrc("Hours"))+cdbl(rsSrc("OverTime")) & "</td>"
				strOut = strOut & "</tr>" & chr(13)
				dblSumHours=dblSumHours+cdbl(rsSrc("Hours"))
				dblSumOT=dblSumOT+cdbl(rsSrc("OverTime"))
			rsSrc.MoveNext
		loop	
		strOut=strOut& "<tfoot style='background-color:#8ca0d1'><tr><td colspan=3 style='text-align:right;padding-right:20px'><b>Sub Total</b></td>"
		strOut=strOut& "<td id='sumHours'>" & dblSumHours & "</td>"
		strOut=strOut& "<td  id='sumOTHours'>" & dblSumOT & "</td>"
		strOut=strOut& "<td id='sumTotal'>" & dblSumHours + dblSumOT & "</tr></tfoot>"
		strOut=strOut & "</tbody>"
	End If
	OutbodyByAPK = strOut
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
	strGroupby=request.form("radGroupBy")
	if strGroupby="" then strGroupby=1

	strStartDate=request.form("start")
	strEndDate=request.form("end")
	
	if strStartDate="" or strEndDate="" then
		strStartDate=DateSerial(Year(date()), Month(date()), 1)
		strEndDate=DateSerial(Year(date()), Month(date()) + 1, 0)
	else
		varFrom		= split(strStartDate,"/")
		strStartDate		= CDate(varFrom(1) & "/" & varFrom(0) & "/" & varFrom(2))
		varTo		= split(strEndDate,"/")
		strEndDate		= CDate(varTo(1) & "/" & varTo(0) & "/" & varTo(2))			
	end if
	
	

	select case cint(strGroupby)
		case 1
		strSql="SELECT  Fullname, CompanyName, ProjectID, SubTaskName, Hours, OverTime  FROM   " & _
			"(SELECT StaffID,AssignmentID, SUM(hours) as Hours,  SUM(OverTime) as OverTime FROM [rpt_TimesheetAll] WHERE (AssignmentID <> 1) AND Tdate BETWEEN '" & strStartDate &"' AND '" & strEndDate & "' GROUP BY StaffID,AssignmentID ) as a " & _
			"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
			"INNER JOIN ATC_Tasks c ON c.SubtaskID=b.SubtaskID " & _
			"INNER JOIN [dbo].[HR_TPStaff] as d ON a.StaffID=d.TPUserID "
			Call GetRecordset(strSQL,rsData)
	
			strLast=OutbodyByStaff(rsData,intStatus)
		case 2
			strSql="SELECT  ProjectID, SubTaskName,CompanyName, SUM(hours) as Hours,  SUM(OverTime) as OverTime  FROM   " & _
			"(SELECT StaffID,AssignmentID, SUM(hours) as Hours,  SUM(OverTime) as OverTime FROM [rpt_TimesheetAll] WHERE (AssignmentID <> 1) AND Tdate BETWEEN '" & strStartDate &"' AND '" & strEndDate & "' GROUP BY StaffID,AssignmentID ) as a " & _
			"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
			"INNER JOIN ATC_Tasks c ON c.SubtaskID=b.SubtaskID " & _
			"INNER JOIN [dbo].[HR_TPStaff] as d ON a.StaffID=d.TPUserID GROUP BY ProjectID,SubTaskName,CompanyName"
			
			Call GetRecordset(strSQL,rsData)

			strLast=OutbodyByAPK(rsData,intStatus)
			
		case 3
			strSql="SELECT  CompanyName, ProjectID, SubTaskName, SUM(hours) as Hours,  SUM(OverTime) as OverTime  FROM   " & _
			"(SELECT StaffID,AssignmentID, SUM(hours) as Hours,  SUM(OverTime) as OverTime FROM [rpt_TimesheetAll] WHERE (AssignmentID <> 1) AND Tdate BETWEEN '" & strStartDate &"' AND '" & strEndDate & "' GROUP BY StaffID,AssignmentID ) as a " & _
			"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
			"INNER JOIN ATC_Tasks c ON c.SubtaskID=b.SubtaskID " & _
			"INNER JOIN [dbo].[HR_TPStaff] as d ON a.StaffID=d.TPUserID GROUP BY CompanyName,ProjectID,SubTaskName"

			Call GetRecordset(strSQL,rsData)

			strLast=OutbodyByCompany(rsData,intStatus)
	end select
'response.write strSql
'response.end

	
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
<head>
<link href="../../bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../bootstrap/css/bootstrap-datepicker.min.css" rel="stylesheet" type="text/css">

<link href="../../css/timesheet.css" rel="stylesheet" >
<link href="../../css/style.css" rel="stylesheet" type="text/css">  
<style>
.screenwidth {
  width: 97%;
}
.rowcontent
{
	padding-left:45px
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


</style>

</head>
<body data-pinterest-extension-installed="cr1.39.1">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
%>

<!----AIS content-->
<div class="container-fluid" >
	<div class="row rowcontent">
		<h3>Summary of TP Hours</h3>
		<form   id="frmSearch" method="post" action="rpt_sum_tpstaff.asp" style="width: 30%">
			
				
			<div class="form-group row">
				<label for ="" class="col-sm-2">Group by</label>
				<div class="col-sm-6">
					<div class="col-sm-1 no-padding width-auto">
						<input type="radio" name="radGroupBy" id="radStaff" value="1" class="no-padding" <%if cint(strGroupby)=1 then%>checked<%end if%>>
					</div>

					<label class="col-sm-3 padding-left5 no-blod" for="radStaff">Staff</label>

					<div class="col-sm-1 no-padding width-auto">
						<input type="radio" name="radGroupBy" id="radAPK" value="2"  class="no-padding" <%if cint(strGroupby)=2 then%>checked<%end if%>>
					</div>

					<label class="col-sm-3 padding-left5 no-blod" for="radAPK">APK</label>

					<div class="col-sm-1 no-padding width-auto">
						<input type="radio" name="radGroupBy" id="radCompany" value="3"  class="no-padding" <%if cint(strGroupby)=3 then%>checked<%end if%>>
					</div>

					<label class="col-sm-3 padding-left5 no-blod" for="radCompany">Company</label>
				</div>
			</div>			
			<div class="form-group row" id="sandbox-container" >
				  <div class="col-sm-2"><label>Select</label></div>
				  <div class="col-sm-8">
				  		<div class="input-daterange input-group " id="datepicker">
							<span class="input-group-addon">from</span>
							<input type="text" class="input-sm form-control" id="start" name="start"/>
							<span class="input-group-addon">to</span>
							<input type="text" class="input-sm form-control" id="end" name="end"/>	
						</div>
					</div>
				  <div class="col-sm-2"><button type="submit" class="btn btn-primary btn-sm" id="btnSearch">Submit</button></div>
				</div>		
							<div><span class="help-block" id="error" style="display: none;"></span></div>		
		</form>

	</div>
	
	<div class="row rowcontent" >
		<div class="screenwidth" style="background-color: lightblue;">	
			
			<button type="button" class="btn btn-default btn-filter" style="float: right;margin-bottom:10px"><span class="glyphicon glyphicon-filter"></span> Filter</button>				
		</div>

		<div class="filterable" >
			<table class="screenwidth table">
			  <thead style="background-color:#8ca0d1">
			  <tr class="filters" >
<%if cint(strGroupby)=1 then%>
				<th style="width:15%"><input type="text" class="form-control blue-normal" placeholder="Full Name" disabled></th>
				<th style="width:25%"><input type="text" class="form-control blue-normal" placeholder="Company" disabled></th>
				<th style="width:15%"><input type="text" class="form-control blue-normal" placeholder="APK" disabled></th>
				<th style="width:15%"><input type="text" class="form-control blue-normal" placeholder="Subtask" disabled></th>
<%elseif cint(strGroupby)=2 then%>
				<th style="width:20%"><input type="text" class="form-control blue-normal" placeholder="APK" disabled></th>
				<th style="width:20%"><input type="text" class="form-control blue-normal" placeholder="Subtask" disabled></th>
				<th style="width:30%"><input type="text" class="form-control blue-normal" placeholder="Company" disabled></th>
<%else%>
				<th style="width:30%"><input type="text" class="form-control blue-normal" placeholder="Company" disabled></th>
				<th style="width:20%"><input type="text" class="form-control blue-normal" placeholder="APK" disabled></th>
				<th style="width:20%"><input type="text" class="form-control blue-normal" placeholder="Subtask" disabled></th>
<%end if%>
				
				<th style="width:10%">Hours</th>
				<th style="width:10%">Overtime</th>
				<th style="width:10%">Total</th>
			  </tr>
			  </thead>
			  <%=strLast%>
			</table>
		</div>
			
    </div> 

</div>	
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------

	'Response.Write(arrPageTemplate(1))
%>
<!----END AIS content-->

</body>
</html>

<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap-datepicker.js" charset="UTF-8"></script>

<script type="text/javascript" src="../../js/library.js"></script>
<script type="text/javascript">

$(document).ready(function(){
	$('#sandbox-container .input-daterange').datepicker({		
			format: 'dd/mm/yyyy'			
	});
	

	
	$("#start").datepicker("setDate", new Date('<%=strStartDate%>'));
	
	$("#start").datepicker().on('changeDate',function(e){
		sDate = new Date($(this).datepicker('getUTCDate'));
		checkDate();
	});

	$("#end").datepicker("setDate", new Date('<%=strEndDate%>'));
	
	$("#end").datepicker().on('changeDate',function(date){
		eDate = new Date($(this).datepicker('getUTCDate'));
		checkDate();
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
        $table.find('tfoot').show();
        /* Show all rows, hide filtered ones (never do that outside of a demo ! xD) */
        $rows.show();
        $filteredRows.hide();

        /* Prepend no-result row if all rows are filtered */
        if ($filteredRows.length === $rows.length) {
            $table.find('tbody').prepend($('<tr class="no-result text-center"><td colspan="'+ $table.find('.filters th').length +'">No result found</td></tr>'));
            $table.find('tfoot').hide();
        }
        else
        {

        	$("#sumHours").html(function(){
        		var totalHours=0;

        		$filteredRows.each(function() {
       				 $(this).find('.hours').each(function(i){        
            			totalHours+=parseFloat( $(this).html());
        			});
       			});
    			
        		return (<%=dblSumHours%>-totalHours);
        	});

        	$("#sumOTHours").html(function(){
        		var totalOTHours=0;
        		
        		$filteredRows.each(function() {
       				 $(this).find('.OThours').each(function(i){        
            			totalOTHours+=parseFloat( $(this).html());
        			});
       			});
    			
        		return (<%=dblSumOT%>-totalOTHours);
        	});

        	$("#sumTotal").html(parseFloat( $("#sumHours").html())+parseFloat( $("#sumOTHours").html()));
        }
		


		var $inputClears=$panel.find('.filters input');		 
		$inputClears.val('');
		
		$input.val(inputContent);
    });	
	
	$('#btnSearch').click(function(e){
		e.preventDefault(); 

		$("#frmSearch").attr('action', 'rpt_sum_tpstaff.asp?a=1').submit();
		
	});
	
});


function checkDate()
{
    if(sDate && eDate && (eDate<sDate))
    {
		$("#error").parent().addClass("has-error");
        $('#error').html("The start date must be ealier than the end date");
		$("#error").show();		
    }
    else
    {
		$("#error").parent().remove("has-error");
        $('#error').html("");
		$("#error").hide();
    }
	
	return (sDate && eDate && (eDate<sDate));
}
</script>