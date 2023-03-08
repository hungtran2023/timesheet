<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

function GetWorkingHourFromTo(staffID, dateF, dateTo)
	dim dblHours
	strSQL="SELECT dbo.[GetWorkingHoursFromTo]("& staffID & ",'" &  dateF & "','" & dateTo & "')"
	Call GetRecordset(strSQL,rsData)
	if not rsData.EOF then
		dblHours=rsData.fields(0).value
	end if
	GetWorkingHourFromTo=dblHours
end Function
	
'-------------------------------------------------------------
'
'-------------------------------------------------------------
Function GetTimesheetForApproveByTeam()
	dim rs, strOut
	dim idxDate, intNameID,strClass,dblHours, strProjectID
	dim objConn 
	
	arrSubTitle = array("Normal Hours", "OT Hours","Total Hours") 
	dim dblSubTotal(2)
	
	set objConn= Server.CreateObject("ADODB.Connection")
	objConn.Open strConnect  
	
	If objConn.State=1 Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objConn
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "TimsheetForApprovalByTeam"
		
		Set myParam = myCmd.CreateParameter("staffIDs",adVarChar,adParamInput,4000)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("startDate",adDate,adParamInput)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("endDate",adDate,adParamInput)
		myCmd.Parameters.Append myParam

		'response.write intStaffIDs


		myCmd("staffIDs") = intStaffIDs
		myCmd("startDate") = strWeekStart
		myCmd("endDate") = strWeekEnd

		SET rs=myCmd.Execute

		strOut=""
		intNameID=-1
		
		if Not rs.EOF then
			do while not rs.EOF	
			
				if cdbl(intNameID)<>cdbl(rs("StaffID")) then			
					if strOut<>"" then
						dblSubTotal(2)=dblSubTotal(0)+dblSubTotal(1)
						for ii=0 to 2
							strOut=strOut & "<tr><td scope='row'></td><td></td>"
							strOut=strOut & "<td class='project' style='text-align:right'>" & arrSubTitle(ii)  & "</td>"
							strOut=strOut & "<td class='weekend-sunday' colspan=2>" & dblSubTotal(ii) 
							if ii=0 and dblSubTotal(ii)>=cdbl(dblLimitHours) then 
								strOut=strOut & "<img src='../../images/TsDone.jpg'/ style='padding-left:20px'>"
							elseif ii=0 then
								strOut=strOut & "<img src='../../images/TsMissing.jpg' style='padding-left:20px'/>"
							end if
							strOut=strOut & "</td><td colspan=5></td></tr>"
						next
					end if
					
					rsTmsApproval.filter="staffID=" & rs("StaffID")
					blnNotLock= (rsTmsApproval.EOF)
					
					rsTmsApproval.filter=""
					
					strOut=strOut& "<tr style='border-top: 2px solid  #617DC0'>"
					if blnNotLock then
						strOut=strOut& "<td scope='row'><input type='checkbox' class='editor-active' value='"& rs("StaffID") &"' ></td>"					
					else
						strOut=strOut& "<td scope='row'><button type='button' class='btn btn-primary btn-xs btn-unlock' >Unlock</button></td>"					
					end if
					strOut=strOut& "<td id='" & rs("StaffID") & "' class='fullname'>" &  rs("fullname") & "</td>"
					intNameID=rs("StaffID")
					dblSubTotal(0)=0
					dblSubTotal(1)=0
					dblLimitHours=GetWorkingHourFromTo(rs("StaffID"),strWeekStart,strWeekEnd)
				else
					strOut=strOut& "<tr>"
					strOut=strOut& "<td scope='row' ></td>"					
					strOut=strOut& "<td ></td>"	
				end if
			
			
				strOut=strOut& "<td class='project'><b>" & rs("ProjectID") & "<b></td>"		
				strProjectID = rs("ProjectID")
				
				for idxDate=strWeekStart to strWeekEnd
					
					strClass=""
					if Weekday(idxDate)=1  then 
						strClass= " class='weekend-sunday' "
					elseif Weekday(idxDate)=7 then
						strClass= " class='weekend-saturday' "
					end if
					dblHours=""
					if Not rs.EOF then
						if year(idxDate)=year(rs("TDate")) AND month(idxDate)=month(rs("TDate")) AND day(idxDate)=day(rs("TDate")) and strProjectID = rs("ProjectID") then
						
							dblHours=cdbl(rs("hours")) + cdbl(rs("OThours"))

							dblSubTotal(0)=dblSubTotal(0)+cdbl(rs("hours"))
							dblSubTotal(1)=dblSubTotal(1)+cdbl(rs("OThours"))

							strOut=strOut& "<td " & strClass &" >" & dblHours & "</td>"					
							if Not rs.EOF then rs.MoveNext
						else 
							strOut=strOut& "<td " & strClass &" ></td>"						
						end if
					else
						strOut=strOut& "<td " & strClass &" ></td>"	
					end if
					
				next

				strOut=strOut& "</tr>"				
				'if Not rs.EOF then rs.MoveNext
			loop
			
			if strOut<>"" then
				for ii=0 to 2
					strOut=strOut & "<tr><td scope='row'></td><td></td>"
					strOut=strOut & "<td class='project' style='text-align:right'>" & arrSubTitle(ii)  & "</td>"
					strOut=strOut & "<td class='weekend-sunday' colspan=2>"& dblSubTotal(ii) 
					if ii=0 and dblSubTotal(ii)>= cdbl(dblLimitHours) then 
						strOut=strOut & "<img src='../../images/TsDone.jpg'/ style='padding-left:20px'>"
					elseif ii=0 then
						strOut=strOut & "<img src='../../images/TsMissing.jpg' style='padding-left:20px'/>"
					end if
					strOut=strOut & "</td><td colspan=5></td></tr>"
				next
			end if					
		end if		
	end if
	
	GetTimesheetForApproveByTeam=strOut
	
End Function
	
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					
'--------------------------------------------------	
	dim intStaffIDs
	dim strWeekStart, strWeekEnd,strConnect
	intStaffIDs  = Request.Form("txtstaffIDs")
	
	intUserID	= session("USERID")
	'strWeekStart= Date - Weekday(Date, vbSaturday)+1
	strWeekStart= Request.Form("txtstartDate")
	if strWeekStart="" then 
		strWeekStart=Date - Weekday(Date, vbSaturday)+1 -7
	else 
		strWeekStart=cdate(strWeekStart)
	end if
	strWeekEnd=strWeekStart + 6	
'--------------------------------------------------
' Initialize appoval timesheet records
'--------------------------------------------------
	strConnect = Application("g_strConnect")	' Connection string 	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strConnect) Then
		strSQL = "SELECT * FROM ATC_TimesheetApproval WHERE DateFrom='" & strWeekStart & "' AND DateTo='" & strWeekEnd & "' AND StaffID In(" & intStaffIDs & ")"

		Set rsTmsApproval = Server.CreateObject("ADODB.Recordset")
		Set rsTmsApproval.ActiveConnection = objDatabase.cnDatabase
		rsTmsApproval.CursorLocation = adUseClient
			
		rsTmsApproval.LockType=3
			
		rsTmsApproval.Open strSQL
			
		If Err.number =>0 then	
			strError = Err.Description
		else
			set rsTmsApproval.ActiveConnection=nothing
		end if
	Else
			Response.Write objDatabase.strMessage		
	End If

	strAct=Request.QueryString("act")
	if strAct="app" then
		strError=""
		strStaffIDApps=split(request.form("txtstaffIDApp"),",")
		for each intStaffID in strStaffIDApps
			
			rsTmsApproval.filter="staffID=" & intStaffID
			if rsTmsApproval.EOF then
				rsTmsApproval.AddNew()
				rsTmsApproval("DateFrom")=strWeekStart
				rsTmsApproval("DateTo")=strWeekEnd
				rsTmsApproval("StaffID")=intStaffID
				rsTmsApproval("ApprovalID")=intUserID
			else
				strError=strError & "," & intStaffID
			end if
		next
		
		set rsTmsApproval.ActiveConnection = objDatabase.cnDatabase
		rsTmsApproval.UpdateBatch
		rsTmsApproval.Requery()
		set rsTmsApproval.ActiveConnection=nothing
		
	elseif strAct="un" then
		rsTmsApproval.filter="staffID=" & request.form("txtstaffIDApp")

		if rsTmsApproval.RecordCount=1 then
			rsTmsApproval.Delete(adAffectCurrent)
			set rsTmsApproval.ActiveConnection = objDatabase.cnDatabase
			rsTmsApproval.UpdateBatch
			rsTmsApproval.Requery()
			set rsTmsApproval.ActiveConnection=nothing
		end if
		
	end if

	strOut=GetTimesheetForApproveByTeam()
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Timesheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"

	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "templates/template1/")
%>	

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

.tmstable {
border: none ;
}
.tmstable th {
  background-color: #617DC0;
  color: #ffffff;
  text-align: center !important;
  border: 1px solid #8FA4D3 !important;
}
.tmstable th.weekend-saturday,
.tmstable th.weekend-sunday,
.tmstable th.authorizing__full-name {
  background-color: #617DC0;
}

.tmstable th.weekend-saturday,
.tmstable th.weekend-sunday {
	color: #ff9999;
}
.tmstable tr {
  background-color: #FFFFFF;
  color: #003399;
}
.tmstable td.weekend-saturday {
  background-color: #D2DAEC;
}

.tmstable td.weekend-sunday {
  background-color: #C2CCE7;
}

.tmstable td.project {
  background-color: #fff2f2;
}

.tmstable td.fullname {
	cursor: pointer;
	font-weight: bold;
	text-align: left;

.tmstable tbody td.num{
	text-align: center !important;
}

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
%>
<table width="95%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td style="vertical-align: top;">
		
		
<div class="container-fluid">
    <div class="row" style="padding:20px 0px 0px 0px;">
	<h2>Timesheet Approval by Team</h2>
	</div>
	
	<div class="row" >
	
	<form class="form-inline"  id="frmSearch" method="post" action="tms_approvalByTeam.asp">
		 <div class="form-group" id="week-picker-wrapper">
			<label for="week" class="control-label">Select Week</label>
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
		  <button type="submit" class="btn  btn-primary" id="btnSubmit">Submit</button>
		  <input type="hidden" id="txtstaffIDs" name="txtstaffIDs" value="<%=intStaffIDs%>">
		  <input type="hidden" id="txtstartDate" name="txtstartDate" value="<%=strWeekStart%>">
		  <input type="hidden" id="txtstaffIDApp" name="txtstaffIDApp" value="">
		  <input type="hidden" id="txtFrApp" name="txtFrApp" value="">
		
	</div>
	
	<div class="row">
			
					
				<button type="button" class="btn  btn-primary" id="btnApprove" style="margin-bottom:10px" 
						data-toggle="confirmation" data-placement="right"
						data-btn-ok-label="Continue" 
						data-btn-cancel-label="Cancel" 
						data-content="Before continuing make sure that you really checked these timesheets.">Approve </button>
	<div id='idWarning'></div>
				<table class="table tmstable" >
					  <thead>
						<tr>
						  <th scope="col"><input type="checkbox" class="editor-active" id="select_all_existent"></th>
						  <th scope="col">Fullname</th>
						  <th scope="col">Project</th>
						<%for idxDate=strWeekStart to strWeekEnd%>			  
						  <th scope='col' <%if Weekday(idxDate)=1 OR Weekday(idxDate)=7 then %> class='weekend-saturday'<%end if%>><%=ddmmmyyyy(idxDate)%></th>
						<%next%>
						</tr>
					  </thead>
					  <tbody>
							<%=strOut%>
					  </tbody>
				</table>				 
				 
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

<script type="text/javascript" src="../../jQuery/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/bootstrap-datepicker.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/bootstrap-confirmation.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>


<script language="javascript">
<!--

var weekpicker, start_date, end_date;

$(document).ready(function() {
 
  weekpicker = $('.week-picker');
    
    weekpicker.datepicker({
		weekStart: 6, 
		autoclose: true,
        forceParse: false,
        container: '#week-picker-wrapper',		
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
	
	$('#btnSubmit').click(function(e){
		e.preventDefault(); 
		//alert(start_date);
		$("#txtstartDate").val((start_date.getMonth() + 1)+ '/' + (start_date.getDate())  + '/' + start_date.getFullYear());
		$("#frmSearch").submit();
	});
	
	$('.fullname').click(function(e){
		e.preventDefault(); 
		var tdfullname = $(this).closest('td');
		
		//alert(tdfullname.attr('id'));
		$("#txtstartDate").val((start_date.getMonth() + 1)+ '/' + (start_date.getDate())  + '/' + start_date.getFullYear());
		$("#txtFrApp").val('app');
		$('#frmSearch').attr('action', 'timesheet.asp?app=1&id='+ tdfullname.attr('id')).submit();
		
	});
		
		 
	 $('#select_all_existent').change(function(){
		var $table = $(this).parents('.table');
		var $cells = $table.find('tbody tr td');
		$cells.find(':checkbox').prop('checked', $(this).is(':checked'));
	});
	
	$('[data-toggle=confirmation]').confirmation({
			rootSelector: '[data-toggle=confirmation]',
			popout: true
	  // other options
	});
	
	$('#btnApprove').click(function(e){
		e.preventDefault(); 
		
		var dataArr = [];
       $('input:checked').not('#select_all_existent').each(function(){
            dataArr.push($(this).val()); // insert rowid's to array
		});		
		if (dataArr!='')
		{
			$("#txtstaffIDApp").val(dataArr);			
			$('#frmSearch').attr('action', 'tms_approvalByTeam.asp?act=app').submit();
		}
		else
		{
			showalert("You must select at least one checkbox before a full name.","alert-danger");
		}
	});
	
	$('.btn-unlock').click(function(e){
		e.preventDefault(); 
		var tdfullname = $(this).parents('td').next();
		$("#txtstaffIDApp").val(tdfullname.attr('id'));
		$('#frmSearch').attr('action', 'tms_approvalByTeam.asp?act=un').submit();
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

function back_menu(){
  var url;
	url = "tms_listfor_approval.asp";
	
		document.location = url;
}

function logout()
{
	var url;
	url = "../../logout.asp";
	
		document.location = url;
	
}
//-->
</script>
</body>
</html>
