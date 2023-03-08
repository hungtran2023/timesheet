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
	


'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					
'--------------------------------------------------
' Check Approving Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgUnlock = False
	Else
		varGetRight = session("RightOn")
		fgUnlock = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "Unlock Timesheet" Then
				fgUnlock = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------	

	intStaffIDs  = Request.Form("txthidden")
	
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view timesheet&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
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

<link href="../../css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../css/timesheet.css" rel="stylesheet" >
<link rel="stylesheet" href="../../bootstrap/css/bootstrap-datetimepicker.min.css"/>     
     
      
<script type="text/javascript" src="../../jQuery/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>

<script type="text/javascript" src="../../bootstrap/js/bootstrap-datepicker.min.js"></script>

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
  font-weight: bold;
  text-align: left;

.tmstable tbody td.num{
text-align: center !important;
}

.bootstrap-datetimepicker-widget tr:hover {
    background-color: #808080;
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
	<form class="form-inline"  id="frmSearch" method="post" action="FingerprintByTeam.asp">
		  <div class="form-group">
			<label for="lblYear">Search in: </label>
			<select class="form-control" id="lblYear" name="lblYear">			
				<option value="<%=Year(Date())%>" <%if cint(intYear)=Year(Date()) then%>selected<%end if%>><%=Year(Date())%></option>
				<option value="<%=Year(Date())-1%>" <%if cint(intYear)=Year(Date())-1 then%>selected<%end if%>><%=Year(Date())-1%></option>
				
			  </select>
		  </div>
		  <div class="form-group">
			<label for="lblMonth"></label>
			<select class="form-control" id="lblMonth" name="lblMonth">
<%for ii=1 to 12%>			
				<option value="<%=ii%>" <%if cint(intMonth)=ii then%>selected<%end if%>><%=MonthName(ii)%></option>
<%next%>
			  </select>
		  </div>
		  
		  <button type="submit" class="btn  btn-primary" id="btnSubmit">Submit</button>
		</form>
	</div>
	<div class="row">
        <div class="col-sm-6 form-group">
            <div class="input-group" id="DateDemo">
              <input  class="form-control" type='text' id='weeklyDatePicker' name="weeklyDatePicker" placeholder="Select Week" />
			   <input class="form-control" id="date" name="date" placeholder="MM/DD/YYY" type="text"/>
      
          </div>
      </div>
	</div>
	
	<div class="row">
		<table class="table tmstable" >
		
		  <thead>
			<tr>
			  <th scope="col"><input type="checkbox" class="editor-active" id="select_all_existent"></th>
			  <th scope="col">Fullname</th>
			  <th scope="col">Project</th>
			  <th scope="col" class="weekend-saturday">23-Sep</th>
			  <th scope="col" class="weekend-sunday">24-Sep</th>
			  <th scope="col">25-Sep</th>
			  <th scope="col">26-Sep</th>
			  <th scope="col">27-Sep</th>
			  <th scope="col">28-Sep</th>
			  <th scope="col">29-Sep</th>
			</tr>
		  </thead>
		  <tbody>
			
			<tr style="border-top: 2px solid  #617DC0">
			  <td scope="row" ><input type="checkbox" class="editor-active" ></td>
			  <td class="fullname">Anh le Ngoc Van</td>
			  <td class="project">ATL0305C00NVN0Z</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">AXS0041030LGB4A</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">MKE0137010TGB4A</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td></td>
			  <td></td>
			  <td></td>
			  <td></td>
			  <td></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">Other Hours</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project" style="text-align:right">Normal Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project" style="text-align:right">OT Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project"  style="text-align:right">Total Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			
			
			<tr style="border-top: 2px solid  #617DC0">
			  <td scope="row"><input type="checkbox" class="editor-active" ></td>
			  <td class="fullname">Anh le Ngoc Van</td>
			  <td class="project">ATL0305C00NVN0Z</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">AXS0041030LGB4A</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">MKE0137010TGB4A</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td></td>
			  <td></td>
			  <td></td>
			  <td></td>
			  <td></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">Other Hours</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project" style="text-align:right">Normal Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project" style="text-align:right">OT Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project"  style="text-align:right">Total Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr style="border-top: 2px solid  #617DC0">
			  <td scope="row"><input type="checkbox" class="editor-active" ></td>
			  <td class="fullname">Anh le Ngoc Van</td>
			  <td class="project">ATL0305C00NVN0Z</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">AXS0041030LGB4A</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">MKE0137010TGB4A</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td></td>
			  <td></td>
			  <td></td>
			  <td></td>
			  <td></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project">Other Hours</td>
			  <td class="weekend-saturday">1.0</td>
			  <td class="weekend-sunday">1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			  <td>1.0</td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project" style="text-align:right">Normal Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project" style="text-align:right">OT Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
			<tr>
			  <td scope="row"></td>
			  <td></td>
			  <td class="project"  style="text-align:right">Total Hours</td>
			  <td class="weekend-sunday" colspan=2>1.0</td>
			  <td colspan=5></td>
			</tr>
		  </tbody>
	</table>
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



<script language="javascript">
<!--

$(document).ready(function() {
 

  var date_input=$('input[name="date"]'); //our date input has the name "date"
      //var container=$('.bootstrap-iso form').length>0 ? $('.bootstrap-iso form').parent() : "body";
      var options={
        format: 'mm/dd/yyyy',
        //container: container,
        todayHighlight: true,
        autoclose: true,
      };
      date_i'nput.datepicker(options);
	  $("#date").datepicker(options);
	  
	   var date_input1=$('input[name="weeklyDatePicker"]'); //our date input has the name "date"
	  date_input1.datepicker(options);
	  
	  
	  
	  
   //Get the value of Start and End of Week
  $('#weeklyDatePicker').on('dp.change', function (e) {
      var value = $("#weeklyDatePicker").val();
      var firstDate = moment(value, "MM-DD-YYYY").day(0).format("MM-DD-YYYY");
      var lastDate =  moment(value, "MM-DD-YYYY").day(6).format("MM-DD-YYYY");
      $("#weeklyDatePicker").val(firstDate + " - " + lastDate);
  });
});

function back_menu(){
  window.history.back();
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
