<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
--> 
<%
	Dim intUserID, intMonth, intYear, intDayNum, intWeekday, intRow, intCount,strAct,intNumberWeek
	Dim varFullName, varFrom, varTo, varPre, getRes, varUser, varInvalidTMS,rsEmailCount,rsWeekDate
	Dim strUserName, strTitle, strFunction, strMenu, strURL, strType, strTitle2, strFrom, strTo, strFirstDay, strCurDate, strDateShow
	
'***************************************************************
'
'***************************************************************
Function NumberOfWeek()
	dim intWeekcount
	dim startDate, endDtae

	startDate=DateSerial(intYear, intMonth, 1)
	endDate=DateSerial(intYear, intMonth+1, 1)-1

	strSql="SELECT * FROM MonthToWeek('" & startDate & "','" & endDate & "')"

    Call GetRecordset(strSql,rsWeekDate)

	NumberOfWeek=rsWeekDate.recordcount
End function

'***************************************************************
'
'***************************************************************
Function GetListMissing()
	dim rs, strOut
	dim ii, intColumns

	strOut=""
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "TimesheetAndNormalHours"
		Set myParam = myCmd.CreateParameter("month",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		
		Set myParam = myCmd.CreateParameter("year",adInteger,adParamInput)
		myCmd.Parameters.Append myParam

		myCmd("month") = intMonth
		myCmd("year") = intyear

		SET rs=myCmd.Execute
		strOut=""
		
		if Not rs.EOF then
			intColumns=intNumberWeek*2+3
			do while not rs.EOF
				
				strOut=strOut& "<tr><td><input type='checkbox' name='id[]' value='" & rs.fields(2) & + "'></td>"
				
				for ii=0 to intColumns-1
					if ii<>2 then
						strField=rs.fields(ii)
						if ii>=3 and ii<intColumns-1 then
							if (ii mod 2=1 and cdbl(rs.fields(ii))>=cdbl(rs.fields(ii+1))) OR (ii mod 2=0 and cdbl(rs.fields(ii))<=cdbl(rs.fields(ii-1))) then
								strField="-"
							end if
						end if
						if (ii=intColumns-1) then
							if cdbl(rs.fields(ii))<=cdbl(rs.fields(ii-1)) then strField="-"
						end if
						strOut=strOut& "<td>" & strField & "</td>"
					end if
				next	
				strOut=strOut& "</tr>"
				rs.MoveNext
			loop
		end if


	end if
	
	GetListMissing=strOut
End Function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID	= session("USERID")
	
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	intYear=request.form("lblYear")
	if intYear="" then intYear=year(date())
	intMonth=request.form("lblMonth")
	if intMonth="" then intMonth=month(date())

	strAct=request.querystring("act")
	
	if strAct="send" then
		strIDList=request.form("txtlistID")
		
		strSQL = "SELECT StaffID,ISNULL(a.EmailID + '@atlasindustries.com','atlas.ais.noreply@gmail.com') as EmailAddress, b.FirstName FROM ATC_Employees a "& _
					" LEFT JOIN ATC_PersonalInfo b ON a.StaffID= b.PersonID WHERE staffID IN (" & strIDList & ") ORDER BY b.FirstName"
					
		
		strSQL ="SELECT a.StaffID,ISNULL(a.EmailID + '@atlasindustries.com','atlas.ais.noreply@gmail.com') as EmailAddress, b.FirstName, ISNULL(c.EmailID + '@atlasindustries.com','atlas.ais.noreply@gmail.com') as EmailLeader FROM ATC_Employees a " & _
						 "LEFT JOIN ATC_PersonalInfo b ON a.StaffID= b.PersonID   " & _
							"LEFT JOIN ATC_Employees c ON a.DirectLeaderID= c.StaffID   " & _
						"WHERE a.staffID IN (" & strIDList & ") ORDER BY b.FirstName"
	
		call GetRecordset(strSQL,rsUsers)
		
		strSubject = Request.Form("txtsubject")
		strContent = Request.Form("txtMessage")
		
		do while not rsUsers.EOF
			
				Set cdoMessage = CreateObject("CDO.Message")  
					With cdoMessage 
						Set .Configuration = getCDOConfiguration()  
						.From = "no-reply@atlasindustries.com" 
						.To = rsUsers("EmailAddress") 
						'.To =  "uyenchi.nguyentai@atlasindustries.com"						
						.cc =rsUsers("EmailLeader")
						.Bcc="uyenchi.nguyentai@atlasindustries.com;"
						.Subject = strSubject
						.TextBody = replace(strContent,"#Name#",rsUsers("FirstName"))
if 	cdbl(intUserID) <>1329	then
						.Send 
end if
				End With

				Set cdoMessage = Nothing  
				Set cdoConfig = Nothing 
			
			rsUsers.MoveNext
		loop
		

	end if
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strFunction = "<a class='c' href='../../tools/preferences.asp' onMouseOver='self.status=&quot;Preferences&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Preferences</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage()' onMouseOver='self.status=&quot;Print missing timesheet page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='../../logout.asp' onMouseOver='self.status=&quot;Log out&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
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

'--------------------------------------------------
' Get current URL
'--------------------------------------------------
	
	If Request.ServerVariables("QUERY_STRING") <> "" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
'--------------------------------------------------
' Get current menu that user is choosing
'--------------------------------------------------
	
	strChoseMenu = Request.QueryString("choose_menu")
	If strChoseMenu = "" Then strChoseMenu = "B"

	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))

	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1) <> "" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if

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
<link href="../../bootstrap/css/dataTables.checkboxes.css" rel="stylesheet" type="text/css">
<link href="../../css/timesheet.css" rel="stylesheet" >
<link href="../../css/style.css" rel="stylesheet" type="text/css">  
<link href="../../css/bootstrapValidator.min.css" rel="stylesheet" type="text/css" />

</head>

<body data-pinterest-extension-installed="cr1.39.1">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
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
	<div class="row" style="padding:20px 0px 0px 20px;"><h3>Missing Timesheets</h3> </div>
		<form class="form-inline"  id="frmSearch" method="post" action="rpt_invalid_tms.asp">
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
<%intNumberWeek=NumberOfWeek()%>
	<div class="row" >
		<form id="frmList" method="post" action="rpt_invalid_tms.asp">
		<table class="table table-hover" id="tblList">
					<thead  class="thead-inverse tableheaderblue">
						<tr>	
							<th colspan="2"><button type="button" class="btn  btn-primary" id="btnSendemail">Send Email</button></th>
							<th  rowspan="2">Report To</th>		
<%rsWeekDate.moveFirst
do while not rsWeekDate.EOF%>							
							<th colspan="2">Week <%=rsWeekDate("weekcount")%> <span class="blue-normal">(<%=ddmmyyyy(rsWeekDate("WeekStart"))%>
										-<%=ddmmyyyy(rsWeekDate("WeekEnd"))%>)</span></th>	
<%
rsWeekDate.MoveNext
loop
%>
						</tr>						
						<tr>	
							<th><input type="checkbox" name="select_all" value="1" id="select-all"></th>
							<th>Full Name</th>
<%for ii=1 to intNumberWeek%>	
							<th >Entered</th>	
							<th >Standard</th>	
<%Next%>							
						</tr>						
					</thead>					
					<tbody>
					<%=GetListMissing()%>
					</tbody>
				</table>	
			</form>
	</div>
	
	
	 <!-- Modal HTML -->
   <!-- Modal HTML -->
    <div id="myModal" class="modal">
        <div class="modal-dialog">
            <div class="modal-content">
			<form id="detailform" class="form-horizontal" action="rpt_invalid_tms.asp?act=send" method="post">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                    <h4 class="modal-title">Reminder for submitting timesheet</h4>
                </div>
                <div class="modal-body">
                    
						<div class="form-group">
							<label class="col-md-12">To *</label>
							<div class="col-md-12">
								<span id="txtTo"></span>
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Subject *</label>
							<div class="col-md-12">
								<input type="text" id="txtsubject" name="txtsubject" class="form-control"  value="Email Timesheet Reminder">
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Message *</label>
							<div class="col-md-12">
								<textarea id="txtMessage" name="txtMessage"  cols="50" rows="10" class="form-control">
								Dear #Name#,								
								</textarea>
							</div>							
							<div class="col-md-12" style="text-align:right; padding:">
								<b>#Name#</b> will be replaced by First Name of recipient
							</div>
						</div>						
						<div class="form-group">
							<div class="col-md-12">
								<div class="col-md-1 no-padding width-auto">
									<input type="checkbox" name="radRecord" id="radRecord" value="1" class="no-padding">
								</div>
								<label class="col-md-3 padding-left5 no-blod" for="radRecord">Recorded</label>
							</div>
						</div>
						<input type="hidden" id="txtlistID" name="txtlistID" value="">
					
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                    <button type="submit" class="btn btn-primary" id="btnSend">Send Email</button>
                </div>
				</form>
            </div>
        </div>
    </div>
    </div>
</div>

<%
	Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(Replace(arrPageTemplate(2),"@@currentYear",Year(Date())))
%>
<script type="text/javascript" src="../../js/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/dataTables.bootstrap.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/dataTables.select.min.js"></script>


<script type="text/javascript" src="../../bootstrap/js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/formValidation.min.js"></script>
<script type="text/javascript" src="../../js/framework/bootstrap.min.js"></script>

<script type="text/javascript" src="../../js/library.js"></script>
<script language="javascript">
<!--
var objNewWindow;
$.extend( true, $.fn.dataTable.defaults, {
     "paging": false,
		"searching": false,
		"processing": true,
		"ordering": false,
		"info":     false
} );
$(document).ready(function() {

    var table = $('#tblList').DataTable();
   
   // Handle click on "Select all" control
	$('#select-all').on('click', function(){
	   // Get all rows with search applied
	   var rows = table.rows({ 'search': 'applied' }).nodes();
	   // Check/uncheck checkboxes for all rows in the table
	   $('input[type="checkbox"]', rows).prop('checked', this.checked);
	});
	
	// Handle click on checkbox to set state of "Select all" control
	$('#tblList tbody').on('change', 'input[type="checkbox"]', function(){
	   // If checkbox is not checked
	   if(!this.checked){
		  var el = $('#select-all').get(0);
		  // If "Select all" control is checked and has 'indeterminate' property
		  if(el && el.checked && ('indeterminate' in el)){
			 // Set visual state of "Select all" control
			 // as 'indeterminate'
			 el.indeterminate = true;
		  }
	   }
	});
	
	$('#btnSendemail').on ('click',function(){
		
		var allVals = [];
		var allNames = [];
		$("input:checked", table.rows().nodes()).each(function(){
			//alert ($(this).val());	
			allVals.push($(this).val());
			allNames.push($(this).parent().next().html());
			//alert ($(this).parent().next().html());						
		});
		if (allNames.length>0)
		{
			$('#txtlistID').val(allVals);
			$('#txtMessage').val("Dear #Name#,");
			document.getElementById("txtTo").innerHTML = allNames;
			$('#myModal').modal({backdrop: "static"});
		}
		
	});
	
	// Align modal when it is displayed
    $(".modal").on("shown.bs.modal", alignModal);
    
    // Align modal when user resize the window
    $(window).on("resize", function(){
        $(".modal:visible").each(alignModal);
    });   
	
	 $('#detailform').formValidation({
		framework: 'bootstrap',
        icon: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },
        fields: {
			txtsubject: {	
                validators: {
                    notEmpty: {
                        message: 'The Subject is required'
                    }
                }
            },
			txtMessage: {	
                validators: {
                    notEmpty: {
                        message: 'The Message is required'
                    }
                }
            }
		}
	 });
});

function alignModal(){
	var modalDialog = $(this).find(".modal-dialog");	
	// Applying the top margin on modal dialog to align it vertically center
	modalDialog.css("margin-top", Math.max(0, ($(window).height() - modalDialog.height()) / 2));
}

function printpage() 
{ //v2.0
var row = "<%=intRow%>";
//	if ("<%=intRow%>" != "" && "<%=intRow%>" >= 0)
	//if (row != "" && row >= 0)
	//{
		window.status = "";
 
		strFeatures = "top=1,left="+(screen.width/2-350)+",width=630,height=680,toolbar=no," 
		          + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";
              
		if((objNewWindow) && (!objNewWindow.closed))
			objNewWindow.focus();	
		else 
		{
			objNewWindow = window.open('rpt_print_preview.asp?title=' + '<%=strtitle2%>', "MyNewWindow", strFeatures);
		}
		window.status = "Opened a new browser window.";  
	//}	
}

//-->
</script>
</body>
</html>
