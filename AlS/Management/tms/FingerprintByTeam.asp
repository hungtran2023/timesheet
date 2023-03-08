<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim intUserID, intMonth, intYear, intDayNum, intWeekday, intRow, intCount,strAct,intNumberWeek
	Dim varFullName, varFrom, varTo, varPre, getRes, varUser, varInvalidTMS,rsEmailCount,rsWeekDate
	Dim strUserName, strTitle, strFunction, strMenu, strURL, strType, strTitle2, strFrom, strTo, strFirstDay, strCurDate, strDateShow
'--------------------------------------------------
'
'--------------------------------------------------	
Function Format_hhmm(dblMinuteIn)
	dim strOut
	dim dblHour, dblMinute
	strOut=""
	if dblMinuteIn<>0 then
		dblHour=cint(abs(dblMinuteIn)\60)
		dblMinute=abs(dblMinuteIn) mod 60
		strOut=dblHour
		if dblHour<10 then strOut="0" & strOut
		strOut=strOut &":" 
		if dblMinute<10 then
			strOut=strOut & "0" & dblMinute
		else
			strOut=strOut & dblMinute
		end if
		if dblMinuteIn<0 then strOut="-" & strOut
	end if
	
	Format_hhmm=strOut
end function	
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
Function GetFingerprintByTeam()
	dim rs, strOut
	dim ii, intColumns
	dim objConn 
	strOut=""
	strConnect = Application("g_strConnect")
	set objConn= Server.CreateObject("ADODB.Connection")
	objConn.Open strConnect  
	
	If objConn.State=1 Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objConn
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "FingerprintByTeam"
		Set myParam = myCmd.CreateParameter("month",adInteger,adParamInput)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("year",adInteger,adParamInput)
		myCmd.Parameters.Append myParam

		myCmd("month") = intMonth
		myCmd("year") = intyear

		SET rs=myCmd.Execute
		strOut=""
		
		if Not rs.EOF then
			intColumns=intNumberWeek*3+1
			do while not rs.EOF	

				strOut=strOut& "<td>" & rs.fields(1) & "</td>"
				strOut=strOut& "<td>" & rs.fields(2) & "</td>"
				for ii=1 to intNumberWeek
			
					strOut=strOut& "<td>" & Format_hhmm(cdbl(rs.fields(2+ii*2-1))) & "</td>"
					strOut=strOut& "<td>" & Format_hhmm(cdbl(rs.fields(2+ii*2))) & "</td>"
					strOut=strOut& "<td><b>" & Format_hhmm(cdbl(rs.fields(2+ii*2))- cdbl(rs.fields(2+ii*2-1))) & "</b></td>"
				next
				'strOut=strOut& "<td>" & strField & "</td>"
				strOut=strOut& "</tr>"
				rs.MoveNext
			loop
		end if

		
	end if
	
	GetFingerprintByTeam=strOut
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
	If strChoseMenu = "" Then strChoseMenu = "AA"

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
	
	<div class="row" style="padding:20px 0px 0px 20px;"><h3>Fingerprint By Team</h3> </div>
	<div class="row">
        <div class="col-sm-6 col-sm-offset-3">
            <div id="imaginary_container"> 
			<form name="frmFilterform" method="post" action="FingerprintByTeam.asp">
				<div class="form-group">
					<div class="input-group stylish-input-group">
						<input type="text" name="txtSearch" id="txtSearch" onkeyup="myFunction()" class="form-control"  placeholder="Filter by Manager" >
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
<%intNumberWeek=NumberOfWeek()%>
	<div class="row" >
		<form id="frmList" method="post" action="FingerprintByTeam.asp">
		<table class="table table-hover" id="tblList">
					<thead  class="thead-inverse tableheaderblue">
						<tr>	
							<th  rowspan="2">Full Name</th>
							<th  rowspan="2">Report To</th>		
<%rsWeekDate.moveFirst
do while not rsWeekDate.EOF%>							
							<th colspan="3">Week <%=rsWeekDate("weekcount")%> <span class="blue-normal">(<%=rsWeekDate("WeekStart")%>-<%=rsWeekDate("WeekEnd")%>)</span></th>	
<%
rsWeekDate.MoveNext
loop
%>
						</tr>						
						<tr>							
							
<%for ii=1 to intNumberWeek%>	
							<th><span style="font-size:8pt">Timesheet</span></th>	
							<th><span style="font-size:8pt">Fingerprint</span></th>	
							<th><span class="blue" style="font-size:8pt">Weekly Balance</span></th>	
<%Next%>							
						</tr>						
					</thead>					
					<tbody>
					<%=GetFingerprintByTeam()%>
					</tbody>
				</table>	
			</form>
	</div>
	
	
	 <!-- Modal HTML -->
   <!-- Modal HTML -->
 
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
   
	});
function myFunction() {
   var input, filter, table, tr, td, i;
  input = document.getElementById("txtSearch");
  filter = input.value.toUpperCase();
  table = document.getElementById("tblList");
  tr = table.getElementsByTagName("tr");

  // Loop through all table rows, and hide those who don't match the search query
  for (i = 0; i < tr.length; i++) {
    td = tr[i].getElementsByTagName("td")[1];
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
