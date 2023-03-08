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

	dim dateFirstTMS, dateLastTMS, rsTimesheet

Function SumTimesheet(dateFrom, dateTo)
	dim totalHour
	totalHour=0

	if rsTimesheet.recordcount>0 then
		rsTimesheet.moveFirst
		do while NOT rsTimesheet.EOF
			if (rsTimesheet("Tdate")>=dateFrom) AND (rsTimesheet("Tdate")<=dateTo) then 
				totalHour=totalHour + cdbl(rsTimesheet("Hours"))+ cdbl(rsTimesheet("Overtime"))
				'if Cdate(dateFrom)="4/6/2019" then	response.write rsTimesheet("Tdate") & "->" & dateTo & ": " &  cdbl(rsTimesheet("Hours"))	& "<br>"		
			end if
			rsTimesheet.MoveNext
		Loop
	
	end if
'if Cdate(dateFrom)="4/6/2019" then	response.write dateFrom & "->" & dateTo & ": " & totalHour	& "<br>"	
	SumTimesheet=totalHour
End function	

'***************************************************************
'
'***************************************************************	
Function NumberOfWeek(intYear,intMonth,ByRef rsWeekDate)
	dim intWeekcount
	dim startDate, endDate

	startDate=DateSerial(intYear, intMonth, 1)
	endDate=DateSerial(intYear, intMonth+1, 1)-1

	strSql="SELECT * FROM MonthToWeek('" & startDate & "','" & endDate & "')"
	

    Call GetRecordset(strSql,rsWeekDate)

	NumberOfWeek=rsWeekDate.recordcount
	
End function

'***************************************************************
'
'***************************************************************	
Function OverviewDataOfAPK()
	
	dim strOut,dblBurn
	strOut=""
	
	strSQL="SELECT SUM(Hours+OTHours) as ActualHours, SUM(InvoiceValue) as Sales, SUM(CSOHours) as CSOHOurs, SUM(CSOPayment) as CSOPayment, SUM(InvoiceValueUSD) as InvoiceValueUSD FROM rp_ProjectPerformanceByPeriod WHERE ProjectID='" & strProID &"'"
	Call GetRecordset(strSql,rsOverPro)
	
	if rsOverPro.recordcount>0 then
	if cdbl(rsOverPro("CSOHOurs"))>0 then dblBurn=(cdbl(rsOverPro("ActualHours"))/cdbl(rsOverPro("CSOHOurs")))*100
	'style='background-color:#C2CCE7'
	strOut="<tr style='background-color:#E7EBF5' ><td></td><td></td><td><b>" & _
			formatnumber(cdbl(rsOverPro("CSOPayment")),2) & "</b></td><td><b>" & _
			formatnumber(cdbl(rsOverPro("Sales")),2) &"</b></td><td><b>" & _
			formatnumber(cdbl(rsOverPro("CSOHOurs")),2) & "</b></td><td><b>" & _
			formatnumber(cdbl(rsOverPro("ActualHours")),2) & "</b></td><td><b>" & _
			"</b></td><td><b>" & _
			"</b></td><td><b>" & _
			"</b></td><tr>"
	
	end if
	
	OverviewDataOfAPK=strOut
End function

'***************************************************************
'
'***************************************************************
Function GetEstDetails(periodMonth, periodYear,startDate, endDate, weekCount, byref dblActualHours)
	dim strOut,rs
	dim dblHours, dblEst, strComment, dblWeekID

	dblHours=SumTimesheet(startDate,endDate)

	dblActualHours=dblActualHours+ cdbl(dblHours)
	dblEst="--"
	strComment=""	
	dblWeekID=-1
	strButtonEst=""
	if endDate<=date then
		strSql="SELECT * FROM ATC_EstPerCompleteByWeek WHERE ProjectID='" & strProID& "' AND periodMonth="& periodMonth &" AND periodYear=" & periodYear & " AND WeekNum=" & weekCount & " ORDER BY WeekNum"	
		Call GetRecordset(strSql,rs)
		'response.write strSql & "<br>"' & rs.recordcount
		dblEst="--"
		strComment=""	
		dblWeekID=-1
		if  rs.recordcount>0 then	
			'dblEst=rs("EstValue")
			dblEst=rs("EstHours")
			strComment=rs("Comments")
			dblWeekID=rs("WeekID")
		end if
		strButtonEst="<button type=button' class='btn btn-default editEst' data-datac='" & periodMonth & "#" & periodYear & "#" & weekCount & "#" & dblWeekID & "#" & endDate & "'>" & dblEst  & "</button>"
	end if 
	
	strOut="<td><table width='100%'><tr><td width='70%'>W" & weekCount & " (" & ddmmyyyy(startDate) & "->" & ddmmyyyy(endDate) & ")</td><td> <b>" & dblHours  & "</b></td></tr></table></td>" &_
					 "<td >"& strButtonEst &"</td>" &_
			         "<td>" & strComment & "</td>" 
	GetEstDetails=strOut	
End function

'***************************************************************
'
'***************************************************************
Function Outbody(ByRef rsSrc)
	dim i,strMonthNameDisplay,intWeeksofMonth
	dim rsWeek, strRowClass
	dim dateWStart, dateWEnd, idxWcount,blnNotFirstRow, dblActualHours
	strOut = ""
	If Not rsSrc.EOF Then
		i=0
		do while not rsSrc.EOF
			i=i+1
			strMonthNameDisplay=monthname(rsSrc("periodMonth"),2) & "-" & rsSrc("periodYear")
			intWeeksofMonth=NumberOfWeek( rsSrc("periodYear"),rsSrc("periodMonth"),rsWeek)
			idxWcount=1
			dblActualHours=0
			blnNotFirstRow=false
			if (i mod 2=0) then 
				strRowClass="even"
			else
				strRowClass="odd"
			end if
			
			strOut = strOut & "<tr class='"& strRowClass &"'>" &_
			         "<td >" & i & ".</td>" &_
			         "<td>" & strMonthNameDisplay  & "</td>" &_
					 "<td>" &  formatnumber(rsSrc("CSOPayment"),2)  & "</td>" &_
			         "<td>" & formatnumber(rsSrc("InvoiceValue"),2) & "</td>" &_
					 "<td>" & formatnumber(rsSrc("CSOHours"),2) & "</td>"
				strOutW=""

			do while not rsWeek.EOF

			 
'				if rsWeek("WeekEnd")>=dateFirstTMS then

					dateWStart=rsWeek("WeekStart")
					dateWEnd= rsWeek("WeekEnd")

					if idxWcount=1 AND dateWStart<> DateSerial(rsSrc("periodYear"),rsSrc("periodMonth") ,1) then dateWStart=DateSerial(rsSrc("periodYear"),rsSrc("periodMonth") ,1)

					if idxWcount= intWeeksofMonth  AND dateWEnd<DateSerial( rsSrc("periodYear"),rsSrc("periodMonth")+1 ,0) then dateWEnd=DateSerial( rsSrc("periodYear"),rsSrc("periodMonth")+1 ,0)
										
					if blnNotFirstRow then 
						strOutW = strOutW & "</tr><tr class='"& strRowClass &"'><td></td><td></td><td></td><td></td><td></td><td></td>"
					else
						blnNotFirstRow=true						
					end if					
					strOutW = strOutW & GetEstDetails( rsSrc("periodMonth"),rsSrc("periodYear"),dateWStart,dateWEnd,rsWeek("WeekCount"),dblActualHours)				
					
					
'				end if
				idxWcount=idxWcount+1
				rsWeek.MoveNext	
			loop
			
			strOut=strOut & 	 "<td>" & formatnumber(dblActualHours,2) & "</td>" & strOutW
						
			strOut=strOut &"</tr>" & chr(13)
					 
			         			
			rsSrc.MoveNext
		loop		
		
	End If
	Outbody = strOut
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
	strProID  = Request.Form("txthidden")
	
'--------------------------------------------------
' Initialize appoval timesheet records
'--------------------------------------------------
	strConnect = Application("g_strConnect")	' Connection string 	
	
	strAct=request.querystring("act")
	if strAct<>"" then
		strTemp=split(Request.Form("weekID"),"#")
		if strAct="save" then			
			dblEst=request.form("estValue")
			strComment="NULL"
			if request.form("comment")<>"" then
				strComment="'" & replace(request.form("comment"),"'","''") & "'"
			end if
			
				if strTemp(3)="-1" then
					strSql="INSERT INTO ATC_EstPerCompleteByWeek ([periodYear],[periodMonth],[WeekNum],[LastDateOfWeek],[EstHours],[Comments],projectID) VALUES (" & _
							strTemp(1) & "," & strTemp(0) & "," & strTemp(2) & ",'" & strTemp(4) & "'," & dblEst & "," & strComment	& ",'"& strProID &  "')"
				else
					strSql="UPDATE ATC_EstPerCompleteByWeek SET " & _
							"[EstHours]=" & dblEst & _
							",[Comments]=" & strComment & _
							" WHERE WeekID=	" & strTemp(3)					 
				end if

		elseif strAct="del" then
			strSql="DELETE ATC_EstPerCompleteByWeek WHERE WeekID=	" & strTemp(3)		
		end if
		Set objDb = New clsDatabase
'response.write strSql		
		If objDb.dbConnect(strConnect) then
				
			ret = objDb.runActionQuery(strSql)
					
			if ret=false then				
				strMessage = objDb.strMessage
			end if
				  
		else
			strMessage=objDb.strMessage
		end if
	end if
	
	strSql="SELECT * FROM [TimesheetByAPK]('"&strProID& "') a INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID INNER JOIN ATC_Tasks c ON b.SubTaskID=c.SubTaskID " & _
			"WHERE c.fgBillable<>0"
	Call GetRecordset(strSql,rsTimesheet)

	strSQL="SELECT  ISNULL(MIN(Tdate),GetDAte()) as firstDate,ISNULL(Max(Tdate),getDate()) as LastDate FROM [TimesheetByAPK]('"&strProID&"') a INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID INNER JOIN ATC_Tasks c ON b.SubTaskID=c.SubTaskID WHERE  c.fgBillable<>0 "
	Call GetRecordset(strSql,rs)
	
	dateFirstTMS=rs("firstDate")	
	dateLastTMS=rs("LastDate")
'response.write dateFirstTMS & "--" & dateLastTMS &"<br>"

	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strConnect) Then
		
		'strSql=" SELECT Period,periodMonth, periodYear,  InvoiceValue, CSOHours, CSOPayment FROM rp_CSOandInvoiceByPeriod WHERE  ProjectID='" & strProID & "' ORDER BY Period "
		strSql="SELECT Period,periodMonth, periodYear,  InvoiceValue, CSOHours, CSOPayment, (Hours+OTHours) as hours_ FROM rp_ProjectPerformanceByPeriod WHERE  ProjectID='" & strProID & "' ORDER BY Period "

		Set rsCSODetails = Server.CreateObject("ADODB.Recordset")
		Set rsCSODetails.ActiveConnection = objDatabase.cnDatabase
		rsCSODetails.CursorLocation = adUseClient
			
		rsCSODetails.LockType=3
			
		rsCSODetails.Open strSQL
			
		If Err.number =>0 then	
			strError = Err.Description
		else
			set rsCSODetails.ActiveConnection=nothing
		end if
	Else
			Response.Write objDatabase.strMessage		
	End If

	strOut=Outbody(rsCSODetails)
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	
	strFunction = "<div align='right'><a href='javascript:back_menu()' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_

				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Timesheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"

	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../../templates/template1/")
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

.tmstable tr.odd{
  background-color: #FFFFFF;
  color: #003399;
}

.tmstable tr.even{
  background-color: #E7EBF5;
  color: #003399;
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
    <div class="row" style="padding:20px 0px 0px 0px;">	<h3>Performance of <b><%=strProID%></b></h3>	</div>
	
	<button class='btn btn-default' onclick="goBack()" > &laquo;&laquo; Go Back</button>
	<p>
	<form class="form-inline"  id="frmSearch" method="post">	
		
		<div class="row">						
			<div id='idWarning'></div>
					<table class="table tmstable" >
						  <thead>
						  
							<tr>
							  <th scope="col">No.</th>
							  <th scope="col">Month</th>
							  <th scope="col">Invoice Schedule</th>
							  <th scope="col">Invoice</th>
							  <th scope="col">CSO Hours</th>
							  <th scope="col">Actual Hours</th>
							  <th scope="col">Detail of Burn Hours</th>
							  <th scope="col">Est. EOJ hours</th>
							  <th width='30%'>Comments</th>			  
							</tr>
						  </thead>
						  <tbody>
								<%=OverviewDataOfAPK()%>
								<%=strOut%>
						  </tbody>
					</table>				 
					 
		</div>
	</form>
</div>
		</td>
	</tr>
</table>

<!-- The Modal -->
  <div class="modal" id="myModal" role="dialog">
    <div class="modal-dialog">
      <div class="modal-content">
		<form id="frmEst" method="post" action="ProjectPerformanceDetails.asp">
			<!-- Modal Header -->
			<div class="modal-header">
			<button type="button" class="close" data-dismiss="modal">&times;</button>
			  <h4 class="modal-title">Percent of completion</h4>
			  
			</div>
			
			<!-- Modal body -->
			<div class="modal-body">
				
					<div class="form-group">
					  <label for="estValue">% of completion:</label>
					  <input type="text" class="form-control" id="estValue"  placeholder="Enter % of completion" name="estValue">
					  <span class="help-block" id="errEstValue" style="display:none"></span>				  
					</div>
					
					<div class="form-group">
					  <label for="comment">Comment:</label>
					  <textarea class="form-control" rows="5" id="comment" name="comment"></textarea>
					</div>
					<input type="hidden" id="weekID" name="weekID" value="">
					<input type="hidden" name="txthidden" id="txthidden" value="<%=strProID%>">
			</div>		
			<!-- Modal footer -->
			<div class="modal-footer">
			  <button type="submit" class="btn btn-default" id="submitEst">Save</button>
			  <button type="button" class="btn btn-default" id="delEst">Delete</button>
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancel</button>
			</div>
        </form>
      </div>
    </div>
  </div>
  
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>

<script type="text/javascript" src="../../jQuery/jquery-3.2.1.min.js"></script>
<!--<script type="text/javascript" src="../../js/bootstrap.min.js"></script>-->
<script type="text/javascript" src="../../bootstrap/js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/bootstrap-datepicker.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/bootstrap-confirmation.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>

<script language="javascript">
<!--


$(document).ready(function() {
 

	$('#submitEst').click(function(e){
		e.preventDefault(); 	
		
		if (validateEst()==true)
		{
			$("#frmEst").attr('action', 'ProjectPerformanceDetails.asp?act=save').submit();
		}
		else
		{
			$("#estValue").parent().addClass("has-error");
			$("#errEstValue").html("This value is require and must be a number.")
			$("#errEstValue").show();
		}
	});
	
	$('.editEst').click(function(e) {

		e.preventDefault(); 	   
		var strKey=$(this).attr("data-datac");
		var arrTemp=strKey.split("#");
	   
		$("#weekID").val(strKey);
		if (arrTemp[3]!="-1")
		{
			$("#estValue").val($(this).text().replace("%",""));
			$("#comment").val($(this).closest("td").next().html());
		}
		else
		{ 
			$("#estValue").val("");
			$("#comment").val("");
		}
		$("#estValue").parent().removeClass("has-error");
		$("#errEstValue").html("")
		$("#errEstValue").hide();
		
		$('#myModal').modal({backdrop: "static"});

    });
	
	$('#delEst').click(function(e) {
		if (confirmationBox())
		{
			$("#frmEst").attr('action', 'ProjectPerformanceDetails.asp?act=del').submit();
		}
	});
	
	$("#estValue").keyup(function(e){
        /* Ignore tab and enter key */
        var code = e.keyCode || e.which;
	
        if (code == '9' || code == '13') return;
		var $input = $(this);
		if ((e.keyCode >=48 && e.keyCode <=57) || (e.keyCode >=96 && e.keyCode <=105))
		{
			$input.parent().removeClass("has-error");
			$("#errEstValue").html("")
			$("#errEstValue").hide();
			return;
		}
		else
		{			
			$input.parent().addClass("has-error");
			$("#errEstValue").html("This value must be a number.")
			$("#errEstValue").show();
		}
       
    });	
		
});


function confirmationBox() {
    var r = confirm("You are about to delete this item. Click Ok to continue");
        if (r) {
            return true;
        }
        else {
            return false;
        }
    }
function validateEst()
{
	var $input = $("#estValue");	
	return $.isNumeric($input.val());
}

function back_menu(){
  var url;
	url = "ProjectPerformance.asp";	
	document.location = url;
}

function goBack(){
 //window.history.back();
 document.location = "ProjectPerformance.asp";
}


//-->
</script>
</body>
</html>
