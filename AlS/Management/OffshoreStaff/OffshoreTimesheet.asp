<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
strStatus= Request.Form("txtstatus")
strUserid=request.querystring("id")
if strUserid="" then strUserid= Request.Form("txtuserid")

'if strStatus="submit" then


'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
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

	strSql="SELECT * FROm ATC_Events WHERE EventID NOT IN (1,2,10) "	
	Call GetRecordset(strSql,rsEvents)
	strEvent= PopulateDataToListWithoutSelectTag(rsEvents,"EventID", "EventName","-1")
'----------------------------------
' Get Full Name and Job Title
'----------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
'----------------------------------	
' Make list of menu
'----------------------------------
	If isEmpty(session("Menu")) then 
		getRes = getarrMenu(session("USERID"))
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End if	
	'0000AP0TES
	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AI"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'----------------------------------------
' analyse query string
'----------------------------------------
	
	'strUserid=Request.Form("txtUserid")	
    'if strUserid="" then strUserid=-1 '--Add new
    
	strAct = Request.QueryString("act")  
    
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
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

<title>Atlas Industries Time Sheet System</title>
   
    <link href="../../css/bootstrap.min.css" rel="stylesheet" type="text/css">
    <link href="../../css/timesheet.css" rel="stylesheet" >     
    <link href="../../css/atlasJquery.css" rel="stylesheet" type="text/css" />
    <link href="../../css/style.css" rel="stylesheet" type="text/css">
    <link href="../../css/datepicker.css" rel="stylesheet" type="text/css">    
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
<!--Content-->			

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tbody>
    <tr> 
        <td style="padding:20px 20px 20px 15px;"> 
        <%if gMessage<>"" then%>
            <div style="font-weight:bold; height:20px; background-color:#E7EBF5;" class="red"><%=gMessage%></div>
        <%end if%>
        <div class="navi-info"> 
                <a class="blue" href="listofOffshoreStaffs.asp" onMouseOver="self.status='Show the list of employees'; return true;" onMouseOut="self.status=''">Offshore Staffs List:</a>
            <span>Timesheet</span>
            </div>
        </td>
    </tr>
</tbody>
</table>

<div class="container-fluid" >
	<div class="row">
        <div class="col-sm-12">
            <ul class="nav nav-tabs">
                <ul class="nav nav-tabs">
                <li><a href="OffshoreStaffProfile.asp?id=<%=strUserid%>"> Employee Profile</a></li>
                <li><a href="OffshoreAtlasInformation.asp?id=<%=strUserid%>">Atlas Information</a></li>
				<li class="active"><a>Enter Timesheet</a></li>
            </ul>
            </ul>
        </div>
    </div>
	<div class="row">
		<div class="col-xs-12">
	
		<form class="form-horizontal row-border" id="contactForm" method="POST" >
<%if gMessage<>"" then%>                        
			<div id="messages" class="alert alert-danger">
				<strong>Error:</strong> <%=gMessage%>.
			</div>
<%end if%>
            		
			<div class="panel panel-default">
				 <div class="panel-body">				 
					<div class="form-group">
						<label class="col-xs-4 control-label">Input hours for </label>
						<div class="col-xs-4">
							<select class="form-control input-sm" data-val="true" data-val-number="The field Type Of Absence must be a number." data-val-required="The Type Of Absence field is required." id="AbsenceType" name="AbsenceType">
								<%=strEvent%>
							</select>
						</div>
					</div>
					<div class="form-group">
						<label class="col-xs-4 control-label">From Date</label>
						<div class="col-xs-4">							
							<div class="input-group date" id="StartDatePicker">
								<input class="form-control datepicker" data-val="true" data-val-required="The From Date field is required." id="StartDate" name="StartDate" placeholder="DD/MM/YYYY" type="text" value="" />
								<span class="input-group-addon">
									<span class="ic-calendar"></span>
								</span>
							</div>
							
						</div>
					</div>
					
					<div class="form-group">
						<label class="col-xs-4 control-label">To Date</label>
						<div class="col-xs-4">
							<div class="input-group date" id="EndDatePicker">
								<input class="form-control datepicker" data-val="true" data-val-required="The To Date field is required." id="EndDate" name="EndDate" placeholder="DD/MM/YYYY" type="text" value="" />
								<span class="input-group-addon">
									<span class="ic-calendar"></span>
								</span>
							</div>
						</div>
					</div>
					
					<div class="form-group">
						<label class="col-xs-4 control-label">Hours</label>
						<div class="col-xs-4">
							<input class="form-control" data-val="true" data-val-number="The field Hours must be a number." data-val-required="The Hours field is required." id="Hours" name="Hours" placeholder="0" rows="4" type="text" value="8.50" />
						</div>
					</div>
					<div class="form-group">
						<label class="col-xs-4 control-label">Note</label>
						<div class="col-xs-4">
							<textarea class="form-control" cols="20" id="Note" name="Note" rows="4"></textarea>
						</div>
					</div>
					<div class="form-group">
						<div class="col-xs-8 text-right button-control">
							<button id="view-timesheet" class="btn btn-default">View Timesheet</button>
							<button id="submit-timesheet" class="btn btn-primary">Submit</button>
						</div>
					</div>
				</div>
			</div>
		<input type="hidden" name="txtuserid" value="<%=strUserid%>"/>
		</form>
		</div>
	</div>
</div>
	<!--<div class="row" style="background-color:blue;">
    <div class="col-xs-4 col-xs-offset-2" style="background-color:lavender;">.col</div>
    <div class="col-xs-4" style="background-color:orange;">.col</div>-->
<%
Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
Response.Write(arrPageTemplate(2))    
%>
<script type="text/javascript" src="../../js/jquery.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>
<script type="text/javascript" src="../../js/bootstrap-datepicker.js" charset="UTF-8"></script>
<script type="text/javascript" src="../../js/js-control.js"></script>
<script type="text/javascript" src="../../js/formValidation.min.js"></script>
<script type="text/javascript" src="../../js/framework/bootstrap.min.js"></script>

<script type="text/javascript">

$(document).ready(function() {
//alert ("test");
	 $('#EndDate')
        .on('changeDate', function(e) {
            // Revalidate the date field
            $('#contactForm').formValidation('revalidateField', 'EndDate');
        });
    $('#StartDate')
    .on('changeDate', function(e) {
        // Revalidate the date field
        $('#contactForm').formValidation('revalidateField', 'StartDate');
    });
	
	$('#contactForm').formValidation({
        framework: 'bootstrap',
        button: {
            selector: '#btnNext',
            disabled: 'disabled'
        },
        icon: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },        
        fields: {
            StartDate      : {
                validators: {
                    notEmpty: {
                        message: 'The From date is required'
                    },
                    date: {
                        format: 'DD/MM/YYYY',
                        message: 'The From date is not a valid'
                    }
                }
            },
             EndDate:{
                validators: {
					notEmpty: {
                        message: 'The To date is required'
                    },
                     date: {
						format: 'DD/MM/YYYY',
                        min: 'StartDate',
                        max: '<%=Date()+365*10%>',
                        message: 'The end date is not a valid or the End Date must be after the Start Date.'
                    }
               }       
                    
            },
			Hours:{
				validators: {
                        numeric: {
                            message: 'The value is not a number',
                            // The default separators
                            thousandsSeparator: '',
                            decimalSeparator: '.'
                        }
                    }
			}
			
        }
    })
	
	.on('success.form.fv', function(e) {
            var $form        = $(e.target),     // Form instance
            $statusField = $form.find('[name="txtstatus"]');
            $statusField.val('submit');           
   }); //end of validation
   
})
  
</script>

</body>
</html>

