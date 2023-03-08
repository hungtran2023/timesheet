<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%
dim strCompany, strJobtitle, strLocation,strStartDate, strEndDate, strNote
Dim dblEmpHistoryID
Dim strUserid, rsSrc

function PopulateListEmpHistory
	
	Dim strSql,strReturn,i
	Dim objDatabase, rs
	Dim strConnect, strID,  strError,strEntryType, strLastdateTemp
	
	strSql="SELECT EmpHistoryID , StaffID , CompanyName , Location , Jobtitle , StartDate , EndDate , HistoryNote FROM ATC_EmployeeHistory  " &_
            " WHERE StaffID=" & strUserid & " ORDER By StartDate"
    
    Call GetRecordset(strSql,rs)  
    strReturn=""
    i=1
    if not rs.Eof then
	  Do Until rs.EOF
	      strID=rs("EmpHistoryID")
	      
	      if cdbl(strID)=cdbl(dblEmpHistoryID) then
                strCompany=rs("CompanyName")
                strJobtitle=rs("Jobtitle")
                strLocation=rs("Location")
                strStartDate=rs("StartDate")
                if strStartDate<>"" then strStartDate=day(strStartDate) & "/" & month(strStartDate) & "/" & year(strStartDate) 
                strEndDate=rs("EndDate")
                if strEndDate<>"" then strEndDate=day(strEndDate) & "/" & month(strEndDate) & "/" & year(strEndDate) 

                strNote=rs("HistoryNote")
	      end if
     
          strLastdateTemp=day(rs("StartDate")) & "/" & month(rs("StartDate")) & "/" & year(rs("StartDate")) & "-" & day(rs("EndDate")) & "/" & month(rs("EndDate")) & "/" & year(rs("EndDate"))
          
	      strReturn= strReturn & "<tr idValue='"  & strID & "'><td>" & i & ".</td>"
          strReturn= strReturn & "<td class='editrow'>" & rs("CompanyName") & "</td>"
          strReturn= strReturn & "<td class='editrow'>" & rs("Jobtitle") & "</td> "
          strReturn= strReturn & "<td class='editrow'>" &  strLastdateTemp &  "</td> "
          strReturn= strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item' data-id='" & rs("EmpHistoryID")  & "'></button></td>  </tr>"		
	    rs.MoveNext
	    i=i+1
	  Loop       
	end if         
	
	PopulateListEmpHistory=strReturn
end function


'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If
'-----------------------------------
'Check ACCESS right
'-----------------------------------
	
	tmp = Request.Form("txtpreviouspage")
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	fgRight=true
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if	
'--------------------------------------------------

'--------------------------------------------------

strUserid= Request.querystring("id")
if strUserid="" then strUserid=Request.Form("txtuserid")
'response.write strUserid

strAct= Request.querystring("act")


if strAct="u" or strAct="d" then

    if strAct="u" then
    
        strCompany=request.form("txtCompany")
        strJobtitle=request.form("txtJobTitle")
        strLocation=request.form("txtLocation")
        strStartDate=request.form("txtStart")
        if strStartDate<>"" then strStartDate=ConvertTommddyyyy(strStartDate)
        strEndDate=request.form("txtEndDate")
        if strEndDate<>"" then strEndDate=ConvertTommddyyyy(strEndDate)
        
        strNote=request.form("txtNote")
        
        dblEmpHistoryID=request.form("txtEmpHistoryID")
        
        if dblEmpHistoryID="" then dblEmpHistoryID=-1   
          
        if cint(dblEmpHistoryID)=-1 then
            strSql="INSERT INTO ATC_EmployeeHistory (StaffID,CompanyName , Location , Jobtitle , StartDate , EndDate , HistoryNote ) VALUES (" & _
                    strUserid & ",'" & strCompany & "','" & strLocation & "','" & strJobtitle & "','" & strStartDate & "','" & strEndDate & "'," & _
                    IIF(strNote="", "null","'" & strNote & "'") & ")"
        else
            strSql="UPDATE  ATC_EmployeeHistory  " & _
                        "SET CompanyName  = '" & strCompany & "'" & _
                            ", Location  =  '" & strLocation & "'" & _
                            ", Jobtitle  = '" & strJobtitle & "'" & _
                            ", StartDate  = '" & strStartDate & "'" & _
                            ", EndDate  = '" & strEndDate & "'" & _
                            ", HistoryNote = '" & strNote & "'" & _
                    " WHERE EmpHistoryID=" & dblEmpHistoryID
        end if
    else
        strSql="DELETE FROM ATC_EmployeeHistory WHERE EmpHistoryID=" & Request.querystring("subid")
    end if
	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase     
    strError=""
    
	If objDatabase.dbConnect(strCnn) Then              
		if not objDatabase.runActionQuery(strSql) then 
		   gMessage = objDatabase.strMessage
        end if			  
    end if
    
    strCompany=""
    strJobtitle=""
    strLocation=""
    strStartDate=""
    strEndDate=""
      
    strNote=""
    
    dblEmpHistoryID=-1
    
else
    dblEmpHistoryID=strAct
end if

 strListEmpHistory=PopulateListEmpHistory()
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
	
	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AB"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'----------------------------------------
' analyse query string
'----------------------------------------
	gMessage = ""

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
                <a class="blue" href="listofemployee.asp" onMouseOver="self.status='Show the list of employees'; return true;" onMouseOut="self.status=''">Employee List:</a>
            <span>Employee Details</span>
            </div>
        </td>
    </tr>
</tbody>
</table>
<div class="container-fluid">
    <div class="row">
        <div class="col-sm-12">
            <ul class="nav nav-tabs">
                <li><a href="employeeProfile.asp?id=<%=strUserid%>">Employee Profile</a></li>
                <li><a href="atlasinformation.asp?id=<%=strUserid%>">Atlas Information</a></li>
                <li><a href="educationskill.asp?id=<%=strUserid%>">Education/Skill</a></li>
                <li><a href="replacementhistory.asp?id=<%=strUserid%>">Replacement History</a></li>
                <li class="active"><a>Employment History</a></li>
            </ul>
        </div>
    </div>
    <div class="row">
        <div class="col-sm-12">
            <div class="tab-content employee-details-form">
                <div class="row">
                    <div class="col-md-12 col-sm-6 col-xs-12">
                        <form id='frmHistory' class="form-horizontal row-border" method="POST" action="employmenthistory.asp?id=<%=strUserid%>&act=u">
<%if gMessage<>"" then%>   	 
                            <div class="alert alert-danger">
                                <strong>Error:</strong> <%=gMessage%>
                            </div>
<%end if%> 
                            <div class="panel panel-default">
                                <div class="panel-body">
                                    <div class="col-sm-6">
                                        <div class="form-group">
                                            <label class="col-md-12">Company</label>
                                            <div class="col-md-12">
                                                <input type="text" id="txtCompany" name="txtCompany" class="form-control" value="<%=strCompany%>" >
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Location (City)</label>
                                            <div class="col-md-12">
                                                <input type="text" id="txtLocation" name="txtLocation" class="form-control" value="<%=strLocation%>">
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Job Title</label>
                                            <div class="col-md-12">
                                                <input type="text" id="txtJobTitle" name="txtJobTitle" class="form-control" value="<%=strJobTitle%>">
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-sm-6">
                                        <div class="form-group">
                                            <label class="col-md-12">Start Date</label>
                                            <div class="col-md-12">
                                                <div class="input-group date">
                                                    <input type="text" class="form-control datepicker" id="txtStart" name="txtStart" placeholder="DD/MM/YYYY" value="<%=strStartDate%>">
                                                    <span class="input-group-addon">
                                                        <span class="ic-calendar"></span>
                                                    </span>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">End Date</label>
                                            <div class="col-md-12">
                                                <div class="input-group date">
                                                    <input type="text" class="form-control datepicker" id="txtEndDate" name="txtEndDate" placeholder="DD/MM/YYYY" value="<%=strEndDate%>">
                                                    <span class="input-group-addon">
                                                        <span class="ic-calendar"></span>
                                                    </span>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-sm-12">
                                        <div class="form-group">
                                            <label class="col-md-12">Note</label>
                                            <div class="col-md-12">
                                                <textarea  id="txtNote" name="txtNote"  cols="30" rows="3" class="form-control" ><%=strNote%></textarea>
                                            </div>
                                        </div>
                                    </div>
                                    <input type="hidden" name="txtEmpHistoryID" id="txtEmpHistoryID" value="<%=dblEmpHistoryID%>"/>
                                    <div class="col-sm-12">
                                        <div class="form-group text-right">
                                            <div class="col-md-12 no-padding">
                                                <button type="submit" class="btn btn-primary" id="btnSave">Save</button>
                                        <button type="button" class="btn btn-primary <%if cdbl(dblEmpHistoryID)<=0 then%> hide<%end if%>" id="btnCancel">Cancel</button>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-md-12">
                                        <table class="table table-striped table-bordered table-hover margin-bottom10 table-responsive" id="tblList">
                                            <thead class="thead-inverse">
                                                <tr>
                                                    <th class="col-md-1 col-md-checkbox"></th>
                                                    <th>Company</th>
                                                    <th>Job Title</th>
                                                    <th>Duration</th>
                                                    <th class="col-action"></th>
                                                </tr>
                                             </thead>
                                             <tbody>
                                                <%=strListEmpHistory%>                       
                                            </tbody>
                                        </table>
                                    </div>
                    
                                </div>
                            </div>            
                         </form>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
<!-- Modal for displaying the messages -->
<div class="modal fade" id="confirm-delete" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                  <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                    <h4 class="modal-title" id="myModalLabel">Confirm Delete</h4>
                </div>
            
                <div class="modal-body">
                    <p id="modal_message"></p>
                    <p>Do you want to proceed?</p>
                </div>
                
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
                    <a class="btn btn-danger btn-ok">Delete</a>
                </div>
        </div>
    </div>
</div>      
               
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
<script type="text/javascript" src="../../js/bootstrap-table.js"></script>
<script type="text/javascript" src="../../js/js-control.js"></script>
<script type="text/javascript" src="../../js/formValidation.min.js"></script>
<script type="text/javascript" src="../../js/framework/bootstrap.min.js"></script>

<script type="text/javascript">

$(document).ready(function() {

         
    $('#txtStart').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmHistory').formValidation('revalidateField', 'txtStart');
        });
        

    $('#txtEndDate').on('changeDate', function(e) {
            $('#frmHistory').formValidation('revalidateField', 'txtEndDate');
        });


    $('#frmHistory')
        .formValidation({
            framework: 'bootstrap',
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },        
            fields: {
                txtCompany:{
                    validators: {
                        notEmpty: {
                            message: 'The Company name is required.'
                        }
                    }
                },
                txtLocation:{
                    validators: {
                        notEmpty: {
                            message: 'The Location is required.'
                        }
                    }
                },
                txtJobTitle:{
                    validators: {
                        notEmpty: {
                            message: 'The JobTitle is required.'
                        }
                    }
                },
                txtStart: {
                    validators: {
                        notEmpty: {
                            message: 'The start date is required'
                        },
                        date: {
                            format: 'DD/MM/YYYY',
                            max: 'txtEndDate',
                            min:'1/1/1900',
                            message: 'The start date is not a valid or the Start Date must be before the End Date.'
                        }
                    }
                },
                txtEndDate: {
                    validators: {
                        notEmpty: {
                            message: 'The end date is required'
                        },
                        date: {
                            format: 'DD/MM/YYYY',
                            min: 'txtStart',
                            max: '<%=Date()+365*10%>',
                            message: 'The end date is not a valid or the End Date must be after the Start Date.'
                        }
                    }
                }
               
                
            }
    })
     .on('success.field.fv', function(e, data) {
            if (data.field === 'txtStart' && !data.fv.isValidField('txtEndDate')) {
                // We need to revalidate the end date
                data.fv.revalidateField('txtEndDate');
            }

            if (data.field === 'txtEndDate' && !data.fv.isValidField('txtStart')) {
                // We need to revalidate the start date
                data.fv.revalidateField('txtStart');
            }
        });
    
   $(".editrow").on('click', function(e) {
   
        var res = $(this).parent().attr("idValue");
        window.location = 'employmenthistory.asp?id=<%=strUserid%>&act='+ res  ; // redirect
   });
   
   $("#btnCancel").click(function(){
       
            $("#txtEmpHistoryID").val(-1);
            $("#txtCompany").val("");
            $("#txtLocation").val("");
            $("#txtJobTitle").val("");
            $("#txtStart").val("");
            $("#txtEndDate").val("");
            $("#txtNote").val("");

            $(this).addClass('hide');
        });
        
        
        $('.btn-remove-item').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to remove this duration.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'employmenthistory.asp?id=<%=strUserid%>&act=d&subid='+id );
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        });      
});

</script>

</body>
</html>

