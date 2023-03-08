<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->

<%
Dim strFirstname,strLastname,strGender,strDateofBirth,strCardNumber,strPassport,strExpireddate
Dim strNationality,strTaxcode,strInsuranceBookNo,strPrivateEmail,strMarialstatus,strBankaccount1,strBankaccount2
Dim strCurrentaddress,strHomeaddress,strMobilephone,strContactVietnam,strHomecountry,strEmergency
Dim strUserid, rsSrc


'***************************************************************
'
'***************************************************************
function PromotionList(rsSrc)
    dim strReturn
    dim i
   if rsSrc.Recordcount>0 then
        strReturn="<tbody>"
        i=1
        Do while not rsSrc.EOF
            
            dblJobTitleID=rsSrc("JobtitleID") 
            strApplyFrom=day(rsSrc("ApplyFrom")) & "/" & month(rsSrc("ApplyFrom")) & "/" & Year(rsSrc("ApplyFrom"))
           
            dblPromotionID=rsSrc("PromotionID")
            
            strReturn=strReturn & "<tr idValue='"  & rsSrc("PromotionID") & "'><td class='col-md-1 col-md-checkbox'>" & i &".</td>"
            strReturn=strReturn & "<td class='jobtitle' jobid='" & rsSrc("JobtitleID") & "'>" & rsSrc("Jobtitle") & "</td>"
            strReturn=strReturn & "<td class='applyfrom' >" & strApplyFrom & "</td>"
            
            rsSrc.MoveNext
            if not rsSrc.EOF then 
                strReturn=strReturn & "<td>" & day(rsSrc("ApplyFrom")-1) & "/" & month(rsSrc("ApplyFrom")-1) & "/" & Year(rsSrc("ApplyFrom")-1)  & "</td>"
                
            else
				if strLastDate<>"" then
					strReturn=strReturn & "<td>"& strLastDate & "</td>"
				else
					strReturn=strReturn & "<td>... Now ...</td>"
				end if
            end if
            if strApplyFrom<>strStartDate then
                strReturn=strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item' data-id='" & dblPromotionID  & "'></button></td></tr>"
            else 
                strReturn=strReturn & "<td class='col-sm-1 col-action text-center'></td></tr>"
            end if
            i=i+1
        loop
         strReturn=strReturn & "</tbody>"
   end if
   
   PromotionList=strReturn
                                             
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
strStatus= Request.Form("txtstatus")
strUserid=request.querystring("id")
if strUserid="" then strUserid= Request.Form("txtuserid")


    strSql= "SELECT a.FirstName, a.LastName,a.Gender,a.Birthday,c.Nationality,b.JoinDate,b.LeaveDate,e.FirstName + ' ' + e.LastName as ReportTo,d.Department FROM ATC_PersonalInfo a " & _
                " INNER JOIN ATC_Employees b on a.PersonID=b.StaffID " & _
	            "INNER JOIN ATC_Countries c ON a.NationalityID=c.CountryID  " & _
	            "INNER JOIN ATC_Department d ON b.DepartmentID=d.DepartmentID " & _
	            "LEFT JOIN ATC_PersonalInfo e ON b.DirectLeaderID=e.PersonID " & _
                " WHERE a.PersonID=" & strUserid
	

    Call GetRecordset(strSql,rsSrc)
      	
    if rsSrc.RecordCount>0 then
    
    
        strFirstname=rsSrc("FirstName")
        strLastname=rsSrc("LastName")
        strGender=iif(RsSrc("Gender"),1,0)
        
        strDateofBirth=day(rsSrc("Birthday")) & "/" & month(rsSrc("Birthday")) & "/" & year(rsSrc("Birthday"))             
      
        strNationality=rsSrc("Nationality")
        strStartDate=day(rsSrc("JoinDate")) & "/" & month(rsSrc("JoinDate")) & "/" & Year(rsSrc("JoinDate"))
    'strStartDate=ConvertToddmmyyyy(rsSrc("JoinDate"))
    
        strLastDate=rsSrc("LeaveDate")
        if strLastDate<>"" then
            strLastDate=day(rsSrc("LeaveDate")) & "/" & month(rsSrc("LeaveDate")) & "/" & Year(rsSrc("LeaveDate"))
        end if
        strDepartment=rsSrc("Department")
        stReportTo=rsSrc("ReportTo")
    
    end if    
    
    strSql="SELECT * FROM ATC_Promotion a INNER JOIN ATC_JobTitle b ON a.JobTitleID=b.JobTitleID WHERE StaffID=" & strUserid & " ORDER BY ApplyFrom"
    Call GetRecordset(strSql,rsSrc)

    strPromotionList=PromotionList(rsSrc)          
'--------------------------------------------------
' Photo
'--------------------------------------------------	    
	strPhoto="male"
	if strGender=0 then strPhoto="female"
	
	strSql = "SELECT  ISNULL(b.Photo,b.Username) as Username FROM ATC_PersonalInfo a LEFT JOIN ATC_Users b ON a.PersonID = b.UserID " &_
				"WHERE PersonID=" & strUserid
				
	Call GetRecordset(strSql,rsSrc)
	if rsSrc.RecordCount>0 then strPhoto=rsSrc("Username")

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
    
	'strAct = Request.QueryString("act")  
    
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
     <link href="../../css/datepicker.css" rel="stylesheet" type="text/css">

     <link href="../../css/atlasJquery.css" rel="stylesheet" type="text/css" />
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
<!--Content-->			

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tbody>
    <tr> 
        <td style="padding:20px 20px 20px 15px;"> 
        <%if gMessage<>"" then%>
            <div style="font-weight:bold; height:20px; background-color:#E7EBF5;" class="red"><%=gMessage%></div>
        <%end if%>
        <div class="navi-info"> 
                <a class="blue" href="listofretired.asp" onMouseOver="self.status='Show the list of employees'; return true;" onMouseOut="self.status=''">Retired Employees List</a>
            </div>
        </td>
    </tr>
</tbody>
</table>

<div class="container-fluid">
<!-- Tab functions -->
    
    <div class="row">
        <div class="col-sm-12">
            <div class="tab-content employee-details-form">
                <div class="row">
                    <div class="col-md-12 col-sm-6 col-xs-12">
                        <form class="form-horizontal row-border" id="contactForm" method="POST" >
<%if gMessage<>"" then%>                        
                            <div id="messages" class="alert alert-danger">
                                <strong>Error:</strong> Indicates a dangerous or potentially negative action.
                            </div>
<%end if%>
                            <div class="panel panel-default">
                                <div class="panel-heading clearfix">
                                    <i class="icon-calendar"></i>
                                    <h3 class="panel-title">Personal Information</h3>
                                </div>
                                <div class="panel-body">
                                    <div class="col-sm-6">
										<div class="row">
											<div class="col-md-3">
												<img height="130" hspace="0" src="http://ais.atlasindustries.com/staff/images/<%=strPhoto%>.jpg" width="100"
												onerror="this.src='http://ais.atlasindustries.com/staff/images/<%if cint(strGender)=0 then response.Write "Fe" end if%>male.jpg';">
											</div>
											<div class="col-md-9">
												<div class="form-group">
													<label class="col-md-12">First name</label>
													<div class="col-md-12">
														<input type="text" id="txtFirstName" name="txtFirstName" class="form-control" placeholder="Enter first name" value="<%=strFirstname%>">
													</div>
												</div>												
												<div class="form-group">
													<label class="col-md-12">Last name</label>
													<div class="col-md-12">
														<input type="text" id="txtLastName" name="txtLastName" class="form-control" placeholder="Enter last name" value="<%=strLastname%>">
													</div>
												</div>
												<div class="form-group">
													<label class="col-md-12">Gender</label>
													<div class="col-md-12">
														<div class="col-md-1 no-padding width-auto">
															<input type="radio" name="radGender" id="radMale" value="1" class="no-padding" <%if cint(strGender)=1 then%>checked<%end if%>>
														</div>
														<label class="col-md-3 padding-left5 no-blod" for="radMale">Male</label>
														<div class="col-md-1 no-padding width-auto">
															<input type="radio" name="radGender" id="radFemale" value="0"  class="no-padding" <%if cint(strGender)=0 then%>checked<%end if%>>
														</div>
														<label class="col-md-3 padding-left5 no-blod" for="radFemale">Female</label>
													</div>
													 <div id="radGenderMessage" class="col-md-12" style="margin-bottom:0px;"></div>
												</div>
											</div>
                                        </div>

									</div>
                                
									<div class="col-sm-6">
										<div class="form-group">
												<label class="col-md-12">Date of Birth</label>
												<div class="col-md-12">
													<div class="input-group date">
															<input type="text" id="txtBirthDate" name="txtBirthDate" class="form-control " placeholder="DD/MM/YYYY"  value="<%=strDateofBirth%>" >
															<span class="input-group-addon">
																<span class="ic-calendar"></span>
															</span>
														</div>
												</div>
									    </div>
                                        
											<div class="form-group">
												<label class="col-md-12">Nationality</label>
												<div class="col-md-12">
													<input type="text" class="form-control" value="<%=strNationality%>" />													
												</div>
											</div>
									 
										
									</div>
								</div>
								<div class="panel-heading clearfix">
									<i class="icon-calendar"></i>
									<h3 class="panel-title">Atlas Information</h3>
								</div>
								<div class="panel-body">
									<div class="col-sm-6">
										<div class="form-group">
				                            <label class="col-md-12">Start Date</label>
				                            <div class="col-md-12">
				                                <div class="input-group date">
				                                    <input type="text"  id="txtStartDate" name="txtStartDate" class="form-control " placeholder="DD/MM/YYYY"  value="<%=strStartDate%>">
				                                    <span class="input-group-addon">
				                                        <span class="ic-calendar"></span>
				                                    </span>
				                                </div>
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Last Date</label>
				                            <div class="col-md-12">
				                                <div class="input-group date">
				                                    <input type="text"  id="txtLastDate" name="txtLastDate" class="form-control " placeholder="DD/MM/YYYY"  value="<%=strLastDate%>">
				                                    <span class="input-group-addon">
				                                        <span class="ic-calendar"></span>
				                                    </span>
				                                </div>
				                            </div>
				                        </div>
										
									</div>
															  
									<div class="col-sm-6">
										<div class="form-group">
				                            <label class="col-md-12">Department</label>
				                            <div class="col-md-12">
				                                <input type="text" class="form-control" value="<%=strDepartment%>">
				                            </div>
				                        </div>
                                         <div class="form-group">
				                            <label class="col-md-12">Report To</label>
				                            <div class="col-md-12">
				                                <input type="text" class="form-control" value="<%=stReportTo%>">
				                            </div>
				                        </div>
									</div>
								</div>
							</div>
                            <div class="col-md-12">
                                        <table class="table table-striped table-bordered table-hover " id="tblListJobtitle">
                                            <thead class="thead-inverse">
                                                <tr>
                                                    <th class="col-md-1 col-md-checkbox">
                                                        No.</th>
                                                    <th>Job Title</th>
                                                    <th>From</th>
                                                    <th>To</th>
                                                    <th class="col-action"></th>
                                                </tr>
                                            </thead>
                                             <%=strPromotionList%> 
                                        </table>
                                    </div>				
							<div class="col-sm-12">
								<div class="form-group text-right">
								
									<button type="button" id="btnCancel" class="btn btn-default">Close</button>
								</div>
							</div>
							<input type="hidden" name="txtuserid" id="txtuserid" value="<%=strUserid%>"/>
							<input type="hidden" name="txtstatus" value="<%=strStatus%>"/>
						</form>
					</div>
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

    $(document).ready(function () {
        $("#btnCancel").click(function () {
            window.location = 'listofretired.asp';
        }
    );
    });

    </script>
</body>
</html>

