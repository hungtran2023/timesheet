<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%

Dim strUserid, rsSrc
Dim strAtlasCourseID
Dim strAtlasCourseCode,strCourseName,strInstructor,strLessonTime,strCourseTotal,strRoomName,strSchedule
Dim strSkills,dblSkillID,strStartDate,strLastDate,strCourseAPK,strRequestor,strNumberofLessons,strNote
dim blnNew

function CheckUnique (strAtlasCourseID)
    dim blnCheck, rsSrc
    
    blnError=""
    
    strSql="SELECT * FROM ATC_AtlasCourses WHERE AtlasCourseID<>"& strAtlasCourseID & " AND AtlasCourseCode='" & strAtlasCourseCode & "'"
'response.write strsql	
    Call GetRecordset(strSql,rsSrc)
    if not rsSrc.EOF then
        blnError= blnError & "<br>" & "The Course Code must be unique."
    end if
       
    CheckUnique=blnError
end function
'--------------------------------------------------
'
'--------------------------------------------------
function CheckDeleteCourse (strAtlasCourseID)
    dim blnCheck, rsSrc
    
    blnError=""
    
    strSql="SELECT count(*) as numberof FROM ATC_AtlasCourseParticipant WHERE AtlasCourseID="& strAtlasCourseID 
	
    Call GetRecordset(strSql,rsSrc)
	
    if rsSrc("numberof")>0 then
        blnError= blnError & "<br>" & "Please remove all assinments before delete this course."
    end if
       
    CheckDeleteCourse=blnError
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

strAct=Request.Querystring("act")

strAtlasCourseID=request.form("txtAtlasCourseID")
if strAtlasCourseID="" then strAtlasCourseID=-1

if strAct<>"" then
	
	if strAct="u" then
		strAtlasCourseCode=request.form("txtAtlasCourseCode")
		strCourseName=request.form("txtCourseName")
		strInstructor=request.form("txtInstructor")
		strLessonTime=request.form("txtLessonTime")
		strCourseTotal=request.form("txtCourseTotal")
		strRoomName=request.form("txtRoomName")
		strSchedule=request.form("txtSchedule")
		dblSkillID=request.form("lstSkill")
		strStartDate=request.form("txtStartDate")
		strLastDate=request.form("txtLastDate")
		strCourseAPK=request.form("txtCourseAPK")
		strRequestor=request.form("txtRequestor")
		strNumberofLessons=request.form("txtNumberofLessons")
		strNote=request.form("txtNote")
		
		strAtlasCourseID=request.form("txtAtlasCourseID")
		if strAtlasCourseID="" then strAtlasCourseID=-1
		gMessage= CheckUnique(strAtlasCourseID) 
		
		if cint(strAtlasCourseID)=-1 then
		
			strsql=	"INSERT INTO [ATC_AtlasCourses] ([AtlasCourseCode],[CourseName],[CourseAPK],[Instructor],[Requestor],[SkillID],[StartDate],[EndDate],[LessonTime],[CourseTotal],[NumberofLessons],[RoomName],[Schedule],[CourseNote]) " & _
						"VALUES (" & _
							"'" & strAtlasCourseCode & "'" & _
							",'" & strCourseName & "'" & _
							"," & iif (strCourseAPK="", "NULL","'" & strCourseAPK & "'") & _
							",'" &  strInstructor & "'" & _
							"," & iif (strRequestor="", "NULL","'" & strRequestor & "'") & _
							"," & dblSkillID & _
							",'" & ConvertTommddyyyy(strStartDate) & "'" & _
							",'" & ConvertTommddyyyy(strLastDate) & "'" & _
							"," & strLessonTime & _
							"," & strCourseTotal & _
							"," & iif (strNumberofLessons="", "NULL", strNumberofLessons)  & _
							",'" & strRoomName & "'" & _						
							",'" & strSchedule & "'" & _
							"," & iif (strNote="","NULL", "'" &  strNote & "'") & ")" 
		else
			strsql="UPDATE  ATC_AtlasCourses " & _
						"SET  AtlasCourseCode  = '" & strAtlasCourseCode & "'" & _
							", CourseName  = '" & strCourseName & "'" & _
							", CourseAPK  = " & iif (strCourseAPK="", "NULL","'" & strCourseAPK & "'") & _
							", Instructor  = '" &  strInstructor & "'" & _
							", Requestor  = " & iif (strRequestor="", "NULL","'" & strRequestor & "'") & _
							", SkillID  = " & dblSkillID & _
							", StartDate  = '" & ConvertTommddyyyy(strStartDate) & "'" & _
							", EndDate  = '" & ConvertTommddyyyy(strLastDate) & "'" & _
							", LessonTime  = " & strLessonTime & _
							", NumberofLessons  = " & iif (strNumberofLessons="", "NULL", strNumberofLessons)  & _
							", CourseTotal  = " & strCourseTotal & _
							", RoomName  = '" & strRoomName & "'" & _	
							", Schedule  = '" & strSchedule & "'" & _
							", CourseNote  = " & iif (strNote="", "NULL", "'" &  strNote & "'") & _ 
						" WHERE AtlasCourseID=" & strAtlasCourseID
	 
		end if
	else
		strAtlasCourseID=request.querystring("subid")
		gMessage=CheckDeleteCourse(strAtlasCourseID)
		if gMessage="" then
			strsql="DELETE FROM ATC_AtlasCourses WHERE AtlasCourseID=" & strAtlasCourseID
		end if
	end if
		
	if gMessage="" then
		strCnn = Application("g_strConnect")	
		Set objDatabase = New clsDatabase     
		
		If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strSql) then 
				gMessage = objDatabase.strMessage
			end if			  
		end if
		
		if gMessage="" then Response.Redirect("AtlasCourseList.asp")
		
	end if
end if

strsql="SELECT [AtlasCourseID],[AtlasCourseCode],[CourseName],[CourseAPK],[Instructor],[Requestor],[SkillID],[StartDate],[EndDate],[LessonTime],[NumberofLessons]" & _
			",[CourseTotal],[RoomName],[Schedule],[CourseNote] FROM [ATC_AtlasCourses] " & _
		"WHERE AtlasCourseID=" & strAtlasCourseID

Call GetRecordset(strSql,rsSrc)
        	
if rsSrc.RecordCount>0 then
	strAtlasCourseCode=rsSrc("AtlasCourseCode")
	strCourseName=rsSrc("CourseName")
	strInstructor=rsSrc("Instructor")
	strLessonTime=rsSrc("LessonTime")
	strCourseTotal=rsSrc("CourseTotal")
	strRoomName=rsSrc("RoomName")
	strSchedule=rsSrc("Schedule")
	dblSkillID=rsSrc("SkillID")
	strStartDate=day(rsSrc("StartDate")) & "/" & month(rsSrc("StartDate")) & "/" & year(rsSrc("StartDate"))
	if rsSrc("EndDate")<>"" then 
		strLastDate=day(rsSrc("EndDate")) & "/" & month(rsSrc("EndDate")) & "/" & year(rsSrc("EndDate"))
	else
		strLastDate=""
	end if
	strCourseAPK=rsSrc("CourseAPK")
	strRequestor=rsSrc("Requestor")
	strNumberofLessons=rsSrc("NumberofLessons")
	strNote=rsSrc("CourseNote")
else
	strAtlasCourseCode=""
	strCourseName=""
	strInstructor=""
	strLessonTime=""
	strCourseTotal=""
	strRoomName=""
	strSchedule=""
	dblSkillID=-1
	strStartDate=""
	strLastDate=""
	strCourseAPK=""
	strRequestor=""
	strNumberofLessons=""
	strNote=""
end if

'--------------------------------------------------
' Initialize recordset
'--------------------------------------------------	

		strSql="SELECT SkillID ,SkillName ,GroupOfSkill FROM ATC_Skills WHERE fgActivate=1"	
		Call GetRecordset(strSql,rs)
		strSkills =PopulateDataToListWithoutSelectTag(rs,"SkillID", "SkillName",cdbl(dblSkillID))


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
	if strChoseMenu = "" then strChoseMenu = "AH"
	
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
        <div class="navi-info"> 
                <a class="blue" href="AtlasCourseList.asp" onMouseOver="self.status='Show the list of employees'; return true;" onMouseOut="self.status=''">Atlas Training:</a>
            <span>Course Details</span>
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
                        <form class="form-horizontal row-border" id="contactForm" method="POST" action="AtlasCourseDetail.asp?act=u">
<%if gMessage<>"" then%>                        
                            <div id="messages" class="alert alert-danger">
                                <strong>Error:</strong> <%=gMessage%>.
                            </div>
<%end if%>
    						<div class="panel panel-default">
				                <div class="panel-body">
				                    <div class="col-sm-6">
				                        <div class="form-group has-error">
				                            <label class="col-md-12" >AtlasCourseCode</label>
				                            <div class="col-md-12">
				                                <input type="text" id="txtAtlasCourseCode" name="txtAtlasCourseCode" class="form-control" value="<%=strAtlasCourseCode%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Course Title:</label>
				                            <div class="col-md-12">
				                                <input type="text" id="txtCourseName" name="txtCourseName" class="form-control"  value="<%=strCourseName%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Instructor</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtInstructor" name="txtInstructor" class="form-control"  value="<%=strInstructor%>">
				                            </div>
				                        </div>
				                        
				                        <div class="form-group">
				                            <label class="col-md-12">Lesson Time (hrs)</label>
				                            <div class="col-md-12">
				                                <input type="number"  id="txtLessonTime" name="txtLessonTime" class="form-control"  value="<%=strLessonTime%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Course Total (hrs)</label>
				                            <div class="col-md-12">
				                                <input type="number"  id="txtCourseTotal" name="txtCourseTotal" class="form-control"  value="<%=strCourseTotal%>">
				                            </div>
				                        </div>
				                        
				                       <div class="form-group">
				                            <label class="col-md-12">Location</label>
				                            <div class="col-md-12">				                              
				                                    <input type="text"  id="txtRoomName" name="txtRoomName" class="form-control"  value="<%=strRoomName%>">				                              
				                            </div>
				                        </div>
				                         
				                        <div class="form-group">
				                            <label class="col-md-12">Schedule</label>
				                            <div class="col-md-12">				                              
				                                    <input type="text"  id="txtSchedule" name="txtSchedule" class="form-control"  value="<%=strSchedule%>">				                              
				                            </div>
				                        </div>
				                        <div class="form-group">
			                                <label class="col-md-12">Skill</label>
			                                <div class="col-md-12">
			                                    <select id="lstSkill" name="lstSkill" class="form-control">
			                                        <option value=""></option>
			                                        <%=strSkills%>
			                                    </select>
			                                </div>
				                        </div>
				                    </div>
				                    <div class="col-sm-6">
				                        <div class="form-group">
				                            <label class="col-md-12">Start Date</label>
				                            <div class="col-md-12">
				                                <div class="input-group date">
				                                    <input type="text"  id="txtStartDate" name="txtStartDate" class="form-control datepicker" placeholder="DD/MM/YYYY"  value="<%=strStartDate%>">
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
				                                    <input type="text"  id="txtLastDate" name="txtLastDate" class="form-control datepicker" placeholder="DD/MM/YYYY"  value="<%=strLastDate%>">
				                                    <span class="input-group-addon">
				                                        <span class="ic-calendar"></span>
				                                    </span>
				                                </div>
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Course APK</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtCourseAPK" name="txtCourseAPK" class="form-control"  value="<%=strCourseAPK%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Requestor</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtRequestor" name="txtRequestor" class="form-control"  value="<%=strRequestor%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Number of Lessons</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtNumberofLessons" name="txtNumberofLessons" class="form-control"  value="<%=strNumberofLessons%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
                                            <label class="col-md-12">Note</label>
                                            <div class="col-md-12">
                                                <textarea  id="txtNote" name="txtNote"  cols="30" rows="4" class="form-control" ><%=strNote%></textarea>
                                            </div>
                                        </div>
				                    </div>
            				        
                                    <input type="hidden" name="txtAtlasCourseID" id="txtAtlasCourseID" value="<%=strAtlasCourseID%>"/>
				                    
				                </div>
				                	                                      
<%if strPromotionList<>"" then%>				                    
                                    <div class="col-md-12">
                                        <table class="table table-striped table-bordered table-hover " id="tblListJobtitle">
                                            <thead class="thead-inverse">
                                                <tr>
                                                    <th class="col-md-1 col-md-checkbox">
                                                        <input type="checkbox"></th>
                                                    <th>Job Title</th>
                                                    <th>From</th>
                                                    <th>To</th>
                                                    <th class="col-action"></th>
                                                </tr>
                                            </thead>
                                             <%=strPromotionList%> 
                                        </table>
                                    </div>				                    
				                    
				                </div>
<%end if%>				                
				            </div>

				            
				            <div class="col-sm-12">
				                <div class="form-group text-right">
				                    <button type="submit" id="btnNext" class="btn btn-primary btnNext">Save</button>
				                    <button type="button" id="btnDelete" class="btn btn-primary btnDelete">Delete</button>
				                    <button type="button" id="btnCancel" class="btn btn-default">Cancel</button>				                    
				                </div>
				            </div>                            
				            
				            <input type="hidden" name="txtID" value="<%=strID%>"/>
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
                    <p>You are about to delete this course.</p>
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
            
    $('#txtStartDate')
        .on('changeDate', function(e) {
            // Revalidate the date field
            $('#contactForm').formValidation('revalidateField', 'txtStartDate');
        });
    $('#txtLastDate')
    .on('changeDate', function(e) {
        // Revalidate the date field
        $('#contactForm').formValidation('revalidateField', 'txtLastDate');
    });
        
   $('#contactForm').formValidation({
        framework: 'bootstrap',
        icon: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },        
        fields: {
            txtAtlasCourseCode: {
                validators: {
                    notEmpty: {
                        message: 'The Course Code is required and cannot be empty'
                    },
                    stringLength: {
                        max: 16,
                        min:16,
                        message: 'The Staff Code must be 16 characters long'
                    }
                }
            },
            txtCourseName : {
                validators: {
                    notEmpty: {
                        message: 'The Course Name is required and cannot be empty'
                    },
                    stringLength: {
                        max: 200,
                        message: 'The Course Name must be less than 200 characters long'
                    }
                }
            },
             
            txtInstructor : {
                validators: {
                    notEmpty: {
                        message: 'The Instructor is required and cannot be empty'
                    }                   
                },
                
            },
            txtLessonTime : {
                validators: {
                    notEmpty: {
                        message: 'The Lesson Time is required and must be a number'
                    }                   
                }
            },
            txtCourseTotal : {
                validators: {
                    notEmpty: {
                        message: 'The Course Total hours is required and must be a number'
                    }                   
                }
            },
            txtRoomName : {
                validators: {
                    notEmpty: {
                        message: 'The Location is required and cannot be empty'
                    }                   
                }
            },
            txtSchedule : {
                validators: {
                    notEmpty: {
                        message: 'The Schedule is required and cannot be empty'
                    }                   
                }
            },    
            lstSkill : {
                validators: {
                    notEmpty: {
                        message: 'The Skill is required and cannot be empty'
                    }                   
                }
            },
            
            
            txtStartDate: {
                    validators: {
                        notEmpty: {
                            message: 'The start date is required'
                        },
                        date: {
                            format: 'DD/MM/YYYY',
                            max: 'txtLastDate',
                            min:'1/1/1900',
                            message: 'The start date is not a valid or the Start Date must be before the End Date.'
                        }
                    }
                },
                txtLastDate: {
                    validators: {
                        notEmpty: {
                            message: 'The end date is required'
                        },
                        date: {
                            format: 'DD/MM/YYYY',
                            min: 'txtStartDate',
                            max: '<%=Date()+365*10%>',
                            message: 'The end date is not a valid or the End Date must be after the Start Date.'
                        }
                    }
                }            
           
            }
         })
         .on('success.field.fv', function(e, data) {
            if (data.field === 'txtStartDate' && !data.fv.isValidField('txtLastDate')) {
                // We need to revalidate the end date
                data.fv.revalidateField('txtLastDate');
            }

            if (data.field === 'txtLastDate' && !data.fv.isValidField('txtStartDate')) {
                // We need to revalidate the start date
                data.fv.revalidateField('txtStartDate');
            }
        });
        
   
        $("#btnCancel").click(function(){
       
            window.location = 'AtlasCourseList.asp';
        });
		
		$("#btnDelete").on('click', function(e) {
            e.preventDefault();
 
            var id = $("#txtAtlasCourseID").val();
 
            $("#modal_message").html("You are about to remove this course.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'AtlasCourseDetail.asp?act=d&subid='+id );
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        });      
    
   
 
 });
 
 
 

</script>

</body>
</html>

