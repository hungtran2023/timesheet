<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%
dim dblInstitution,dblField,dblDegree,strYearofGraduation,dblAtaslEducationID
dim dblAtlasCourseID,strParticipantNote,dblAtlasParticipantID
dim dblGroupOfSkill,dblSkill,intSelfAssessment,	intDeFacto,strDateDeFacto,dblStaffSkillID
Dim strUserid, rsSrc
'"Beginner (Basic training with <3 months project experience)","Intermediate(>3 months project experience)","Advanced(>12 months project experience)","Trainer"

'***************************************************************
'
'***************************************************************
function EducationProcess
	dim strSql
	
	Dim objDatabase, dblSubId
	Dim strConnect, strDigest,  strError	
	
	dblSubId=Request.QueryString("Subid")
	
	dblInstitution=request.form("lstInstitution")
	dblField=request.form("lstField")
	dblDegree=request.form("lstDegree")
	strYearofGraduation=request.form("txtYofGraduation")
	dblAtaslEducationID=request.form("txStaffEducationID")
	if dblAtaslEducationID="" then dblAtaslEducationID=-1
	
	if dblSubId="" then
	    if cint(dblAtaslEducationID)=-1 then
	        strSql="INSERT INTO ATC_StaffEducation (StaffID,InstitutionID,FieldID,DegreeID,YearOfGraduate) VALUES(" & _
	                strUserid & "," & dblInstitution & "," & dblField & "," & dblDegree & "," & strYearofGraduation & ")"
	    else
	        strSql="UPDATE ATC_StaffEducation SET " &_
	                " InstitutionID = " & dblInstitution & _
	                " ,FieldID = " & dblField & _
	                " ,DegreeID = " & dblDegree & _
	                " ,YearOfGraduate = " & strYearofGraduation & _
	                " WHERE StaffEducationID=" & dblAtaslEducationID
	    end if
	else
	    strSql="DELETE FROM ATC_StaffEducation WHERE StaffEducationID=" & dblSubId
	end if
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase     
    strError=""
	If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strSql) then 
			   strError = objDatabase.strMessage
            end if			  
        'response.write strSql
    end if
	EducationProcess=strError
end function

'***************************************************************
'
'***************************************************************
function PopulateListEducation
	
	Dim strSql,strReturn,i
	Dim objDatabase, rs
	Dim strConnect, strID,  strError	
	
	strSql="SELECT a.StaffEducationID, a.InstitutionID, a.FieldID, a.DegreeID, a.YearOfGraduate, b.InstitutionName, c.FieldName, d.DegreeName " & _
            "FROM        ATC_StaffEducation AS a INNER JOIN " & _
                         "ATC_Institutions AS b ON a.InstitutionID = b.InstitutionID INNER JOIN " & _
                         "ATC_FieldofStudy AS c ON a.FieldID = c.FieldID INNER JOIN " & _
                         "ATC_Degree AS d ON a.DegreeID = d.DegreeID " & _
            "WHERE StaffID=" & strUserid
     
    Call GetRecordset(strSql,rs)  
    strReturn=""
    i=1
    if not rs.Eof then
	  Do Until rs.EOF
	      strID=rs("StaffEducationID") & "#" & rs("InstitutionID") & "#" & rs("FieldID") & "#" & rs("DegreeID")
	      strReturn= strReturn & "<tr idValue='"  & strID & "'><td>" & i & ".</td>"
          strReturn= strReturn & "<td>" & rs("InstitutionName") & "</td>"
          strReturn= strReturn & "<td>" & rs("FieldName") & "</td> "
          strReturn= strReturn & "<td class='YOG' >" & rs("YearOfGraduate") & "</td> "
          strReturn= strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item eduitem' data-id='" & rs("StaffEducationID")  & "'></button></td>  </tr>"		
	    rs.MoveNext
	    i=i+1
	  Loop       
	end if         
	
	PopulateListEducation=strReturn
end function

'***************************************************************
'
'***************************************************************
function AtlasTrainingProcess
	dim strSql
	
	Dim objDatabase, dblSubId
	Dim strConnect, strDigest,  strError	
	
	dblSubId=Request.QueryString("Subid")
		
	dblAtlasCourseID=request.form("lstCourse")
	strParticipantNote=request.form("txtParticipantNote")
	dblAtlasParticipantID=request.form("txtAtlasParticipantID")
	if dblAtlasParticipantID="" then dblAtlasParticipantID=-1
	
	if dblSubId="" then
	    if cint(dblAtlasParticipantID)=-1 then
	        strSql="INSERT INTO ATC_AtlasCourseParticipant  (StaffID,AtlasCourseID,ParticipantNote) VALUES(" & _
	                strUserid & "," & dblAtlasCourseID & "," & IIF(strParticipantNote="", "null","'" & strParticipantNote & "'") & ")"
	    else
	        strSql="UPDATE ATC_AtlasCourseParticipant SET " &_
	                " AtlasCourseID = " & dblAtlasCourseID & _
	                " ,ParticipantNote = " & IIF(strParticipantNote="", "null","'" & strParticipantNote & "'") & _
	                " WHERE AtlasParticipantID=" & dblAtlasParticipantID
	    end if
	else
	    strSql="DELETE FROM ATC_AtlasCourseParticipant WHERE AtlasParticipantID=" & dblSubId
	end if
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase     
    strError=""
	If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strSql) then 
			   strError = objDatabase.strMessage
            end if			  
        'response.write strSql
    end if
	AtlasTrainingProcess=strError
end function

'***************************************************************
'
'***************************************************************
function PopulateListAtlasTrainingParticipant
	
	Dim strSql,strReturn,i
	Dim objDatabase, rs
	Dim strConnect, strID,  strError	
	
	strSql="SELECT        a.AtlasCourseCode, a.CourseName, b.ParticipantNote, b.AtlasParticipantID, b.StaffID, b.AtlasCourseID " & _
            " FROM        ATC_AtlasCourseParticipant AS b INNER JOIN " & _
                         " ATC_AtlasCourses AS a ON b.AtlasCourseID = a.AtlasCourseID " &_
            " WHERE b.StaffID=" & strUserid
   
    Call GetRecordset(strSql,rs)  
    strReturn=""
    i=1
    if not rs.Eof then
	  Do Until rs.EOF
	      strID=rs("AtlasParticipantID") & "#" & rs("AtlasCourseID") 
	      strReturn= strReturn & "<tr idValue='"  & strID & "'><td>" & i & ".</td>"
          strReturn= strReturn & "<td>" & rs("AtlasCourseCode") & "-" & rs("CourseName") & "</td>"
          strReturn= strReturn & "<td class='parNote' >" & rs("ParticipantNote") & "</td> "
          strReturn= strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item coursetem' data-id='" & rs("AtlasParticipantID")  & "'></button></td>  </tr>"		
	    rs.MoveNext
	    i=i+1
	  Loop       
	end if         
	
	PopulateListAtlasTrainingParticipant=strReturn
end function


'***************************************************************
'
'***************************************************************
function AtlasSkillProcess
	dim strSql
	
	Dim objDatabase, dblSubId
	Dim strConnect, strDigest,  strError	
	
	dblSubId=Request.QueryString("Subid")
	
    dblGroupOfSkill=request.form("lstGroupSkill")
	dblSkill=request.form("lstSkill")
	intSelfAssessment=request.form("lstSelfAssessment")
	intDeFacto=request.form("lstDeFacto")
	strDateDeFacto=request.form("txtDateDeFacto")
	if strDateDeFacto<>"" then strDateDeFacto=ConvertTommddyyyy(strDateDeFacto)
	dblStaffSkillID=request.form("txtStaffSkillID")
	if dblStaffSkillID="" then dblStaffSkillID=-1
	
	if dblSubId="" then
	    if cint(dblStaffSkillID)=-1 then
	    
	    strsql ="INSERT INTO ATC_StaffSkills(StaffID,SkillID,SelfAssessment,Defactor,DateDeFactor)VALUES (" & _
	            strUserid & "," & dblSkill & "," & intSelfAssessment & "," & IIF(intDeFacto="", "null",intDeFacto) & "," & IIF(strDateDeFacto="", "null","'" & strDateDeFacto & "'")  & ")"
	    else
		        strSql="UPDATE ATC_StaffSkills SET " & _
	                " SkillID = " & dblSkill & _
                    ",SelfAssessment =" & intSelfAssessment & _
                    ",Defactor =" &  IIF(intDeFacto="", "null",intDeFacto) & _
                    ",DateDeFactor =" & IIF(strDateDeFacto="", "null","'" & strDateDeFacto & "'") & _
                    " WHERE StaffSkillID=" & dblStaffSkillID
	    end if
	else
	    strSql="DELETE FROM ATC_StaffSkills WHERE StaffSkillID=" & dblSubId
	end if
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase     
    strError=""
	If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strSql) then 
			   strError = objDatabase.strMessage
            end if			  
        'response.write strSql
    end if
	AtlasSkillProcess=strError
end function

'***************************************************************
'
'***************************************************************
function PopulateListAtlasSkills
	
	Dim strSql,strReturn,strSkillFilter,strGroupSkillFilter
	Dim objDatabase, rs
	Dim strConnect, strID,  strError		
	dim arrLevel
	
	
	arrLevel=array("","Beginner","Intermediate","Advanced","Trainer")
	strSkillFilter=Request.querystring("ss")
	strGroupSkillFilter=Request.querystring("gs")
	
	strSql="SELECT        a.StaffSkillID, a.StaffID, a.SkillID, a.SelfAssessment, ISNULL(a.Defactor,0) as Defactor, DateDeFactor, b.SkillName, b.GroupOfSkill " & _
                "FROM    ATC_StaffSkills AS a INNER JOIN " & _
                         " ATC_Skills AS b ON a.SkillID = b.SkillID" &_
            " WHERE a.StaffID=" & strUserid

    if strGroupSkillFilter<>"" then strSql=strSql & " AND GroupOfSkill=" & strGroupSkillFilter
    if strSkillFilter<>"" then strSql=strSql & " AND a.SkillID=" & strSkillFilter
    
    strSql=strSql & " ORDER BY SkillName"
    Call GetRecordset(strSql,rs)  
    strReturn=""
    i=1
    
    if not rs.Eof then
	  Do Until rs.EOF
	      strID=rs("StaffSkillID") & "#" & rs("SkillID") & "#" & rs("SelfAssessment") & "#" & rs("Defactor") & "#" & rs("GroupOfSkill") 
	      strReturn= strReturn & "<tr idValue='"  & strID & "'>"
          strReturn= strReturn & "<td>" & rs("SkillName") & "</td>"
          strReturn= strReturn & "<td>" & arrLevel(rs("SelfAssessment")) & "</td>"
          strReturn= strReturn & "<td>" & arrLevel(rs("Defactor")) & "</td>"
          strReturn= strReturn & "<td class='DateDeFactor' >" & IIF(ISNULL(rs("DateDeFactor")),"",day(rs("DateDeFactor")) & "/" & month(rs("DateDeFactor")) & "/" & Year(rs("DateDeFactor"))) & "</td> "
          strReturn= strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item atlskillitem' data-id='" & rs("StaffSkillID")  & "'></button></td>  </tr>"		
	    rs.MoveNext
	    i=i+1
	  Loop       
	end if         
	
	PopulateListAtlasSkills=strReturn
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
'----------------------------------------
' analyse query string
'----------------------------------------
	gMessage = ""
    strUserid= Request.querystring("id")
    if strUserid="" then strUserid=Request.Form("txtuserid")    

	strAct = Request.QueryString("act")  
	
	Select Case strAct
      Case "edu"    
            gMessage= EducationProcess()        
      Case "atl"
            gMessage= AtlasTrainingProcess()        
      Case "ski"
            gMessage= AtlasSkillProcess()        
      Case else
        
    End Select
	
'strYofGraduation= Request.Form("txtYofGraduation")
'response.write "stract:" & stract & "YofGraduation:" & strYofGraduation

'--------------------------------------------------
' Initialize recordset
'--------------------------------------------------	
		strSql="SELECT * FROM  ATC_Institutions WHERE  (fgActivate = 1) ORDER BY InstitutionName"	
		Call GetRecordset(strSql,rs)

	    strInstitution= PopulateDataToListWithoutSelectTag(rs,"InstitutionID", "InstitutionName",-1)
	    
	    strSql="SELECT * FROM ATC_FieldofStudy WHERE fgActivate=1 ORDER BY FieldName"	
		Call GetRecordset(strSql,rs)

	    strField= PopulateDataToListWithoutSelectTag(rs,"FieldID", "FieldName",-1)
	    
	    strSql="SELECT * FROM ATC_Degree ORDER BY DegreeName"	
		Call GetRecordset(strSql,rs)
	    stDegree= PopulateDataToListWithoutSelectTag(rs,"DegreeID", "DegreeName",-1)
	    
	    strSql = "SELECT GroupOfSkillID,GroupOfSkillName FROM ATC_GroupOfSkills WHERE  fgActivate=1"
	    Call GetRecordset(strSql,rs)
	    stGroupSkill= PopulateDataToListWithoutSelectTag(rs,"GroupOfSkillID", "GroupOfSkillName",-1)
	    
	    strSql="SELECT SkillID ,SkillName ,GroupOfSkill FROM ATC_Skills WHERE fgActivate=1"	
		Call GetRecordset(strSql,rs)
	    strSkills =PopulateDataToListWithoutSelectTag(rs,"SkillID", "SkillName",-1)
	    
	    strSql="SELECT AtlasCourseID,AtlasCourseCode +' - ' + CourseName as AtlasCourseCodeName FROM ATC_AtlasCourses ORDER BY StartDate"	
		Call GetRecordset(strSql,rs)
	    strAtlasCourses =PopulateDataToListWithoutSelectTag(rs,"AtlasCourseID", "AtlasCourseCodeName",-1)
   
	    strEducation=PopulateListEducation()
	 	    
	    strAtlasEducation=PopulateListAtlasTrainingParticipant()	    
	    strAtlasSkills= PopulateListAtlasSkills()

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
<!-- Tab functions -->
    <div class="row">
        <div class="col-sm-12">
            <ul class="nav nav-tabs">
                <li><a href="employeeProfile.asp?id=<%=strUserid%>">Employee Profile</a></li>
                <li><a href="atlasinformation.asp?id=<%=strUserid%>">Atlas Information</a></li>
                <li class="active"><a>Education/Skill</a></li>
                <li><a href="replacementhistory.asp?id=<%=strUserid%>">Replacement History</a></li>
                <li><a href="employmenthistory.asp?id=<%=strUserid%>">Employment History</a></li>
            </ul>
        </div>
    </div>
	<div class="row">
	    <div class="col-sm-12">
	        <div class="tab-content employee-details-form">
	            <div class="row">
	                <div class="col-md-12 col-sm-6 col-xs-12">
<%if gMessage<>"" then%>   	                
                        <div class="alert alert-danger">
                            <strong>Error:</strong> <%=gMessage%>
                        </div>  
<%end if%>	                                              
                        <div class="panel panel-default">
                            <div class="panel-heading clearfix">
                                <h3 class="panel-title">Education</h3>
                            </div>
                             <form id='frmEducation' class="form-horizontal row-border" method="POST" action="educationskill.asp?id=<%=strUserid%>&act=edu" >
                                <div class="panel-body">
                                    <div class="col-sm-6">
                                        <div class="form-group">
                                            <label class="col-md-12">Institution</label>
                                            <div class="col-md-12">
                                                <select id="lstInstitution" name="lstInstitution" class="form-control sel-institution-list" >
                                                    <option value=""></option>
                                                    <%=strInstitution%>
                                                </select>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Field of Study</label>
                                            <div class="col-md-12">
                                            <select  id="lstField" name="lstField"  class="form-control sel-field-of-study">
                                                <option value=""></option>
                                                <%=strField%>
                                            </select>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-sm-6">
                                        <div class="form-group">
                                            <label class="col-md-12">Degree</label>
                                            <div class="col-md-12">
                                                <select id="lstDegree" name="lstDegree" class="form-control sel-degree">
                                                    <option value=""></option>
                                                    <%=stDegree%>
                                                </select>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Year of Graduation</label>
                                            <div class="col-md-12">
                                                <input type="number" id="txtYofGraduation" name="txtYofGraduation" class="form-control"  value="">
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-sm-12">
                                        <div class="form-group text-right">
                                            <div class="col-md-12">
                                                <button type="submit" class="btn btn-primary" id="btnEducationSave">Save</button>
                                                <button type="button" class="btn btn-primary hide" id="btnEducationCancel" >New</button>
                                            </div>
                                        </div>
                                    </div>
                                    <input type="hidden" name="txStaffEducationID" id="txStaffEducationID"/>
				            
                                    <div class="col-md-12">
                                        <table class="table table-striped table-bordered table-hover " id="tblListAtlasEducation">
                                            <thead class="thead-inverse">
                                                <tr>
                                                    <th>No.</th>
                                                    <th>Institution</th>
                                                    <th>Field of Study</th>
                                                    <th>Year of Graduation</th>
                                                    <th class="col-action"></th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <%=strEducation%>
                                            </tbody>
                                            <tbody class="education-listing"></tbody>
                                        </table>
                                    </div>
                                </div>
                            </form>
                            <div class="panel-heading clearfix">
                                <h3 class="panel-title">Atlas Traning</h3>
                            </div>
                            <form id='frmAtlasTraing' class="form-horizontal row-border" method="POST" action="educationskill.asp?id=<%=strUserid%>&act=atl" > 
                                <div class="panel-body">
                                    <div class="col-sm-12 no-padding">
                                        <div class="col-sm-6">
                                            <div class="form-group">
                                                <label class="col-md-12">Course</label>
                                                <div class="col-md-12">
                                                    <select id="lstCourse" name="lstCourse" class="form-control sel-course-name">
                                                        <option value=""></option>
                                                        <%=strAtlasCourses%>                                               
                                                    </select>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12">
                                            <div class="form-group">
                                                <label class="col-md-12">Note</label>
                                                <div class="col-md-12">
                                                    <textarea id="txtParticipantNote" name="txtParticipantNote"  cols="30" rows="3" class="form-control"></textarea>
                                                </div>
                                            </div>
                                        </div>	                                        
                                        <div class="col-sm-12">
                                            <div class="form-group text-right">
                                                <div class="col-md-12">
                                                    <button type="submit" class="btn btn-primary" id="btnAtlasTrainingSave">Save</button>
                                                    <button type="button" class="btn btn-primary hide" id="btnAtlasTrainingCancel">Cancel</button>
                                                    
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    <input type="hidden" name="txtAtlasParticipantID" id="txtAtlasParticipantID"/>
                                    <div class="col-md-12">
                                        <table class="table table-striped table-bordered table-hover" id="tblListAtlasParticipant">
                                            <thead class="thead-inverse">
                                                <tr>
                                                    <th class="col-md-1"> No.</th>
                                                    <th class="col-md-5">Course</th>
                                                    <th class="col-md-9">Note</th>
                                                    <th class="col-md-1 col-action"></th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <%=strAtlasEducation%>                                                                                              
                                             </tbody>   
                                        </table>
                                    </div>
	                            </div>                          
                            </form>

			                <div class="panel-heading clearfix">
			                    <h3 class="panel-title">Skills</h3>
			                </div>
			                <form id='frmSkill'  class="form-horizontal row-border" method="POST" action="educationskill.asp?id=<%=strUserid%>&act=ski" >
                				<div class="panel-body">
				                    <div class="col-sm-6">
				                        <div class="form-group">
				                            <label class="col-md-12">Group of Skills</label>
				                            <div class="col-md-12">
				                                <select id="lstGroupSkill" name="lstGroupSkill"  class="form-control">
				                                    <option value=""></option>
				                                    <%=stGroupSkill%>  
				                                </select>
				                            </div>
				                        </div>
				                    </div>
				                    <div class="col-sm-6">
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
                    				<div class="col-sm-12 no-padding">
				                        <div class="col-sm-6">
				                            <div class="form-group">
				                                <label class="col-md-12">Self Assessment</label>
				                                <div class="col-md-12">
				                                    <select id="lstSelfAssessment" name="lstSelfAssessment"  class="form-control">
				                                        <option value=""></option>
				                                       <option value="1">Beginner (Basic training with <3 months project experience)</option>
				                                       <option value="2">Intermediate(>3 months project experience)</option>
				                                       <option value="3">Advanced(>12 months project experience)</option>
				                                       <option value="4">Trainer</option>
				                                    </select>
				                                </div>
				                            </div>
				                        </div>
				                        <div class="col-sm-6">
				                            <div class="form-group">
				                                <div class="col-md-8 no-padding">
				                                    <label class="col-md-12">De Facto</label>
				                                    <div class="col-md-12">
				                                        <select id="lstDeFacto" name="lstDeFacto"  class="form-control">
				                                            <option value=""></option>
				                                            <option value="1">Beginner</option>
				                                            <option value="2">Intermediate</option>
				                                            <option value="3">Advanced</option>
				                                            <option value="4">Trainer</option>
				                                        </select>
				                                    </div>
				                                </div>
				                                <div class="col-md-4 padding-right15">
				                                    <label class="col-md-12">Date</label>
				                                    
				                                     <div class="input-group date">
				                                        <input type="text" name="txtDateDeFacto" id="txtDateDeFacto" class="form-control datepicker inp-apply-from-date" placeholder="DD/MM/YYYY" value="">
				                                        <span class="input-group-addon">
				                                            <span class="ic-calendar"></span>
				                                        </span>
				                                    </div>
				                                </div>
				                            </div>
				                        </div>
	                   				 </div>
                				</div>
                				 <input type="hidden" name="txtStaffSkillID" id="txtStaffSkillID"/>
                				<div class="col-sm-12">
                                    <div class="form-group text-right">
                                        <div class="col-md-12">
                                            <button type="submit" class="btn btn-primary" id="btnSkillSave">Save</button>
                                            <button type="button" class="btn btn-primary hide" id="btnSkillCancel">Cancel</button>                                            
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="col-sm-6">
				                        <div class="form-group">
				                            <div class="col-md-12">
				                                <label class="col-md-2">Fillter by Skill</label>
				                                <div class="col-md-3">
				                                    <select id="lstGroupSkillSearch" name="lstGroupSkillSearch" class="form-control">
				                                        <option value=""></option>
				                                         <%=stGroupSkill%>  
				                                    </select>
				                                </div>
				                                <div class="col-md-5">
				                                    <select id="lstSkillSearch" name="lstSkillSearch" class="form-control">
				                                        <option value=""></option>
				                                        <%=strSkills%>
				                                    </select>
				                                </div>
				                                <div class="col-md-2">
				                                    <button type="button" class="btn btn-primary" id="btnSkillSearch">Search</button> 
				                                </div>
				                            </div>
				                        </div>
				                    </div>
				                <div class="panel-body">
				                    <div class="col-md-12">
				                        <table class="table table-striped table-bordered table-hover"  id="tblListSatffSkill">
				                            <thead class="thead-inverse">
				                                <tr>
				                                    <th data-sortable="true" class="col-md-3">Skill</th>
				                                    <th data-sortable="true" class="col-md-3">Self Assessment</th>
				                                    <th data-sortable="true" class="col-md-3">De Facto</th>					                                    
				                                    <th data-sortable="true" class="col-md-3">Date</th>
				                                    <th class="col-action"></th>
				                                </tr>
				                            </thead>
				                            <tbody>
				                                <%=strAtlasSkills%>
				                            </tbody>
				                        </table>
				                    </div>
				                </div>
				            </form>
            			</div>    				
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
    
   $("#txStaffEducationID").val(-1);
   
   $('#frmEducation').formValidation({
            framework: 'bootstrap',
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },        
            fields: {
                lstInstitution:{
                    validators: {
                        notEmpty: {
                            message: 'The Institution is required.'
                        }
                    }
                },
                lstField:{
                    validators: {
                        notEmpty: {
                            message: 'The Field of study is required.'
                        }
                    }
                },
                lstDegree:{
                    validators: {
                        notEmpty: {
                            message: 'The Field of study is required.'
                        }
                    }
                },
                
                txtYofGraduation: {
                    validators: {
                        notEmpty: {
                            message: 'The year of graduation is required and must be a number'
                        },
                        between: {
                                min: 1970,
                                max:<%=Year(Date())-2%>,
                                message: 'The year of graduation must be after 1970 and before <%=Year(Date())-2%>'
                        }
                    }
                }
            }
        });//end of validation
        
          
        $("#tblListAtlasEducation tbody tr").click(function(){
            var res = $(this).attr("idValue").split("#");
            $("#txStaffEducationID").val(res[0]);
            $("#lstInstitution").val(res[1]);
            $("#lstField").val(res[2]);
            $("#lstDegree").val(res[3]);
            $("#txtYofGraduation").val($(this).children("td.YOG").text());
            
            $("#btnEducationCancel").removeClass('hide');
            //$("#btnCancelPromotion").addClass('hide');
            });
        
        $("#btnEducationCancel").click(function(){
            $("#txStaffEducationID").val(-1);
            $("#lstInstitution").val("");
            $("#lstField").val("");
            $("#lstDegree").val("");
            $("#txtYofGraduation").val("");
            $("#btnEducationCancel").addClass('hide');
        });
       
        $('.eduitem').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to delete this education.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'educationskill.asp?id=<%=strUserid%>&act=edu&subid=' + id);
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        }); 
        
//**********************************************************
// Atlas Trainning
//**********************************************************
        $('#frmAtlasTraing').formValidation({
            framework: 'bootstrap',
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },        
            fields: {
                lstCourse:{
                    validators: {
                        notEmpty: {
                            message: 'The Course is required.'
                        }
                    }
                }
            }
        });//end of validation
        
       $("#tblListAtlasParticipant tbody tr").click(function(){
            var res = $(this).attr("idValue").split("#");

            $("#txtAtlasParticipantID").val(res[0]);
            $("#lstCourse").val(res[1]);
            $("#txtParticipantNote").val($(this).children("td.parNote").text());
            
            $("#btnAtlasTrainingCancel").removeClass('hide');
            //$("#btnCancelPromotion").addClass('hide');
            });
 
      $("#btnAtlasTrainingCancel").click(function(){
            $("#txtAtlasParticipantID").val(-1);
            $("#lstCourse").val("");
            $("#txtParticipantNote").val("");
            $(this).addClass('hide');
        });
        
        $('.coursetem').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to remove user out this course.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'educationskill.asp?id=<%=strUserid%>&act=atl&subid=' + id);
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        }); 

//**********************************************************
// Atlas Skill
//**********************************************************
        $('#txtDateDeFacto')
        .on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmSkill').formValidation('revalidateField', 'txtDateDeFacto');
        });
         $('#frmSkill').formValidation({
            framework: 'bootstrap',
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },        
            fields: {
                lstGroupSkill:{
                    validators: {
                        notEmpty: {
                            message: 'The Group Of Skill is required.'
                        }
                    }
                },                
                lstSkill:{
                    validators: {
                        notEmpty: {
                            message: 'The Skill is required.'
                        }
                    }
                },
                lstSelfAssessment:{
                    validators: {
                        notEmpty: {
                            message: 'The Self Assessment is required.'
                        }
                    }
                },
                txtDateDeFacto:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The Date of De Facto is not a valid'
                        }
                    }  
                 }               
            }
        });//end of validation
        
         $("#tblListSatffSkill tbody tr").click(function(){
            var res = $(this).attr("idValue").split("#");
            
            $("#txtStaffSkillID").val(res[0]);
            $("#lstGroupSkill").val(res[4]);
            $("#lstSkill").val(res[1]);
            $("#lstSelfAssessment").val(res[2]);
            $("#lstDeFacto").val(res[3]);
            $("#txtDateDeFacto").val($(this).children("td.DateDeFactor").text());
            
            $("#btnSkillCancel").removeClass('hide');
            
            });        
        
        $("#btnSkillCancel").click(function(){
            $("#txtStaffSkillID").val(-1);
            $("#lstGroupSkill").val("");
            $("#lstSkill").val("");
            $("#lstSelfAssessment").val("");
            $("#lstDeFacto").val("");
            $("#txtDateDeFacto").val("");
            $(this).addClass('hide');
        });
        
         $('.atlskillitem').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to remove this skill out of user.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'educationskill.asp?id=<%=strUserid%>&act=ski&subid=' + id);
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        });        
        
         $("#btnSkillSearch").click(function(){
            // Revalidate the date field
             window.location = 'educationskill.asp?id=<%=strUserid%>&gs=' + $("#lstGroupSkillSearch").val() +'&ss=' + $("#lstSkillSearch").val()  ; // redirect
        });
        
    });

</script>

</body>
</html>

