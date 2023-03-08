<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%
dim strStaffIDHR,strEmailID,strLoginID,strPassword,strOldPassword,strStartDate,strLastDate,strExTel,intIndirect
Dim dblDepartmentID,dblWorkingHoursID,dblReportToID,dblCSOLevelID, dblCompanyID
Dim strDepartment,strWorkingHours,strReportTo,strCSOLevel, strJobTitle , strCompany
Dim dblJobTitleID,strApplyFrom,dblPromotionID, dblJobRoleID
Dim strUserid, rsSrc
dim blnNew
dim strOldStartDate, dblOldWorkingHoursID


'***************************************************************
'
'***************************************************************
function PromotionList(rsSrc)
    dim strReturn
    
   if rsSrc.Recordcount>0 then
        strReturn="<tbody>"
        Do while not rsSrc.EOF
            
            dblJobTitleID=rsSrc("JobtitleID") 
            strApplyFrom=day(rsSrc("ApplyFrom")) & "/" & month(rsSrc("ApplyFrom")) & "/" & Year(rsSrc("ApplyFrom"))
           
            dblPromotionID=rsSrc("PromotionID")
            
            strReturn=strReturn & "<tr idValue='"  & rsSrc("PromotionID") & "'><td class='col-md-1 col-md-checkbox'><input type='checkbox' value='" & rsSrc("PromotionID") &  "'></td>"
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
        loop
         strReturn=strReturn & "</tbody>"
   end if
   
   PromotionList=strReturn
                                             
end function
'***************************************************************
'
'***************************************************************
function CheckNewAtlasMember(strUserid)
    dim blnCheck, rsSrc
    
    blnCheck=true
    
    strSql="SELECT * FROM ATC_Employees WHERE StaffID=" & strUserid
    Call GetRecordset(strSql,rsSrc)
       
    CheckNewAtlasMember=(rsSrc.EOF OR(rsSrc.RecordCount=0))
end function

'***************************************************************
'
'***************************************************************
function CheckUnique (strUserid)
    dim blnCheck, rsSrc
    
    blnError=""
    
    strSql="SELECT * FROM ATC_Employees WHERE staffID<>"& strUserid & " AND StaffIDHR='" & strStaffIDHR & "'"
	
    Call GetRecordset(strSql,rsSrc)
    if not rsSrc.EOF then
        blnError= blnError & "<br>" & "The Staff ID must be unique"
    end if
    
    strSql="SELECT * FROM ATC_Employees WHERE staffID<>"& strUserid & " AND EmailID='" & strEmailID & "'"
    Call GetRecordset(strSql,rsSrc)
    
    if not rsSrc.EOF then
        blnError= blnError & "<br>" & "The Email ID must be unique."
    end if
    
    strSql="SELECT * FROM ATC_Users WHERE UserID <> "& strUserid & " AND UserName='" & strLoginID & "'"
    Call GetRecordset(strSql,rsSrc)
    
    if not rsSrc.EOF then
        blnError= blnError & "<br>" & "The login ID has aready been used."
    end if        
       
    CheckUnique=blnError
end function

'***************************************************************
'
'***************************************************************
function AddAtlasEmployee
	
	Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then
        
        if strPassword<>"" or blnNew then
            Set objSHA1 = New clsSHA1	
            strDigest = ObjSHA1.SecureHash(strPassword)
            Set ObjSHA1 = Nothing
        end if
        
        objDatabase.cnDatabase.BeginTrans
        
            
            if strLastDate="" then
                strLastDate="NULL"
            else
                strLastDate="'" & month(strLastDate) & "/" & day(strLastDate) & "/" & year(strLastDate) & "'"
            end if
            
            strSql= "INSERT INTO ATC_Employees " & _
			               "([StaffID],[DepartmentID],[DirectLeaderID],[JoinDate],[LeaveDate],[fgIndirect] " & _
			               ",[ExtPhone] ,[EmailID] ,[CSOLevel], JobRoleID ,StaffIDHR,[CompanyID] )" & _
		             "VALUES(" & strUserid & "," & dblDepartmentID & "," & iif(cdbl(dblReportToID)=-1,"NULL", dblReportToID )  & ",'" & ConvertTommddyyyy(strStartDate) & "'," & _
		                    strLastDate & "," & intIndirect & "," & iif(strExTel="","NULL","'" & strExTel & "'") &_
		                     ",'" & strEmailID & "'," & dblCSOLevelID & "," & dblJobRoleID & ",'" & strStaffIDHR &  "'," & dblCompanyID & ")"

			if not objDatabase.runActionQuery(strSql) then 
			    strError = objDatabase.strMessage
			else
            
                strSql="INSERT INTO [ATC_Users] " & _
			                "([UserID],[UserName] ,[Password]) " & _
		                "VALUES (" & strUserid & ",'" & strLoginID & "','" & strDigest & "')"

                if not objDatabase.runActionQuery(strSql) then 
			        strError = objDatabase.strMessage
			    else
                    strSql="INSERT INTO [ATC_SalaryStatus] " & _
                                "([StaffID],[SalaryDate],[WorkingHourID]) " & _
                            "VALUES (" & strUserid & ",'" & ConvertTommddyyyy(strStartDate)& "'," & dblWorkingHoursID  & ")"
										
                    if not objDatabase.runActionQuery(strSql) then 
			            strError = objDatabase.strMessage
						response.write strSql
'response.end
			        else
			            strSql="INSERT INTO ATC_Promotion " & _
			                    "([StaffID],[JobtitleID],[ApplyFrom]) " & _ 
			                "VALUES (" & strUserid & "," & dblJobTitleID & ",'" & ConvertTommddyyyy(strStartDate) & "')" 
				
			            if not objDatabase.runActionQuery(strSql) then 
			                strError = objDatabase.strMessage
			            end if			            
			        end if
			    end if
			    
            end if			  

      if strError<>"" then 
	    objDatabase.cnDatabase.RollbackTrans
	  else
	  	objDatabase.cnDatabase.CommitTrans
	  end if
    end if
	AddAtlasEmployee=strError
	
end function

'***************************************************************
'
'***************************************************************
function UpdateAtlasEmployee
	
	Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then
        
        if strPassword<>"" or blnNew then
            Set objSHA1 = New clsSHA1	
            strDigest = ObjSHA1.SecureHash(strPassword)
            Set ObjSHA1 = Nothing
        else
            strDigest=strOldPassword
        end if
        
        objDatabase.cnDatabase.BeginTrans
        
            if strLastDate="" then
                strLastDate="NULL"
            else
                strLastDate="'" & ConvertTommddyyyy(strLastDate) & "'"
            end if
            
			
            strSql= "UPDATE ATC_Employees " & _
		            "SET " & _
                        "CompanyID = " & dblCompanyID & _
			            ",DepartmentID = " & dblDepartmentID & _
			            ",DirectLeaderID = " & iif(cdbl(dblReportToID)=-1,"NULL", dblReportToID )  & _
			            ",LeaveDate = " & strLastDate & _
			            ",fgIndirect = " & intIndirect & _
			            ",ExtPhone = " & iif(strExTel="","NULL","'" & strExTel & "'") & _
			            ",EmailID = '" & strEmailID & "'" & _
			            ",CSOLevel = " & dblCSOLevelID & _
                        ",JobRoleID = " & dblJobRoleID & _
			            ",StaffIDHR = '" & strStaffIDHR & "'" & _
		            " WHERE StaffID=" & strUserid
			if not objDatabase.runActionQuery(strSql) then 
			    strError = objDatabase.strMessage
			else
                strSql="UPDATE [ATC_Users] " & _
                		   "SET " & _
			                   "[UserName] = '" & strLoginID & "'" & _
			                   ",[Password]= '" & strDigest & "'" & _
		                    " WHERE [UserID]=" & strUserid

                if not objDatabase.runActionQuery(strSql) then 
			       strError = objDatabase.strMessage
			    else
					if cdbl(dblPromotionID)=-1 then
							strSql="INSERT INTO ATC_Promotion " & _
									"([StaffID],[JobtitleID],[ApplyFrom]) " & _ 
								"VALUES (" & strUserid & "," & dblJobTitleID & ",'" & ConvertTommddyyyy(strApplyFrom) & "')" 
					else
						strSql="UPDATE ATC_Promotion SET JobtitleID=" & dblJobTitleID & ",ApplyFrom='" &  ConvertTommddyyyy(strApplyFrom) & "' WHERE PromotionID=" & dblPromotionID
					end if
					if not objDatabase.runActionQuery(strSql) then 
						strError = objDatabase.strMessage
					else
						dblOldWorkingHoursID=request.form("txtOldWorkingHours")
						if dblWorkingHoursID<>dblOldWorkingHoursID then
							strSql="UPDATE [ATC_SalaryStatus] " & _
								"SET " & _
										"[WorkingHourID] = " & dblWorkingHoursID & _
									" WHERE [StaffID]= " & strUserid & " AND [SalaryDate] IN (SELECT MAX([SalaryDate]) FROM [ATC_SalaryStatus] WHERE  [StaffID]="  & strUserid & ")"
						end if

						if strSql<>"" then
							if not objDatabase.runActionQuery(strSql) then 
								strError = objDatabase.strMessage
							else
								strOldStartDate=request.form("txtOldStartDate")
								if ConvertTommddyyyy(strStartDate)<>ConvertTommddyyyy(strOldStartDate) then
									strSql="UPDATE [ATC_SalaryStatus] " & _
											"SET " & _
											"[SalaryDate]= '" & ConvertTommddyyyy(strStartDate) & _
										"' WHERE [StaffID]= " & strUserid & " AND [SalaryDate] IN (SELECT MIN([SalaryDate]) FROM [ATC_SalaryStatus] WHERE  [StaffID]="  & strUserid & ")"
									
								end if
								
								if strSql<>"" then
									if not objDatabase.runActionQuery(strSql) then strError = objDatabase.strMessage
								end if
							end if
						end if
						'response.write strSql			            
			        end if
			    end if
            end if			  

                
		if strError<>"" then 
			objDatabase.cnDatabase.RollbackTrans
		else
			objDatabase.cnDatabase.CommitTrans
		end if
    end if


	UpdateAtlasEmployee=strError
	
end function
'--------------------------------------------------
' 
'--------------------------------------------------
function HideAnEmployee()
    Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then        
        objDatabase.cnDatabase.BeginTrans
        
            strSql= "UPDATE ATC_PersonalInfo   SET fgDelete = 1" & _
		            " WHERE PersonID=" & strUserid

			if not objDatabase.runActionQuery(strSql) then 
			    strError = objDatabase.strMessage			
            end if			  

                
      if strError<>"" then 
	    objDatabase.cnDatabase.RollbackTrans
	  else
	  	objDatabase.cnDatabase.CommitTrans
	  end if
    end if
    HideAnEmployee=strError
end function

'--------------------------------------------------
' 
'--------------------------------------------------
function RemovePromotion(intPromotionID)
    Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then        
        objDatabase.cnDatabase.BeginTrans
        
            strSql= "DELETE ATC_Promotion  WHERE PromotionID=" & intPromotionID

			if not objDatabase.runActionQuery(strSql) then 
			    strError = objDatabase.strMessage			
            end if			  

                
      if strError<>"" then 
	    objDatabase.cnDatabase.RollbackTrans
	  else
	  	objDatabase.cnDatabase.CommitTrans
	  end if
    end if
    RemovePromotion=strError
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
strUserid= Request.querystring("id")

if strUserid="" then strUserid=Request.Form("txtuserid")

strAct=Request.Querystring("act")

if strAct="d" then
    gMessage=HideAnEmployee()
    if gMessage="" then Response.Redirect("listofemployee.asp")
elseif strAct="dj" then
	
	gMessage=RemovePromotion(Request.querystring("subid"))
    'if gMessage="" then Response.Redirect("listofemployee.asp")
end if

blnNew=CheckNewAtlasMember(strUserid)
if strStatus="submit" then

        strStaffIDHR=Request.Form("txtStaffIDHR")
        strEmailID=Request.Form("txtEmailID")
        
        strLoginID=Request.Form("txtLoginID")
        strPassword=Request.Form("txtPassword")
        
        strStartDate=Request.Form("txtStartDate")
        strLastDate=Request.Form("txtLastDate")
        strExTel=Request.Form("txtExTel")
        intIndirect=Request.Form("radDirect")
      
        if intIndirect="" then 
            intIndirect=0
        else
            intIndirect=1
        end if

        dblCompanyID=Request.Form("lstCompany")
        dblDepartmentID=Request.Form("lstDepartment")
        dblWorkingHoursID=Request.Form("lstWorkingHours")
        dblReportToID=Request.Form("lstReportTo")
        dblCSOLevelID=Request.Form("lstCSOLevel")
        dblJobRoleID=Request.Form("lstJobRole")
        
        strOldPassword=Request.Form("txtOldPassword")
            
        dblJobTitleID=Request.Form("lstJobTitle")
        strApplyFrom=Request.Form("txtApplyFrom")
        dblPromotionID=Request.Form("txtPromotionID")
        if dblPromotionID="" then dblPromotionID=-1
        
    gMessage= CheckUnique(strUserid) 
   
    if gMessage ="" then
        if blnNew then
            gMessage=AddAtlasEmployee()
        else
            gMessage=UpdateAtlasEmployee()
        end if
        'if gMessage="" then Response.Redirect("listofemployee.asp")
    end if
       
end if


strSql="SELECT a.DepartmentID, a.CompanyID, a.DirectLeaderID, a.JoinDate, a.LeaveDate, a.fgIndirect, a.ExtPhone, a.EmailID, ISNULL(a.CSOLevel,-1) as CSOLevel,ISNULL(a.JobRoleID,0) as JobRoleID, a.StaffIDHR, " & _
        " b.WorkingHourID, b.SalaryDate, c.UserName, c.Password " & _
        " FROM ATC_Employees AS a INNER JOIN " & _
                "ATC_Users AS c ON a.StaffID = c.UserID INNER JOIN  " & _
                "(select * from atc_salarystatus where staffid =" & strUserid & " and SalaryDate in (select max(SalaryDate) from atc_SalaryStatus where staffid = " & strUserid & " )) AS b ON a.StaffID = b.StaffID " & _
                "WHERE a.StaffID =" & strUserid
'response.Write strsql
Call GetRecordset(strSql,rsSrc)
    	
if rsSrc.RecordCount>0 then
    strStaffIDHR=rsSrc("StaffIDHR")
    strEmailID=rsSrc("EmailID")
    
    strLoginID=rsSrc("UserName")
    strOldPassword=rsSrc("Password")
    
    strStartDate=day(rsSrc("JoinDate")) & "/" & month(rsSrc("JoinDate")) & "/" & Year(rsSrc("JoinDate"))
    'strStartDate=ConvertToddmmyyyy(rsSrc("JoinDate"))
    
    strLastDate=rsSrc("LeaveDate")
    if strLastDate<>"" then
        strLastDate=day(rsSrc("LeaveDate")) & "/" & month(rsSrc("LeaveDate")) & "/" & Year(rsSrc("LeaveDate"))
    end if
    
    strExTel=rsSrc("ExtPhone")
    intIndirect=iif(rsSrc("fgIndirect"),1,0)

    dblCompanyID=rsSrc("CompanyID")
    dblDepartmentID=rsSrc("DepartmentID")
    dblWorkingHoursID=rsSrc("WorkingHourID")
    dblReportToID=rsSrc("DirectLeaderID")
    dblCSOLevelID=rsSrc("CSOLevel")
    dblJobRoleID=cdbl(rsSrc("JobRoleID"))
    
    strSql="SELECT * FROM ATC_Promotion a INNER JOIN ATC_JobTitle b ON a.JobTitleID=b.JobTitleID WHERE StaffID=" & strUserid & " ORDER BY ApplyFrom"
    Call GetRecordset(strSql,rsSrc)

    strPromotionList=PromotionList(rsSrc)    
    
else
    strStaffIDHR=""
    strEmailID=""
    
    strLoginID=""
    strPassword=""
    
    strStartDate=""
    strLastDate=""
    strExTel=""
    intIndirect=0
    dblCompanyID=-1
    dblDepartmentID=-1
    dblWorkingHoursID=-1
    dblReportToID=-1
    dblCSOLevelID=-1
    dblJobRoleID=-1
    
    dblJobTitleID=-1
    strApplyFrom=""
end if            


'--------------------------------------------------
' Initialize recordset
'--------------------------------------------------	

		strSql="SELECT DepartmentID, Department, fgActivate FROM  ATC_Department WHERE  (fgActivate = 1) ORDER BY Department"	
		Call GetRecordset(strSql,rsDepart)
	    strDepartment= PopulateDataToListWithoutSelectTag(rsDepart,"DepartmentID", "Department",dblDepartmentID)
	    
    	strSql="SELECT CompanyID, CompanyName, fgActiv FROM  ATC_Companies WHERE  type=0 AND  (fgActiv = 1) ORDER BY CompanyName"	
		Call GetRecordset(strSql,rsCompany)
	   strCompany= PopulateDataToListWithoutSelectTag(rsCompany,"CompanyID", "CompanyName",dblCompanyID)


	    strSql="SELECT CostingGroupID,CostingGroupName FROM ATC_CSOGroupCosting WHERE fgActivate=1 ORDER BY OrderList"	
  
		Call GetRecordset(strSql,rsCSOLevel)
	    strCSOLevel= PopulateDataToListWithoutSelectTag(rsCSOLevel,"CostingGroupID", "CostingGroupName",dblCSOLevelID)
	    
	    strSql="SELECT WorkingHourID ,[Hours] FROM ATC_WorkingHours"	
		Call GetRecordset(strSql,rsWorkingHour)
	    strWorkingHour= PopulateDataToListWithoutSelectTag(rsWorkingHour,"WorkingHourID", "Hours",dblWorkingHoursID)
    
	    strSql = "SELECT DISTINCT a.UserID, e.Firstname + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') as Fullname " &_
					"FROM ATC_UserGroup a LEFT JOIN ATC_Group b ON a.GroupID = b.GroupID " &_
					"LEFT JOIN ATC_Permissions c ON b.GroupID = c.GroupID " &_
					"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID " &_
					"LEFT JOIN ATC_PersonalInfo e ON a.UserID = e.PersonID " &_
					"WHERE d.Description = 'Receive Report' AND e.fgDelete = 0 ORDER BY Fullname"
	    Call GetRecordset(strSql,rsReportTo)
	    stReportTo= PopulateDataToListWithoutSelectTag(rsReportTo,"UserID", "Fullname",dblReportToID)
	    
	    strSql="SELECT [JobTitleID],[JobTitle]  FROM ATC_JobTitle  WHERE  (fgActivate = 1) ORDER BY JobTitle"	
		Call GetRecordset(strSql,rsJobtitle)
	    strJobtitle= PopulateDataToListWithoutSelectTag(rsJobtitle,"JobTitleID", "JobTitle",dblJobTitleID)

        strSql="SELECT [JobRoleID],[JobRoleName]  FROM ATC_JobRoles  WHERE  (fgActivate = 1) ORDER BY JobRoleName"
        
		Call GetRecordset(strSql,rsJobRole)
	    strJobRole= PopulateDataToListWithoutSelectTag(rsJobRole,"JobRoleID", "JobRoleName",dblJobRoleID)
	    
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
                <ul class="nav nav-tabs">
                <li><a href="employeeProfile.asp?id=<%=strUserid%>"> Employee Profile</a></li>
                <li class="active"><a>Atlas Information</a></li>
                <li><a href="educationskill.asp?id=<%=strUserid%>">Education/Skill</a></li>
                <li><a href="replacementhistory.asp?id=<%=strUserid%>">Replacement History</a></li>
                <li><a href="employmenthistory.asp?id=<%=strUserid%>">Employment History</a></li>
            </ul>
            </ul>
        </div>
    </div>
    
    <div class="row">
        <div class="col-sm-12">
            <div class="tab-content employee-details-form">
                <div class="row">
                    <div class="col-md-12 col-sm-6 col-xs-12">
                        <form class="form-horizontal row-border" id="contactForm" method="POST" >
<%if gMessage<>"" then%>                        
                            <div id="messages" class="alert alert-danger">
                                <strong>Error:</strong> <%=gMessage%>.
                            </div>
<%end if%>
            
    						<div class="panel panel-default">
				                <div class="panel-body">
				                    <div class="col-sm-6">
				                        <div class="form-group has-error">
				                            <label class="col-md-12" >Staff ID (<%=strUserid%>)</label>
				                            <div class="col-md-12">
				                                <input type="text" id="txtStaffIDHR" name="txtStaffIDHR" class="form-control" value="<%=strStaffIDHR%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Email ID</label>
				                            <div class="col-md-12">
				                                <input type="text" id="txtEmailID" name="txtEmailID" class="form-control"  value="<%=strEmailID%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Login ID</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtLoginID" name="txtLoginID" class="form-control"  value="<%=strLoginID%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Password</label>
				                            <div class="col-md-12">
				                                <div class="col-md-10 no-padding">
<%if blnNew then%>				                                
				                                    <p class="text-muted">Default for new user:blank</p>
				                                    </div>
<%else%>
				                                    <div id="passhidden">
				                                        <p class="text-muted"><< hidden for security >></p>				                                    
				                                    </div>
				                                    <div id="passinput" class="hide">
				                                        <input type="Password"  id="txtPassword" name="txtPassword" class="form-control"  value="">	
				                                    </div>
				                                 </div>
				                                
				                                 <div class="col-md-2 padding-left5">
                                                    <button type="button" id="btnReset"  class="btn btn-primary btn-full-width">Reset</button>
                                                    <button type="button" id="btnCancelResetPass"  class="btn btn-primary btn-full-width hide">Cancel</button>
                                                </div>
<%end if%>				                            
				                                
				                                
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
				                            <label class="col-md-12">Last Date</label>
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
				                            <label class="col-md-12">Ex. Tel</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtExTel" name="txtExTel" class="form-control"  value="<%=strExTel%>">
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <div class="col-md-12" >
				                                    <div class="col-md-1 no-padding width-auto" >
				                                        <input type="checkbox" name="radDirect" id="radDirect" value="1" class="no-padding" <%if cint(intIndirect)=1 then%>checked<%end if%>>
				                                    </div>
				                                    <label class="col-md-3 padding-left5 no-blod" >Indirect</label>
				                                    
				                                </div>
				                                 <div id="radDirectMessage" class="col-md-12" style="margin-bottom:0px;"></div>
				                        </div>
				                    </div>
				                    <div class="space-row30"></div>
				                    <div class="col-md-6">
                                            <div class="form-group">
				                            <label class="col-md-12">Company</label>
				                            <div class="col-md-12">
				                                <select class="form-control" name="lstCompany" id="lstCompany">
				                                    <%=strCompany%>
				                                </select>
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Department</label>
				                            <div class="col-md-12">
				                                <select class="form-control" name="lstDepartment" id="lstDepartment">
				                                    <%=strDepartment%>
				                                </select>
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Working Hours</label>
				                            <div class="col-md-12">
				                                <select class="form-control" name="lstWorkingHours" id="lstWorkingHours">
				                                    <%=strWorkingHour%>
				                                </select>
				                            </div>
				                        </div>
				                    </div>
				                    <div class="col-md-6">
				                        <div class="form-group">
				                            <label class="col-md-12">Report To</label>
				                            <div class="col-md-12">
				                                <select class="form-control" name="lstReportTo" id="lstReportTo">
				                                    <option value="-1"></option>
				                                    <%=stReportTo%>
				                                </select>
				                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">CSO Level</label>
				                            <div class="col-md-12">
				                                <select class="form-control" name="lstCSOLevel" id="lstCSOLevel">
													<option value=""></option>
				                                    <%=strCSOLevel%>
				                                </select>
				                            </div>
				                        </div>
                                         <div class="form-group">
				                            <label class="col-md-12">Job Role</label>
				                            <div class="col-md-12">
				                                <select class="form-control" name="lstJobRole" id="lstJobRole">
                                                     <option value="-1"></option>
				                                    <%=strJobRole%>
				                                </select>
				                            </div>
				                        </div>
				                    </div>
				                </div>
				                <div class="panel-heading clearfix">
				                    <h3 class="panel-title">Promotion</h3>
				                </div>
				                <div class="panel-body">
				                    <div class="col-sm-6">
				                        <div class="form-group">
				                            <label class="col-md-12">Job Title</label>
				                            <div class="col-md-12">
				                                <select class="form-control sel-job-title-list" name="lstJobTitle" id="lstJobTitle">
				                                    <%=strJobtitle%>
				                                </select>
				                                <input type="hidden" name="txtPromotionID" id="txtPromotionID"value="<%=dblPromotionID%>">
				                            </div>
				                        </div>
				                    </div>
<%if not blnNew then %>				                    
				                    <div class="col-sm-6">
				                        <div class="form-group">
				                            <label class="col-md-12">Apply From</label>
				                            <div class="col-md-12">
				                                <div class="input-group date">
				                                    <input type="text" name="txtApplyFrom" id="txtApplyFrom" class="form-control datepicker inp-apply-from-date" placeholder="DD/MM/YYYY" value="<%=strApplyFrom%>">
				                                    <span class="input-group-addon">
				                                        <span class="ic-calendar"></span>
				                                    </span>
				                                </div>
				                            </div>				                           
				                        </div>
				                    </div> 
				                    <div class="col-sm-12">
				                        <div class="form-group text-right">
		                                    <button type="button" id="btnNewPromotion" class="btn  btn-default btnNext">New</button>
		                                    <button type="button" id="btnCancelPromotion" class="btn btn-default hide ">Cancel</button>
		                                </div> 
		                            </div>
				                        
				                     
<%end if%>				                                      
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
				                    <button type="button" id="btnBack" class="btn btn-default btnPrevious">Back</button>
				                    <button type="submit" id="btnNext" class="btn btn-primary btnNext">Save & Close</button>
				                    <button type="button" id="btnDelete" class="btn btn-primary btnDelete">Delete</button>
				                    <button type="button" id="btnCancel" class="btn btn-default">Cancel</button>
				                    
				                </div>
				            </div>                            
				            
				            <input type="hidden" name="txtuserid" value="<%=strUserid%>"/>
				            <input type="hidden" name="txtstatus" value="<%=strStatus%>"/>
				            <input type="hidden" name="txtOldPassword" value="<%=strOldPassword%>"/>
							<input type="hidden" name="txtOldWorkingHours" value="<%=dblWorkingHoursID%>"/>
				            <input type="hidden" name="txtOldStartDate" value="<%=strStartDate%>"/>
							
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
//alert ("test");


        FormValidation.Validator.securePassword = {
                validate: function(validator, $field, options) {
                    var value = $field.val();
                    if (value === '') {
                        return true;
                    }

                    // Check the password strength
                    if (value.length < 8) {
                        return {
                            valid: false,
                            message: 'The password must be more than 8 characters long'
                        };
                    }

                    // The password doesn't contain any digit
                    if (value.search(/[0-9]/) < 0) {
                        return {
                            valid: false,
                            message: 'The password must contain at least one digit'
                        }
                    }
                    // The password doesn't contain any digit
                    if (value.search(/[!@#$%^&*-]/) < 0) {
                        return {
                            valid: false,
                            message: 'The password must contain at least one char in !@#$%^&*-'
                        }
                    }                    

                    return true;
                }
            };

            
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
    $('#txtApplyFrom')
    .on('changeDate', function(e) {
        // Revalidate the date field
        $('#contactForm').formValidation('revalidateField', 'txtApplyFrom');
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
            txtStaffIDHR: {
                validators: {
                    notEmpty: {
                        message: 'The Staff code is required and cannot be empty'
                    },
                    stringLength: {
                        max: 10,
                        min:10,
                        message: 'The Staff Code must be 10 characters long'
                    }
                }
            },
            txtEmailID : {
                validators: {
                    notEmpty: {
                        message: 'The Atlas EmailID is required and cannot be empty'
                    },
                    stringLength: {
                        max: 50,
                        message: 'The Atlas EmailID must be less than 50 characters long'
                    }
                }
            },
             
            txtLoginID : {
                validators: {
                    notEmpty: {
                        message: 'The Atlas LoginID is required and cannot be empty'
                    },
                    stringLength: {
                        max: 50,
                        message: 'The Atlas LoginID must be less than 50 characters long'
                    }
                }
            },
            txtStartDate      : {
                validators: {
                    notEmpty: {
                        message: 'The Start Date is required'
                    },
                    date: {
                        format: 'DD/MM/YYYY',
                        message: 'The Start Date is not a valid'
                    }
                }
            },
            txtPassword      : {
                validators: {
                    notEmpty: {
                        message: 'The Password is required'
                    },
                    securePassword: {
                        message: 'The password is not valid'
                    }
                    
                }
            },
             txtLastDate:{
                validators: {
                     date: {
                        format: 'DD/MM/YYYY',
                        message: 'The Last Date is not a valid'
                    }
               }       
                    
            },
			lstCSOLevel:{
                validators: {
                    notEmpty: {
                        message: 'The CSO level is required'
                    }      
                    }                   
             },       
            txtApplyFrom:{
                validators: {
                    notEmpty: {
                        message: 'The Applyform is required'
                    },
                     date: {
                        format: 'DD/MM/YYYY',
<%if not blnNew then%>                        
                        min: '<%=strStartDate%>',
<%end if%>                        
                        message: 'The Date is not a valid or it must be after the start date.'
                    }
               }       
                    
            }
        }
    })
   .on('success.form.fv', function(e) {
            var $form        = $(e.target),     // Form instance
                // Get the clicked button
                $button      = $form.data('formValidation').getSubmitButton(),
                // You might need to update the "status" field before submitting the form
                $statusField = $form.find('[name="txtstatus"]');

            // To demonstrate which button is clicked,
            // I use Bootbox (http://bootboxjs.com/) to popup a simple message
            // You might don't need to use it in real application

            switch ($button.attr('id')) {
                case 'btnSavePromotion':
                    $statusField.val('submitjobtitle');
                    break;
                case 'btnNext':
                    $statusField.val('submit');
                    break;
            }
         
   }); //end of validation
    
   $("#btnCancel").click( function()
        {
            window.location = 'listofemployee.asp';
        }
   );
   
    $("#btnBack").click( function()
        {
            window.location = 'employeeProfile.asp?id=<%=strUserid%>';
        }
     
   );
   
   $("#btnDelete").click(function()
   {
         //e.preventDefault();
		 $("#modal_message").html("You are about to delete this employee.");
         $('#confirm-delete').modal('show');
		 $('#confirm-delete').find('.btn-ok').attr('href', 'atlasinformation.asp?act=d&id=<%=strUserid%>' );	
   }
    );    
       
   $("#btnReset").click( function()
        {
               $(this).addClass('hide');
               $("#btnCancelResetPass").removeClass('hide');
               $("#passhidden").hide();
               $("#passinput").removeClass('hide');
            
        }
    );
    $("#btnCancelResetPass").click( function()
        {
            $(this).addClass('hide');
            $("#btnReset").removeClass('hide');
            $("#passhidden").show();
            $("#passinput").addClass('hide');
            
        }
    );
    
    //For edit user
  $("#tblListJobtitle tbody tr").click(function(){
    
        $("#txtPromotionID").val($(this).attr("idValue"));
        $("#lstJobTitle").val($(this).children("td.jobtitle").attr("jobid"));
        $("#txtApplyFrom").val($(this).children("td.applyfrom").text());
        $("#btnNewPromotion").removeClass('hide');
        $("#btnCancelPromotion").addClass('hide');
    });
   
   $("#btnCancelPromotion").click( function()
    {
        
        $("#txtPromotionID").val("<%=dblPromotionID%>");
        $("#lstJobTitle").val("<%=dblJobTitleID%>");
        $("#txtApplyFrom").val("<%=strApplyFrom%>");
        $("#btnNewPromotion").removeClass('hide');
        $("#btnCancelPromotion").addClass('hide');
           
    }); 
    
    $("#btnNewPromotion").click( function()
        {
            $(this).addClass('hide');
            $("#txtPromotionID").val(-1);
            $("#lstJobTitle").val(6);
            $("#txtApplyFrom").val("<%=Day(Date()) & "/" & Month(Date()) & "/" & Year(Date()) %>");
            $("#btnCancelPromotion").removeClass('hide');
        }
    );       
	
	 $('.btn-remove-item').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to remove this jobtitle.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'atlasinformation.asp?act=dj&id=<%=strUserid%>&subid='+id );
           //alert ('atlasinformation.asp?act=dj&id=<%=strUserid%>&subid='+id);
            //$('#myModal').data('id', id).modal('show');
        });  
 
 });
 
 
 function getjobtitle(varid){
       
//	document.contactForm.txtstatus.value = "submitjobtitle";
	
	document.navi.action = "atlasinformation.asp?id=<%=strUserid%>";
	document.contactForm.target = "_self";
	document.contactForm.submit();
}

</script>

</body>
</html>

