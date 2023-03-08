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
function AddProfile
	
	Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then
    
		        Set myCmd = Server.CreateObject("ADODB.Command")
		        Set myCmd.ActiveConnection = objDatabase.cnDatabase
		        myCmd.CommandType = adCmdStoredProc
		        myCmd.CommandText = "[InsertEmployeeProfile]"		
		
                Set myParam = myCmd.CreateParameter("FirstName", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("LastName", adVarChar,adParamInput,100)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("Gender", adBoolean,adParamInput)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("Birthday", adDate	,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("MarialStatus", adInteger,adParamInput)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("NationalilyID", adInteger,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("PrivateEmail", adVarChar,adParamInput,100)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("MobilePhone",adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("CardNumber",adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("Passport",adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("PassportExpire",adDate,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("CurrentAddress", adVarChar,adParamInput,200)
		        myCmd.Parameters.Append myParam			
 		        
		        Set myParam = myCmd.CreateParameter("EmergencyContact", adLongVarChar,adParamInput, 1000)
                myCmd.Parameters.Append myParam	
              
		        Set myParam = myCmd.CreateParameter("TaxCode", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("InsuranceBookNo", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("BankAccount1", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("BankAccount2", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("HomeAddress", adVarChar,adParamInput,100)
                myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("ContactVietnam", adVarChar,adParamInput,100)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("HomeCountry", adVarChar,adParamInput,100)              
		        myCmd.Parameters.Append myParam						
		        Set myParam = myCmd.CreateParameter("UserType", adUnsignedTinyInt,adParamInput)              
		        myCmd.Parameters.Append myParam
				
		        Set myParam = myCmd.CreateParameter("PersonID", adInteger,adParamOutput)
		        myCmd.Parameters.Append myParam	
	        

                myCmd("FirstName")      =   strFirstname
                myCmd("LastName")       =   strLastname
                myCmd("Gender")         =   strGender
                myCmd("Birthday")       =   ConvertTommddyyyy(strDateofBirth)
                
                myCmd("MarialStatus")   =   strMarialstatus
                myCmd("NationalilyID")  =   strNationality
                myCmd("PrivateEmail")   =   strPrivateEmail
                
                myCmd("MobilePhone")    =   strMobilephone
                myCmd("CardNumber")     =   strCardNumber
                myCmd("Passport")       =   strPassport
              
                if strExpireddate ="" then 
                    myCmd("PassportExpire") =    null
                else
                    myCmd("PassportExpire") =   strExpireddate
                end if              

                myCmd("CurrentAddress") =   strCurrentaddress                     
                myCmd("EmergencyContact")=  strEmergency              
                myCmd("TaxCode")        =   strTaxcode              
                myCmd("InsuranceBookNo")=   strInsuranceBookNo
                myCmd("BankAccount1")   =   strBankaccount1
                myCmd("BankAccount2")   =   strBankaccount2
                myCmd("Homeaddress")    =   strHomeaddress
                 
                myCmd("ContactVietnam") =   strContactVietnam
                myCmd("Homecountry")    =   strHomecountry
				myCmd("UserType")    =   1
				
		        myCmd.Execute

		        If Err.number > 0 Then
			        strError= Err.Description
		        Else
			        strError = ""
				    strUserid=myCmd("PersonID")
			    End If
		        Err.Clear
        	
		        set myCmd=nothing
       end if
	     
	set objDatabase=nothing	
	
	AddProfile=strError
	
end function



'***************************************************************
'
'***************************************************************
function UpdateProfile
	
	Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then
    
		        Set myCmd = Server.CreateObject("ADODB.Command")
		        Set myCmd.ActiveConnection = objDatabase.cnDatabase
		        myCmd.CommandType = adCmdStoredProc
		        myCmd.CommandText = "[UpdateEmployeeProfile]"		
		
                Set myParam = myCmd.CreateParameter("FirstName", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("LastName", adVarChar,adParamInput,100)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("Gender", adBoolean,adParamInput)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("Birthday", adDate	,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("MarialStatus", adInteger,adParamInput)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("NationalilyID", adInteger,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("PrivateEmail", adVarChar,adParamInput,100)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("MobilePhone",adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("CardNumber",adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("Passport",adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("PassportExpire",adDate,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("CurrentAddress", adVarChar,adParamInput,200)
		        myCmd.Parameters.Append myParam			
 		        
		        Set myParam = myCmd.CreateParameter("EmergencyContact", adLongVarChar,adParamInput, 1000)
                myCmd.Parameters.Append myParam	
              
		        Set myParam = myCmd.CreateParameter("TaxCode", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("InsuranceBookNo", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("BankAccount1", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("BankAccount2", adVarChar,adParamInput,50)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("HomeAddress", adVarChar,adParamInput,100)
                myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("ContactVietnam", adVarChar,adParamInput,100)
                myCmd.Parameters.Append myParam	
                Set myParam = myCmd.CreateParameter("HomeCountry", adVarChar,adParamInput,100)              
		        myCmd.Parameters.Append myParam						
		        
		        Set myParam = myCmd.CreateParameter("PersonID", adInteger,adParamInput)
		        myCmd.Parameters.Append myParam	
	        

                myCmd("FirstName")      =   strFirstname
                myCmd("LastName")       =   strLastname
                myCmd("Gender")         =   strGender
                myCmd("Birthday")       =   ConvertTommddyyyy(strDateofBirth) 'month(strDateofBirth) & "/" & day(strDateofBirth) & "/" & year(strDateofBirth)
                
                myCmd("MarialStatus")   =   strMarialstatus
                myCmd("NationalilyID")  =   strNationality
                myCmd("PrivateEmail")   =   strPrivateEmail
                
                myCmd("MobilePhone")    =   strMobilephone
                myCmd("CardNumber")     =   strCardNumber
                myCmd("Passport")       =   strPassport
              
                if strExpireddate ="" then 
                    myCmd("PassportExpire") =    null
                else
                    myCmd("PassportExpire") =  ConvertTommddyyyy(strExpireddate) ' month(strExpireddate) & "/" & day(strExpireddate) & "/" & year(strExpireddate) 
                end if              

                myCmd("CurrentAddress") =   strCurrentaddress                     
                myCmd("EmergencyContact")=  strEmergency              
                myCmd("TaxCode")        =   strTaxcode              
                myCmd("InsuranceBookNo")=   strInsuranceBookNo
                myCmd("BankAccount1")   =   strBankaccount1
                myCmd("BankAccount2")   =   strBankaccount2
                myCmd("Homeaddress")    =   strHomeaddress
                 
                myCmd("ContactVietnam") =   strContactVietnam
                myCmd("Homecountry")    =   strHomecountry
                myCmd("PersonID")       =   strUserid

		        myCmd.Execute

		        If Err.number > 0 Then
			        strError= Err.Description
		        Else
			        strError = ""
				    
			    End If
		        Err.Clear
        	
		        set myCmd=nothing
       end if
	     
	set objDatabase=nothing	
	
	UpdateProfile=strError
	
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


if strStatus="submit" then

    strFirstname=Request.Form("txtFirstName")
    strLastname=Request.Form("txtLastName")
    strGender=Request.Form("radGender")

    strDateofBirth=Request.Form("txtBirthDate")
    strCardNumber=Request.Form("txtCardNumber")
    strPassport=Request.Form("txtPassport")
    strExpireddate=Request.Form("txtExpireddate")
    strNationality=Request.Form("lstNationality")
    strTaxcode=Request.Form("txtTaxcode")
    strInsuranceBookNo=Request.Form("txtInsurance")
    strPrivateEmail=Request.Form("txtPrivateEmail")
    strMarialstatus=Request.Form("lstMarialstatus")
    strBankaccount1=Request.Form("txtBankAcc1")
    strBankaccount2=Request.Form("txtBankAcc2")
    strCurrentaddress=Request.Form("txtCurrentAddr")
    strHomeaddress=Request.Form("txtHomeAddr")
    strMobilephone=Request.Form("txtMobile")
    strContactVietnam=Request.Form("txtContactVN")
    strHomecountry=Request.Form("txtHomeCountry")
    strEmergency=Request.Form("txtEmergency") 
    
    
    'Response.write  strFirstname & "-" &strLastname & "-" &strGender & "-" &strDateofBirth & "-" &strCardNumber & "-" &strPassport & "-" &strExpireddate & "<br>"
    'Response.write  strNationality & "-" &strTaxcode & "-" &strInsuranceBookNo & "-" &strPrivateEmail & "-" &strMarialstatus & "-" &strBankaccount1 & "-" &strBankaccount2 & "<br>"
    'Response.write  strCurrentaddress & "-" &strHomeaddress & "-" &strMobilephone & "-" &strContactVietnam & "-" &strHomecountry & "-" &strEmergency 
    
    if cdbl(strUserid)=-1 then
       gMessage= AddProfile()
    else
       gMessage= UpdateProfile()
    end if
    
    if gMessage="" then Response.Redirect("atlasinformation.asp?id=" & strUserid )
    
end if

	
    strSql="SELECT PersonID ,FirstName ,LastName ,Gender ,Birthday ,MarialStatus ,NationalityID ,PrivateEmail ,MobilePhone ,CardNumber ,Passport , " & _
                "PassportExpire ,CurrentAddress ," & _
                "EmergencyContact ,TaxCode ,InsuranceBookNo ,BankAccount1 ,BankAccount2 ,HomeAddress ,ContactVietnam ,HomeCountry ,fgDelete " & _
                "FROM ATC_PersonalInfo WHERE PersonID=" & strUserid
           
    Call GetRecordset(strSql,rsSrc)
      	
    if rsSrc.RecordCount>0 then
    
    
        strFirstname=rsSrc("FirstName")
        strLastname=rsSrc("LastName")
        strGender=iif(RsSrc("Gender"),1,0)
        
        strDateofBirth=day(rsSrc("Birthday")) & "/" & month(rsSrc("Birthday")) & "/" & year(rsSrc("Birthday"))
                
        strCardNumber=rsSrc("CardNumber")
        strPassport=rsSrc("Passport")
        strExpireddate=rsSrc("PassportExpire")
        if strExpireddate<>"" then
            strExpireddate=day(rsSrc("PassportExpire")) & "/" & month(rsSrc("PassportExpire")) & "/" & year(rsSrc("PassportExpire"))
        end if
        strNationality=rsSrc("NationalityID")
        strTaxcode=rsSrc("Taxcode")
        strInsuranceBookNo=rsSrc("InsuranceBookNo")
        strPrivateEmail=rsSrc("PrivateEmail")
        strMarialstatus=rsSrc("Marialstatus")
        strBankaccount1=rsSrc("BankAccount1")
        strBankaccount2=rsSrc("BankAccount2")
        strCurrentaddress=rsSrc("CurrentAddress")
        strHomeaddress=rsSrc("HomeAddress")
        strMobilephone=rsSrc("MobilePhone")
        strContactVietnam=rsSrc("ContactVietnam")
        strHomecountry=rsSrc("HomeCountry")
        strEmergency=rsSrc("EmergencyContact")
    else
        
        strFirstname=""
        strLastname=""
        strGender=-1
        strDateofBirth=""               
        strCardNumber=""
        strPassport=""
        strExpireddate=""
        strNationality=-1
        strTaxcode=""
        strInsuranceBookNo=""
        strPrivateEmail=""
        strMarialstatus=1
        strBankaccount1=""
        strBankaccount2=""
        strCurrentaddress=""
        strHomeaddress=""
        strMobilephone=""
        strContactVietnam=""
        strHomecountry=""
        strEmergency=""
    end if            
'--------------------------------------------------
' Photo
'--------------------------------------------------	    
	strPhoto="male"
	if strGender=0 then strPhoto="female"
	
	strSql = "SELECT  ISNULL(b.Photo,b.Username) as Username FROM ATC_PersonalInfo a LEFT JOIN ATC_Users b ON a.PersonID = b.UserID " &_
				"WHERE PersonID=" & strUserid
				
	Call GetRecordset(strSql,rsSrc)
	if rsSrc.RecordCount>0 then strPhoto=rsSrc("Username")
'--------------------------------------------------
' Initialize recordset
'--------------------------------------------------	
		strSql="SELECT CountryID,Nationality FROM ATC_Countries WHERE fgActivate=1 ORDER BY Nationality"	
		Call GetRecordset(strSql,rsNationality)
	    strNationality= PopulateDataToListWithoutSelectTag(rsNationality,"CountryID", "Nationality",cint(strNationality))

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
                    <li class="active"><a>Employee Profile</a></li>  
<%if cdbl(strUserid)<>-1 then%>
                    <li><a href="atlasinformation.asp?id=<%=strUserid%>">Atlas Information</a></li>
                    <li><a href="educationskill.asp?id=<%=strUserid%>">Education/Skill</a></li>
                    <li><a href="replacementhistory.asp?id=<%=strUserid%>">Replacement History</a></li>
                    <li><a href="employmenthistory.asp?id=<%=strUserid%>">Employment History</a></li>
<%end if%>                    
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
												<img id="imageProfile" name="imageProfile" height="130"  hspace="0" src="http://vnhcmcode/Timesheet/aisnet/Data/photos/<%=strPhoto%>.jpg?dummy=371662" width="100" onclick="onUploadImage();"
												onerror="this.src='http://ais.atlasindustries.com/staff/images/<%if cint(strGender)=0 then response.Write "Fe" end if%>male.jpg';">
										<!--	<button type="button"  style="margin-top:5px;" name="btnimage" id = "btnimage">Refresh</button>-->
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
										<div class="row">
											<div class="form-group">
												<label class="col-md-12">Date of Birth</label>
												<div class="col-md-12">
													<div class="input-group date">
															<input type="text" id="txtBirthDate" name="txtBirthDate" class="form-control datepicker" placeholder="DD/MM/YYYY"  value="<%=strDateofBirth%>" >
															<span class="input-group-addon">
																<span class="ic-calendar"></span>
															</span>
														</div>
												</div>
											</div>
										   
											<div class="form-group">
												<label class="col-md-12">Card Number</label>
												<div class="col-md-12">
													<input type="text" id="txtCardNumber" name="txtCardNumber" class="form-control"  value="<%=strCardNumber%>">
												</div>
											</div>
											<div class="form-group">
												<div class="col-md-6 no-padding">
													<label class="col-md-12">Passport</label>
													<div class="col-md-12">
														<input  type="text" id="txtPassport" name="txtPassport" class="form-control" value="<%=strPassport%>">
													</div>
												</div>
												<div class="col-md-6 no-padding">
													<label class="col-md-12">Expired date</label>
													<div class="col-md-12">
														<div class="input-group date">
															<input type="text" id="txtExpireddate" name="txtExpireddate" class="form-control datepicker" placeholder="DD/MM/YYYY" value="<%=strExpireddate%>">
															<span class="input-group-addon">
																<span class="ic-calendar"></span>
															</span>
														</div>
													</div>
												</div>
											</div>
											<div class="form-group">
												<label class="col-md-12">Nationality</label>
												<div class="col-md-12">
													<select id="lstNationality" name="lstNationality" class="form-control">
														<%=strNationality%>
													</select>
												</div>
											</div>
										</div>
									</div>
                                
									<div class="col-sm-6">
										<div class="form-group">
											<label class="col-md-12">Tax code</label>
											<div class="col-md-12">
												<input type="text" id="txtTaxcode" name="txtTaxcode" class="form-control" value="<%=strTaxcode%>">
											</div>
										</div>
									 
										<div class="form-group">
											<label class="col-md-12">Insurance Book No.</label>
											<div class="col-md-12">
												<input type="text" id="txtInsurance" name="txtInsurance" class="form-control" value="<%=strInsuranceBookNo%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Private Email</label>
											<div class="col-md-12">
												<input type="text" id="txtPrivateEmail" name="txtPrivateEmail" class="form-control" value="<%=strPrivateEmail%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Marital status</label>
											<div class="col-md-12">
												<select id="lstMarialstatus" name="lstMarialstatus" class="form-control">
													<option value="0" <%if cint(strMarialstatus) =0 then %>selected<%end if%>>Single</option>
													<option value="1" <%if cint(strMarialstatus) =1 then %>selected<%end if%>>Married</option>                                    
													<option value="2" <%if cint(strMarialstatus) =2 then %>selected<%end if%>>Separated</option>
													<option value="3" <%if cint(strMarialstatus) =3 then %>selected<%end if%>>Divorced</option>                                    
													<option value="4" <%if cint(strMarialstatus) =4 then %>selected<%end if%>>Widowed</option>
												</select>
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Bank account 1</label>
											<div class="col-md-12">
												<input type="text" id="txtBankAcc1" name="txtBankAcc1" class="form-control" value="<%=strBankaccount1%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Bank account 2</label>
											<div class="col-md-12">
												<input type="text" id="txtBankAcc2" name="txtBankAcc2" class="form-control" value="<%=strBankaccount2%>">
											</div>
										</div>
									</div>
								</div>
								<div class="panel-heading clearfix">
									<i class="icon-calendar"></i>
									<h3 class="panel-title">Contact Information</h3>
								</div>
								<div class="panel-body">
									<div class="col-sm-6">
										<div class="form-group">
											<label class="col-md-12">Current address</label>
											<div class="col-md-12">
											   <input type="text" id="txtCurrentAddr" name="txtCurrentAddr" class="form-control" value="<%=strCurrentaddress%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Home address</label>
											<div class="col-md-12">
												<input type="text" id="txtHomeAddr" name="txtHomeAddr" class="form-control" value="<%=strHomeaddress%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Mobile phone</label>
											<div class="col-md-12">
												<input type="text" id="txtMobile" name="txtMobile" class="form-control" value="<%=strMobilephone%>">
											</div>
										</div>
									</div>
															  
									<div class="col-sm-6">
										<div class="form-group">
											<label class="col-md-12">Contact Vietnam</label>
											<div class="col-md-12">
												<input type="text" id="txtContactVN" name="txtContactVN" class="form-control" value="<%=strContactVietnam%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Home country</label>
											<div class="col-md-12">
												<input type="text" id="txtHomeCountry" name="txtHomeCountry" class="form-control" value="<%=strHomecountry%>">
											</div>
										</div>
										<div class="form-group">
											<label class="col-md-12">Emergency</label>
											<div class="col-md-12">
												<input type="text" id="txtEmergency" name="txtEmergency" class="form-control" value="<%=strEmergency%>">
											</div>
										</div>
									</div>
								</div>
							</div>
            
							<div class="col-sm-12">
								<div class="form-group text-right">
									
									<button type="submit" id="btnNext" class="btn btn-primary btnNext">Save & Next</button>
									<button type="button" id="btnCancel" class="btn btn-default">Cancel</button>
								</div>
							</div>
							<input type="hidden" name="txtuserid" id="txtuserid" value="<%=strUserid%>"/>
							<input type="hidden" name="txtstatus" value="<%=strStatus%>"/>
							<input type="hidden" name="txtPhoto"  id="txtPhoto" value="<%=strPhoto%>"/>
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
	<input type="file" id="fileupload" name="fileupload" style="display:none" />
	
<script type="text/javascript">


	function onUploadImage() {

		var userId = $("#txtuserid").val();
        window.location.href = 'http://vnhcmcode/Timesheet/aisnet/MessageContent/UploadFile?UserId='+userId;

	}
	

	$(document).ready(function () {

      //  $("#imageProfile").attr('src', $(this).src + '?' + (new Date()).getTime());

		//$("#btnimage").click(function () {
		//	debugger;
  //          var photoname = $("#txtPhoto").val();
		//	var urlimage = "http://vnhcmcode/Timesheet/aisnet/Data/photos/" + photoname + ".jpg";
		//	$("#imageProfile").attr("src", urlimage);
  //          var fullphoto = photoname + ".jpg";
		//	document.getElementById('imageProfile').src = document.getElementById('imageProfile').src + '?' + (new Date()).getTime();
		//	$("#imageProfile").attr("src", urlimage);

         
		//	//location.reload(false);
  //      }); 

    $('#txtBirthDate')
        .on('changeDate', function(e) {
            // Revalidate the date field
            $('#contactForm').formValidation('revalidateField', 'txtBirthDate');
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
            txtFirstName: {
                validators: {
                    notEmpty: {
                        message: 'The First Name is required and cannot be empty'
                    },
                    stringLength: {
                        max: 50,
                        message: 'The First Name must be less than 50 characters long'
                    }
                }
            },
            txtLastName: {
                validators: {
                    notEmpty: {
                        message: 'The Last Name is required and cannot be empty'
                    },
                    stringLength: {
                        max: 100,
                        message: 'The First Name must be less than 100 characters long'
                    }
                }
            },
            radGender: {
                err: '#radGenderMessage',
                validators: {
                    notEmpty: {
                        message: 'The gender is required.'
                    }
                }
            },
            txtBirthDate: {
                validators: {
                    notEmpty: {
                        message: 'The Birth Date is required'
                    },
                    date: {
                        format: 'DD/MM/YYYY',
                        message: 'The Birth Date is not a valid'
                    }
                }
            },
            txtExpireddate: {
                validators: {
                    date: {
                        format: 'DD/MM/YYYY',
                        message: 'The Expired Date is not a valid'
                    }
                }
            },
            
            txtPrivateEmail: {
                    validators: {
                        emailAddress: {
                            message: 'The value is not a valid email address'
                        },
                        regexp: {
                            regexp: '^[^@\\s]+@([^@\\s]+\\.)+[^@\\s]+$',
                            message: 'The value is not a valid email address'
                        }
                    }
                },
             lstMarialstatus:{
                validators: {
                    notEmpty: {
                        message: 'The Marial status is required'
                    }      
                    }                   
             },       
             lstNationality:{
                validators: {
                    notEmpty: {
                        message: 'The Nationality is required'
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
    
   $("#btnCancel").click( function()
        {
            window.location = 'listofemployee.asp';
        }
   );
  
});

function getTab(varid){
		
	document.contactForm.action = "atlasinformation.asp";
	document.contactForm.target = "_self";
	document.contactForm.submit();
}

</script>

</body>
</html>

