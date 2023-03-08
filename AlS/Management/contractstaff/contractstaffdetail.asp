<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%
'-------------------------------------------
	Dim varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strUserid
	Dim strfirst, strmiddle, strsurname
	Dim strmobile, stridnum, strusername
	Dim strpass, dblLevel, intdepartment, dblOT,intCompanyID
	Dim strExemail, strNote
	Dim gMessage, fgChanged

'***************************************************************
'
'***************************************************************
function AddTP( intcompanyID, strTitle ,strLastName,strMiddleName,strFirstName,strex_email,strmobilePhone,strusername,strpassword,intDepartID,dblLevelID,dblOvertimeID,strNote)
	
	Dim objDatabase, objSHA1
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then
	    
	    strQuery="SELECT * FROM ATC_Users WHERE UserName='" & strusername & "'"

	    ret = objDatabase.runQuery(strQuery)
    
		if ret then
		    if objDatabase.noRecord then 

                Set objSHA1 = New clsSHA1	
                strDigest = ObjSHA1.SecureHash(strPassword)
                Set ObjSHA1 = Nothing
        		
		        Set myCmd = Server.CreateObject("ADODB.Command")
		        Set myCmd.ActiveConnection = objDatabase.cnDatabase
		        myCmd.CommandType = adCmdStoredProc
		        myCmd.CommandText = "[InsertAContractStaff]"		

		        Set myParam = myCmd.CreateParameter("companyID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("Title", adVarChar,adParamInput,5)
		        myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("LastName", adVarChar,adParamInput,30)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("MiddleName", adVarChar,adParamInput,20)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("FirstName", adVarChar,adParamInput,20)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("ex_email", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("mobilePhone", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("username", adVarChar,adParamInput,20)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("password", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("DepartID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("LevelID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("OvertimeID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("Note", adLongVarChar,adParamInput,len(strNote)+1)
		        myCmd.Parameters.Append myParam						
		        Set myParam = myCmd.CreateParameter("PersonID", adInteger,adParamOutput)
		        myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("intErrorCode", adInteger,adParamOutput)
		        myCmd.Parameters.Append myParam


		        myCmd("companyID")	= intcompanyID		
		        myCmd("Title")		= strTitle
		        myCmd("LastName")   = strLastName
		        myCmd("MiddleName")	= strMiddleName
		        myCmd("FirstName")	= strFirstName
		        myCmd("ex_email")	= strex_email
		        myCmd("mobilePhone")= strmobilePhone
		        myCmd("username")	= strusername
		        myCmd("password")	= strDigest
		        myCmd("DepartID")   = intDepartID
		        myCmd("LevelID")	= dblLevelID
		        myCmd("OvertimeID")	= dblOvertimeID		
		        myCmd("Note")		= strNote

		        myCmd.Execute

		        If Err.number > 0 Then
			        strError= Err.Description
		        Else
			        if myCmd("intErrorCode")>0 then
				        strError="Failed to create user. Please contact IT supporter for helping."
			        else
				        strError = ""
				        strUserid=myCmd("PersonID")
			        end if
		        End If
		        Err.Clear
        	
		        set myCmd=nothing
            else
                strError="The Username is already existed."
            end if
        end if
    else
        strError=objDatabase.strMessage
    end if
	     
	set objDatabase=nothing	
	
	AddTP=strError
	
end function

'***************************************************************
'
'***************************************************************
function UpdateTP( intcompanyID, strTitle ,strLastName,strMiddleName,strFirstName,strex_email,strmobilePhone,strusername,strpass,intDepartID,dblLevelID,dblOvertimeID,strNote, PersonID)
	
	Dim objDatabase
	Dim strConnect, strDigest,  strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then
	    
	    strQuery="SELECT * FROM ATC_Users WHERE UserID<>" & PersonID & " AND UserName='" & strusername & "'"

	    ret = objDatabase.runQuery(strQuery)
    
		if ret then
		    if objDatabase.noRecord then 
                
                if strpass<>"" then

                    Set objSHA1 = New clsSHA1	
                    strDigest = ObjSHA1.SecureHash(strpass)
                    Set ObjSHA1 = Nothing
                    
                    strQuery="UPDATE ATC_Users SET Password='" & strDigest & "' WHERE UserID=" & PersonID
                    objDatabase.runQuery(strQuery)
                end if

		        Set myCmd = Server.CreateObject("ADODB.Command")
		        Set myCmd.ActiveConnection = objDatabase.cnDatabase
		        myCmd.CommandType = adCmdStoredProc
		        myCmd.CommandText = "[UpdateAContractStaff]"		

		        Set myParam = myCmd.CreateParameter("companyID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("Title", adVarChar,adParamInput,5)
		        myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("LastName", adVarChar,adParamInput,30)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("MiddleName", adVarChar,adParamInput,20)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("FirstName", adVarChar,adParamInput,20)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("ex_email", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("mobilePhone", adVarChar,adParamInput,50)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("username", adVarChar,adParamInput,20)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("DepartID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam
		        Set myParam = myCmd.CreateParameter("LevelID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam		
		        Set myParam = myCmd.CreateParameter("OvertimeID",adInteger,adParamInput)
		        myCmd.Parameters.Append myParam			
		        Set myParam = myCmd.CreateParameter("Note", adLongVarChar,adParamInput,len(strNote)+1)
		        myCmd.Parameters.Append myParam						
		        Set myParam = myCmd.CreateParameter("PersonID", adInteger,adParamInput)
		        myCmd.Parameters.Append myParam	
		        Set myParam = myCmd.CreateParameter("intErrorCode", adInteger,adParamOutput)
		        myCmd.Parameters.Append myParam


		        myCmd("companyID")	= intcompanyID		
		        myCmd("Title")		= strTitle
		        myCmd("LastName")   = strLastName
		        myCmd("MiddleName")	= strMiddleName
		        myCmd("FirstName")	= strFirstName
		        myCmd("ex_email")	= strex_email
		        myCmd("mobilePhone")= strmobilePhone
		        myCmd("username")	= strusername
		        myCmd("DepartID")   = intDepartID
		        myCmd("LevelID")	= dblLevelID
		        myCmd("OvertimeID")	= dblOvertimeID		
		        myCmd("Note")		= strNote
		        myCmd("PersonID")	= PersonID

		        myCmd.Execute

		        If Err.number > 0 Then
			        strError= Err.Description
		        Else
			        if myCmd("intErrorCode")>0 then
				        strError="Failed to update user information. Please contact IT supporter for helping."
			        else
				        strError = "Updated successfully"
			        end if
		        End If
		        Err.Clear
        	
		        set myCmd=nothing
            else
                strError="The Username is already existed."
            end if
        end if
    else
        strError=objDatabase.strMessage
    end if
	     
	set objDatabase=nothing	
	
	UpdateTP=strError
	
end function


'***************************************************************
'
'***************************************************************
function DeleteTP(PersonID)
	

	Dim strConnect,   strError	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
    
    strError=""
	If objDatabase.dbConnect(strCnn) Then

        		
        Set myCmd = Server.CreateObject("ADODB.Command")
        Set myCmd.ActiveConnection = objDatabase.cnDatabase
        myCmd.CommandType = adCmdStoredProc
        myCmd.CommandText = "[DeleteAContractStaff]"		

       
        Set myParam = myCmd.CreateParameter("PersonID", adInteger,adParamInput)
        myCmd.Parameters.Append myParam	
        Set myParam = myCmd.CreateParameter("count", adInteger,adParamOutput)
        myCmd.Parameters.Append myParam	
        Set myParam = myCmd.CreateParameter("intErrorCode", adInteger,adParamOutput)
        myCmd.Parameters.Append myParam

        myCmd("PersonID")	= PersonID

        myCmd.Execute
        
        If Err.number > 0 Then
	        strError= Err.Description
        Else
            if myCmd("count")>0 then
                strError="Some hours had booked by this user, cannot remove."
	        elseif myCmd("intErrorCode")>0 then
		        strError="Failed to remove user. Please contact IT supporter for helping."
	        else
		        strError = ""
	        end if
        End If
        Err.Clear
	
        set myCmd=nothing
            
    end if
	     
	set objDatabase=nothing	
	
	DeleteTP=strError
	
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
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if	

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
	if strChoseMenu = "" then strChoseMenu = "AG"
	
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

	strUserid=Request.Form("txtUserid")	
    if strUserid="" then strUserid=-1 '--Add new
    		
	strAct = Request.QueryString("act")

	if strAct<>"" then
	
	    Select case strAct
	        case "SAVE"
                struserID = Request.Form("txtuserid")
	            strfirst = trim(Request.Form("txtfirst"))
	            strmiddle = trim(Request.Form("txtmiddle"))
	            strsurname = trim(Request.Form("txtsurname"))
	            strtitleU = Request.Form("lsttitle")
            	
	            strExemail=trim(Request.Form("txtExemail"))
            	
	            strmobile = trim(Request.Form("txtmobile"))
	            stridnum = trim(Request.Form("txtidnum"))
	            strusernameU = trim(Request.Form("txtusername"))
	            strpass =  Request.Form("txtpass")
	            dblLevel = trim(Request.Form("lstLevel"))
	            intdepartment = trim(Request.Form("lstdepartment"))
	            dblOT = trim(Request.Form("lstOT"))
	            intCompanyID= trim(Request.Form("lstCompany"))
            	
                strNote = trim(Request.Form("txtNote"))
	            	        
	            if cint(struserID)=-1 then
		            gMessage = AddTP(intCompanyID,strtitleU, strsurname,strmiddle,strfirst,strExemail,strmobile,strusernameU, strpass, intdepartment, dblLevel, dblOT,strNote)
		             if gMessage="" THEN Response.Redirect("listofcontractstaff.asp")
	            else
		            gMessage = UpdateTP(intCompanyID,strtitleU, strsurname,strmiddle,strfirst,strExemail,strmobile,strusernameU,strpass, intdepartment, dblLevel, dblOT,strNote, strUserid)
	            end if
    		
	        case "DELETE"
	            gMessage=DeleteTP(strUserid)
	            if gMessage="" then Response.Redirect("listofcontractstaff.asp")
    	    
	    end select
	  
	  else

	        strSql="SELECT TPUserID, UserName, Title, FirstName, MiddleName, LastName, EmailAddress_Ex, MobilePhone, IDNumber, CompanyID, LevelID, DepartmentID, OvertimeID, Note FROM HR_TPStaff WHERE TPUserID=" & strUserid
        	
	        Call GetRecordset(strSql,rsSrc)
        	
	        if rsSrc.RecordCount>0 then
	            strTitleU=rsSrc("Title")	
	            strsurname=rsSrc("LastName")
	            strmiddle=rsSrc("MiddleName")
	            strfirst=rsSrc("FirstName")
	            strExemail=rsSrc("EmailAddress_Ex")
	            strmobile=rsSrc("MobilePhone")
	            stridnum=rsSrc("IDNumber")
	            intCompanyID=rsSrc("CompanyID")
	            dblLevel=rsSrc("LevelID")
	            intdepartment=rsSrc("DepartmentID")
	            dblOT=rsSrc("OvertimeID")
	            strNote=rsSrc("Note")
                strusernameU=rsSrc("UserName")
	        else
	            strTitleU=""	
	            strsurname=""
	            strmiddle=""
	            strfirst=""
	            strExemail=""
	            strmobile=""
	            stridnum=""
	            dblLevel=-1
	            intdepartment=-1
	            dblOT=-1
	            strNote=""
	            strusername=""
	        end if
	        
	   end if
        	
'----------------------------------------
' Prepare form
'----------------------------------------
    Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		
		ret = objDb.runQuery("SELECT CompanyID, CompanyName FROM ATC_Companies WHERE ([Type]=2) ORDER BY CompanyName")
		strOut1 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut1 = "<select name='lstCompany' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intCompanyID) then strSel=" selected" else strSel="" end if
			    strOut1 = strOut1 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut1 = strOut1 & "</select>"
		end if
		
		ret = objDb.runQuery("SELECT * FROM ATC_TPLevel WHERE fgActivate=1 ORDER BY LevelName")
		strOut3 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut3 = "<select name='lstLevel' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if CDbl(objDb.rsElement(0))=CDbl(dblLevel) then strSel=" selected" else strSel="" end if
			    strOut3 = strOut3 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut3 = strOut3 & "</select>"
		end if
		
		ret = objDb.runQuery("SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department")
		strOut4 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut4 = "<select name='lstdepartment' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if objDb.rsElement(0)=int(intdepartment) then strSel=" selected" else strSel="" end if
			    strOut4 = strOut4 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut4 = strOut4 & "</select>"
		end if
		
		strQuery = "SELECT  OvertimeID , OTName FROM  ATC_TPOverTime WHERE fgActivate =1 ORDER BY OTName"
		ret = objDb.runQuery(strQuery)
        strOut5 = ""
		if not ret then 
			gMessage = objDb.strMessage
		else
			strOut5 = "<select name='lstOT' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px' onFocus='CheckMode(this)'>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
				if cdbl(objDb.rsElement(0))=CDbl(dblOT) then strSel=" selected" else strSel="" end if
			    strOut5 = strOut5 & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
			    objDb.MoveNext
			  Loop
			end if
			strOut5 = strOut5 & "</select>"
		end if
		
	else
		'error in connection
		gMessage = objDb.strMessage
	end if
	objDb.dbdisConnect
	set objDb = nothing
	
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

<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css">

<script type="text/javascript" src="../../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../../jQuery/jquery-ui.min.js"></script>

<link href="../../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<style type="text/css">

fieldset{
    margin: 			0;  
	padding: 			0;
    display:            inline;
    border:             none;
    vertical-align:     top;
}
  
fieldset legend{
	padding-bottom: 	20px;
}

fieldset ol {
    padding: 			0; 
    margin:             0;
	list-style: 		none;
}

fieldset li {  
    padding:0;
	padding-bottom: 	0.5em;
}

fieldset li label 
{	
	width:			    80px;
	clear:				left;
	float:				left;
}

#submit ul
{
      list-style:none;
      text-align:center;
      padding:0;
}

#submit ul li
{    
    margin-left:1px;
    display:inline;
}

#submit ul li a
{
    padding-top:3px;
    display:inline-block;
    width:60px;
    height:22px;
    background-color:#8CA0D1;
    text-align:center;
    font-weight: bold;
    text-decoration:none;
}

#submit ul li a:hover
{
    background-color:#7791D1;
    color:white;
}

</style>

<script type="text/javascript" src="../../library/library.js"></script>

<script type="text/javascript">

    $(document).ready(function() {
        var intUser =<%=strUserid%>;
        //alert((intUser==-1));
        if (intUser!=-1)
            $("#passinput").hide();
        else
            $("#passlink").hide();
            
        $("#lnkPass").click(function(e) {
            $("#passinput").show();
            $("#passlink").hide();
        });
        
    })
//---------------------------------------------
    function _act(kind) {
        var act;
        act = 0;

         if (kind == "DELETE") {
               if (confirm("Are you sure you want to delete this Employee?")) act = 1;  }
          else //(kind == "SAVE") 
          {
             if (checkdata() == true) act = 1;                
                
            }
        
        if (act == 1) {
        
            document.frmdetail.action = "contractstaffdetail.asp?act=" + kind;
            document.frmdetail.target = "_self";
            document.frmdetail.submit();
        }
    }
    //---------------------------------------------

    function checkdata() {

        if (alltrim(document.frmdetail.txtsurname.value) == "") {
            alert("Please enter the surname.");
            document.frmdetail.txtsurname.focus();
            return false;
        }
        if (alltrim(document.frmdetail.txtfirst.value) == "") {
            alert("Please enter the first name.");
            document.frmdetail.txtfirst.focus();
            return false;
        }
        
        if (alltrim(document.frmdetail.txtusername.value) == "") {
            alert("Please enter the user name.");
            document.frmdetail.txtusername.focus();
            return false;
        }

        if (alltrim(document.frmdetail.txtExemail.value) != "") {
            if (!isemail(document.frmdetail.txtExemail.value)) {
                alert("Invalid value external email address \nValid format is: 'itsupport@atlasindustries.com'");
                document.frmdetail.txtExemail.focus();
                return false;
            }
        }

        return CheckPassword();
        
    }

    function CheckPassword() {
        var newPass, errMsg;
        //strPass = document.frmdetail.txtpass.value
        
        errMsg="";
        newPass=$('#txtpass').val();

        //alert(newPass);
        
        if (!$("#txtpass").is(":hidden")){           
           if (newPass.length < 8)
                    errMsg = errMsg + "- Be at least 8 characters \n";
                
                if (newPass.match(/\d/) == null)
                    errMsg = errMsg + "- At least <strong>one number \n";
                               
                if (newPass.match(/[!@#$%^&*-]/) == null)
                    errMsg = errMsg + "- At least one char in [!@#$%^&*-]";
                if (errMsg != "")
                    errMsg = " Password must meet the following requirements:\n" + errMsg ;
                    
                 if (errMsg!="") alert( errMsg);         
         }
          
        return (errMsg == "");
        
    }
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="frmdetail">

    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>

			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr> 
            <td style="padding:20px 20px 0 20px;"> 
              <%if gMessage<>"" then%>
               <div style="font-weight:bold; height:20px; background-color:#E7EBF5;" class="red"><%=gMessage%></div>
              <%end if%>
               <a class="blue" href="listofcontractstaff.asp" >Contact Staff List</a>
               <div class="title" style="padding:10px; text-align:center;">Contact Staff Details</div>
            </td>
          </tr>
          
          <tr> 
            <td height="100%" valign="top" style="padding:0 0 0 20px">            
               <fieldset >
                    <legend class="blue">Personal Information</legend>
                    <ol>                                              
						<li>
                            <label for="lsttitle" class="blue-normal">Title</label>
                                <select class="blue-normal" size="1" id="lsttitle" name="lsttitle"  ">
						              <option value="Mr" <% if ucase(trim(strtitleU))=ucase("Mr") then%> selected <%end if%>>Mr</option>
						              <option value="Mrs" <%if ucase(trim(strtitleU))=ucase("Mrs") then%> selected <%end if%>>Mrs</option>
						              <option value="Ms" <%if ucase(trim(strtitleU))=ucase("Ms") then%> selected<%end if%> >Ms.</option>
						            
					            </select></li>
					    <li>
					        <label for="txtsurname" class="blue-normal">Surname*</label>
					        <input type="text" id="txtsurname" name="txtsurname" maxlength="25" class="blue-normal" style="width:160px" value="<%=showlabel(strsurname)%>" <%if strMode<>"EDIT" then%> "<%end if%>></li>
					    <li>
					        <label for="txtmiddle" class="blue-normal">Middle Name</label>
					        <input type="text" id="txtmiddle" name="txtmiddle" class="blue-normal" maxlength="15" style="width:160px" value="<%=showlabel(strmiddle)%>" <%if strMode<>"EDIT" then%> "<%end if%>></li>
                        
                        <li>
                            <label for="txtfirst" class="blue-normal">First Name*</label>                        
                            <input type="text" id="txtfirst" name="txtfirst" class="blue-normal" maxlength="20" style="width:160px" value="<%=showlabel(strfirst)%>" <%if strMode<>"EDIT" then%> "<%end if%>></li>
                         
                           <li><label class="blue-normal" for="txtExemail">Ex. Email</label> 
                                <input type="text" id="txtExemail" name="txtExemail" class="blue-normal" maxlength="60" style="width:160px" value="<%=showlabel(strExemail)%>" <%if strMode<>"EDIT" then%> "<%end if%>></li>
                            <li><label class="blue-normal" for="txtmobile">Mobile Phone</label>  
                                <input type="text" id="txtmobile" name="txtmobile" class="blue-normal" maxlength="50" style="width:160px" value="<%=showlabel(strmobile)%>" <%if strMode<>"EDIT" then%> "<%end if%>></li>
                    </ol>
               </fieldset>                
               <fieldset style="padding-left:50px;">
                    <legend class="blue">Working Information</legend>
                    <ol>
                  
                       <li><label class="blue-normal" for="txtidnum">Company</label>
                            <%Response.Write strOut1%> </li>
                        <li><label class="blue-normal" for="txtidnum">User Name*</label>
                            <input type="text" id="txtusername" name="txtusername" maxlength="20" class="blue-normal"  style="width:160px" value="<%=showlabel(strusernameU)%>" <%if strMode<>"EDIT" then%> "<%end if%>></li>
                    </ol>                            
<%'if cint(strUserid)=-1 then %>
  <div id="passinput">
                    <ol>
                        <li><label class="blue-normal" for="txtpass">Password</label>
                            <input type="password" id="txtpass" name="txtpass" maxlength="15" class="blue-normal"  style="width:160px">
                        </li>
                     </ol>
</div>
<% 'else%>
  <div id="passlink">
    <ol>
                        <li><label class="blue-normal" for="lnkPass">&nbsp;</label>
                            <a href="#" name="lnkPass" id="lnkPass" >Reset Password</a>
                        </li>
                        </ol>
</div>                        
<% 'end if %>
                        <ol>

                           <li><label class="blue-normal">Level</label>
                            <%Response.Write strOut3%></li>
                        <li><label class="blue-normal">Department</label>
                            <%Response.Write strOut4%></li>
                        <li><label class="blue-normal">Overtime for</label>
                                <%Response.Write strOut5%> </li>
                        
                    </ol>
                </fieldset>
                 <fieldset>
                    <ol>
                        <li>
                            <label class="blue-normal" for="txtNote">Note</label>
                            <textarea id="txtNote" name="txtNote" rows="5" cols="90" class="blue-normal"><%=strNote %></textarea></li>
                    </ol>
                </fieldset>
                
<%if fgUpdate then%>                
               <div id="submit">
                    <ul>
                        <li><a href="javascript:_act('SAVE')">Save</a></li>
                        <%if strUserid<>-1 then %><li><a href="javascript:_act('DELETE');">Delete</a></li><%end if %>
                    </ul> 

               </div>
<%end if %>               
               </td>
          </tr> 
          </table>
			<%
			Response.Write(arrTmp(1))
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
<input type="hidden" name="txtuserid" value="<%=struserid%>" />
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>"/>
</form>
</body>
</html>