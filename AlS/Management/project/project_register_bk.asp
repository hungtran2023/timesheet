<!-- #include file = "project_function.asp"-->
<%
Dim fgError, objDatabase
'Dim objMail, strFrom, strTo, strSubject, strContent
'Declare by Uyen Chi
'User login and privilege for approval and regitry project.
Dim intUserID,fgRegister,fgApproval,fgViewAll
Dim varDepartment,arrlstDay(2),rsProject,rsCompany
dim strProjectIDkey,strDateTranfer,strProjectName,strProjectkey1,intDepartID,strProjectkey2,intCompanyID,strDescription,dblHourTransfer,strServerPath
dim strCSOFilename,dblCSOMainHours,dblValue,dblDailyRate,dblEstRemain
Dim strFullProjectID,strError,dateTransfer,strCompanyName,strCurrency
Dim strAction,fgUpdate,fgStatus
dim intManagerID,blnCSOApproval,intSignContract,blnUtilised,intOldManagerID
dim rsClientCountry,strClientCountry
dim blnisBIM,dblContingency, blnDesigned, dblBDM

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	fgError			= True
	strConnect = Application("g_strConnect")					' Connection string 				
	strAction= Request.QueryString("act")
	strFullProjectID=Space(15)

	Call GetInforProjectID(Request.Form("txthidden"),strFullProjectID,fgStatus,strDateTranfer,intOldManagerID)
	

	strProjectIDkey=ParseAPK(strFullProjectID)		

'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If

	intUserID = session("USERID")
'--------------------------------------------------
' Check VIEWALL project right
' User can update all project
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgViewAll = False
	Else
		varGetRight = session("RightOn")
		fgViewAll = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "View all projects" Then
				fgViewAll = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------
' Check Approving Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgApproval = False
	Else
		varGetRight = session("RightOn")
		fgApproval = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "approving project" Then
				fgApproval = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------
' Check Registration Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRegister = False
	Else
		varGetRight = session("RightOn")
		fgRegister = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "registration" Then
				fgRegister = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If
'--------------------------------------------------
' Check Invoicet right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgInvoice = False
	Else
		varGetRight = session("RightOn")
		fgInvoice = False
		For ii = 0 To Ubound(varGetRight, 2)
'Response.Write 	varGetRight(0, ii)		 & "<br>"
			If varGetRight(0, ii) = "Invoice" Then

				fgInvoice = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If		

'-----------------------------------
'Check ACCESS right
'-----------------------------------

	strTemp = Request.Form("txtpreviouspage")
	
	strFilename = strTemp
	If isEmpty(session("Righton")) Then
		fgRight = False
	Else
		getRight = session("Righton")
		fgRight = False
		For ii = 0 To Ubound(getRight, 2)
			If getRight(0, ii) = strTemp Then
				fgRight = True
				fgUpdate = False
				If getRight(1, ii) = 1 Then fgUpdate = True	'updateable right
				Exit For
			End If
		Next
		Set getRight = Nothing		
	End If	
	If fgRight = False Then
		Response.Redirect("../../welcome.asp")
	End If	

'--------------------------------------------------
' Initialize department array
'--------------------------------------------------	
If Not isEmpty(session("varDepartment")) Then
		varDepartment = session("varDepartment")
Else
		varDepartment = GetDepartment()
		if not isEmpty(varDepartment) then	session("varDepartment") = varDepartment
End If


'--------------------------------------------------
' Initialize client country recordset
'--------------------------------------------------	
		strSql="SELECT * FROM ATC_Countries WHERE fgActivate=1 ORDER BY CountryCode"	
		Call GetRecordset(strSql,rsClientCountry)
'--------------------------------------------------
' Initialize BDM recordset
'--------------------------------------------------	
		strSql="SELECT * FROM HR_BDM ORDER BY Firstname"	
		Call GetRecordset(strSql,rsBDM)

'--------------------------------------------------
' Analyse conditions
'--------------------------------------------------
'For Edit or add new that call from nListProject

if strAction="" then
	call LoadProject(strFullProjectID,strDateTranfer,strError,rsProject)

	if not rsProject.EOF then
		'Load existed project
		strProjectIDkey=ParseAPK(strFullProjectID)
	
		intDepartID=rsProject("DepartmentID")		
		strProjectName=rsProject("ProjectName")

		strDateTranfer=cdate(rsProject("DateTransfer"))
		strCompanyName=rsProject("CompanyName")
		strServerPath=rsProject("SeverPath")
		dblHourTransfer=rsProject("HourTransfer")
		
		strProjectkey2=rsProject("ProjectKey2")
		if isnull(rsProject("ClientCountryCode")) then 
			strClientCountry=""
		else
			strClientCountry=rsProject("ClientCountryCode")
		end if

		blnisBIM=rsProject("isBIM")
		blnDesigned=rsProject("isDesign")
		dblBDM=rsProject("BDMID")
		
		strDescription=""
		'blnCSOApproval=rsProject("CSOApproval")
		intSignContract=rsProject("SignContract")
		if not isnull(rsProject("Description")) then strDescription=rsProject("Description")
		intManagerID=""
		if not isnull(rsProject("ManagerID")) then intManagerID=rsProject("ManagerID")
		'blnCSOCompleted=rsProject("CSOCompleted")
		'blnBillable=rsProject("Billable")
		dblCSOMainHours=rsProject("CSOMainHours")
		dblValue=rsProject("Value")
		dblDailyRate=rsProject("DailyRate")
		dblEstRemain=rsProject("EstRemaining")
		dblExRate=rsProject("ExchangeRate")
		'dblCWFValue=rsProject("CWFValue")
		strCurrency=rsProject("CurrencyCode")
		strCSOFilename=""
		if not isnull(rsProject("CSOFilename")) then strCSOFilename=rsProject("CSOFilename")
		'strCSOComment=""
		'if not isnull(rsProject("CSOComment")) then strCSOComment=rsProject("CSOComment")
		
		dblContingency=rsProject("Contingency")

	end if
else
'For submit	

	if fgStatus="" then
		strProjectIDkey(0)=Ucase(Request.Form("txtClientCode"))
		
	else	
		if fgStatus="New" then
			strProjectIDkey(1)=Request.Form("txtProjectNumber")
			strProjectIDkey(2)=Request.Form("txtVariation")
		end if
	end if	
	
	strProjectIDkey(3)=Request.Form("lstProType")

	if not InvoieAlready(ParseAPK(strFullProjectID)) then
		strProjectIDkey(4)=Request.Form("lstCountry")
		strProjectIDkey(5)=Request.Form("lstServiceType")
		strProjectIDkey(6)=Request.Form("lstSector")
	end if
	
	
	strProjectName=Request.Form("txtproject")		
	'dblHourTransfer=Request.Form("txthour")
	dblHourTransfer=0
	
	strDescription=Request.Form("txtdescription")
	intManagerID=cint(Request.Form("lstManager"))
	intDepartID=Request.Form("lbDepart")
	blnUtilised=Request.Form("chkUtilised")
	dblBDM=Request.Form("lbBDM")
	if dblBDM="" then dblBDM=null
	blnDesigned=(Request.Form("chkDesigned")<>"")
	
    if Request.Form("chkBIM")="" then 
		blnisBIM = false
	else
		blnisBIM = true
	end if		
		
	strCurrency=Request.Form("lbCurrency")	
	strClientCountry=Request.Form("lbClientCountryCode")
    
	strProjectkey2=5
	if blnUtilised<>"" then strProjectkey2=1

	strCompanyName=GetCompanyName(strProjectIDkey(0),strServerPath,intCompanyID)	
	
	Select Case strAction
	
		case "save"
			'For add new project
			If fgStatus="" then
				'Get new tranfer date when update to database.
				strDateTranfer=cdate(Request.Form("lstmonth") & "/" & Request.Form("lstday") & "/" & Request.Form("lstyear"))
						
				call AddnewProject(strProjectIDkey, strProjectName,strDateTranfer,intDepartID ,intManagerID,strDescription,strProjectkey2,intUserID,blnisBIM,blnDesigned, dblBDM, strError)
			'For approval project
			elseif fgStatus="New" then
				'Get new tranfer date when update to database.
				strDateTranfer=cdate(Request.Form("lstmonth") & "/" & Request.Form("lstday") & "/" & Request.Form("lstyear"))
				Call ApprovalProject(strFullProjectID,strProjectIDkey,strProjectName,strDescription,intManagerID,strDateTranfer,strProjectkey2,blnisBIM,blnDesigned,dblBDM,strError)
			'For Edit project
			else
				If fgStatus = "Issued" Then					
					'blnCSOCompleted=Request.Form("chkCSOCompleted")
					'if 	blnCSOCompleted ="" then blnCSOCompleted=0
					
					dblCSOMainHours=Request.Form("txtCSOMainHours")
					if dblCSOMainHours="" then dblCSOMainHours=null
					
					dblValue=Request.Form("txtCost")
					if dblValue="" then dblValue=null
					
					dblDailyRate=Request.Form("txtDailyRate")
					if dblDailyRate="" then dblDailyRate=null	
									
					dblExRate=Request.Form("txtExRate")
					if dblExRate="" then dblExRate=null
								
					dblEstRemain=Request.Form("txtEstRemain")
					if dblEstRemain="" then dblEstRemain=null	
					
					dblContingency=Request.Form("txtContingency")
					if dblContingency="" then dblContingency=null	
					'strCSOComment=Request.Form("txtCSOComment")
					
					'blnCSOApproval=Request.Form("chkCSOApproval")
					'if blnCSOApproval="" then blnCSOApproval=0
					
					intSignContract=Request.Form("radContract")
					if intSignContract="" then intSignContract=0	
					
					strCSOFilename=	Request.Form("txtCSOFileName")			
					
				end if
				
				strDateTranferNew=cdate(Request.Form("lstmonth") & "/" & Request.Form("lstday") & "/" & Request.Form("lstyear"))

'Response.end						
				Call UpdateProject(strFullProjectID,strProjectIDkey,strDateTranfer,intDepartID,strProjectkey2,strProjectName,_
						strDescription,intManagerID,strDateTranferNew,dblHourTransfer,fgApproval,dblCSOMainHours,_
						intSignContract,dblValue,dblDailyRate,dblExRate,dblEstRemain,_
						strClientCountry,strCurrency,strCSOFilename,blnisBIM,blnDesigned,dblContingency,dblBDM,strError)

			end if

			'Send email notification to Manager
			if strError="" AND cint(intManagerID)<>0 then				
				if (cint(intManagerID)<>cint(intOldManagerID) and fgStatus<>"") OR fgStatus="New" then call SendEmailForManager(strProjectIDkey,strProjectName,intManagerID)
			end if		
		
			if strError="" then Response.Redirect("n_projectlist.asp?b=1")
			
		case "del"
			'Delete project
			if BeAbleToDelete(strFullProjectID,strDateTranfer) then
				DeleteProject(strFullProjectID)
				if strError="" then Response.Redirect("n_projectlist.asp?b=1")
			Else
				strError="You can not delete this project because it is used."
			end if
	end select	

end if

'-----------------------------------------------------------------
'Get List for tranfer date
'-----------------------------------------------------------------
arrlstDay(0) = selectmonth("lstmonth", month(strDateTranfer), -1)
arrlstDay(1) = selectday("lstday", day(strDateTranfer), -1)
arrlstDay(2) = selectyear("lstyear", year(strDateTranfer), 2000, year(now()), -1)
'-----------------------------------------------------------------
'Get List of Manager
'-----------------------------------------------------------------
strQuery = "SELECT DISTINCT a.UserID, e.Firstname + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') as Fullname " &_
			"FROM ATC_UserGroup a LEFT JOIN ATC_Group b ON a.GroupID = b.GroupID " &_
			"LEFT JOIN ATC_Permissions c ON b.GroupID = c.GroupID " &_
			"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID " &_
			"LEFT JOIN ATC_PersonalInfo e ON a.UserID = e.PersonID " &_
			"WHERE d.Description = 'Manager' AND e.FirstName <> 'Managers' AND e.fgDelete = 0 ORDER BY Fullname"
'response.write 		strQuery
Set objDb = New clsDatabase
ret = objDb.dbConnect(strConnect)
ret = objDb.runQuery(strQuery)

strOut = ""
if not ret then 
	gMessage = objDb.strMessage
else
	strOut = "<select name='lstManager' class='blue-normal' style='HEIGHT: 22px; WIDTH: 228px'>"
	if intManagerID="" then strSel=" selected" else strSel="" end if
	strOut = strOut & "<option value='0'" & strSel & "></option>"
	if not objDb.noRecord then
	  Do Until objDb.rsElement.EOF
		if objDb.rsElement(0)=int(intManagerID) then strSel=" selected" else strSel="" end if
	    strOut = strOut & "<option value='" & objDb.rsElement(0) & "'" & strSel & ">" & showlabel(objDb.rsElement(1)) & "</option>"
	    objDb.MoveNext
	  Loop
	end if
	strOut = strOut & "</select>"
end if


'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	'for new project, the first key is user'department by default
	if intDepartID=-1 then intDepartID=varFullName(2)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
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
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", session("CurPage"))
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", session("NumPage"))	
End If

%>
<html>
<head>
<title>Timesheet System - Project Detail</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>
<script language="javascript">
<!--

function checkclient()
{
	if (isnull(document.frmreport.txtClientCode.value) == false)
	{
		if (document.frmreport.txtClientCode.value.length < 3)
		{
			alert("The length of Client Prefix accept 3 characters.\n Please check again.");
			document.frmreport.txtClientCode.focus();
		}
		else
		{	
			document.frmreport.action = "project_register.asp?act=sc"
			document.frmreport.submit();
		}	
	}
}

function back_menu()
{
	window.document.frmreport.action = "n_projectlist.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

<%If fgStatus <> "" and fgStatus <> "New" then%>
function SearchAgain()
{
	var tmp = document.frmreport.txtsearch.value;
	tmp = alltrim(tmp);
	document.frmreport.txtsearch.value = tmp;
	tmp = escape(tmp);
	
	window.document.frmreport.action = "n_projectlist.asp?search=" + tmp;
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}
<%end if%>
function checkdata()
{
var strkey3;
var fgStatus="<%=fgStatus%>";
var fgApproval="<%=fgApproval%>"
var companyname="<%=strCompanyName%>";
var dateTranfer=document.frmreport.lstday.value + "/" + document.frmreport.lstmonth.value + "/" + document.frmreport.lstyear.value;

	if (fgStatus=="")
	{
		var tmp = alltrim(document.frmreport.txtClientCode.value);
		document.frmreport.txtClientCode.value = tmp;
		if (tmp=="") {
			alert("Please enter the client code.");
			document.frmreport.txtClientCode.focus();
			return false;
		}		
		if (companyname == "")
		{
			alert("Please add new company for this project." );
			return false;
		}
		if (document.frmreport.lstProType.value==0){
			alert("Please select the project type.");
			document.frmreport.lstProType.focus();
			return false;
		}
		if (document.frmreport.lstCountry.value==0){
			alert("Please select the country.");
			document.frmreport.lstCountry.focus();
			return false;
		}
				
		if (document.frmreport.lstServiceType.value==''){
			alert("Please select the service type.");
			document.frmreport.lstServiceType.focus();
			return false;
		}
		
		if (document.frmreport.lstSector.value==''){
			alert("Please select the project sector type.");
			document.frmreport.lstSector.focus();
			return false;
		}
	}
	if (fgStatus=="New"){
		if (document.frmreport.txtProjectNumber.value==""){
			alert("Please select the project number.");
			document.frmreport.txtProjectNumber.focus();
			return false;
		}
		
		if (document.frmreport.txtVariation.value==""){
			alert("Please select the variation value.");
			document.frmreport.txtVariation.focus();
			return false;
		}
	}
	if (isnull(document.frmreport.txtproject.value) == true)
	{
		alert("Project Name must not be empty.");
		document.frmreport.txtproject.focus();
		return false;
	}
	

	if(isdate(dateTranfer)==false) {
		alert("The date tranfer (" + dateTranfer + ") is invalid.");
		document.frmreport.lstday.focus();
		return false;
	}
	
	if (document.frmreport.lbDepart.value == 0){
		alert("Please select the department.");
		document.frmreport.lbDepart.focus();
		return false;
	}
	return true;
}

function AddCSODetails(varid)
{
	//window.document.frmreport.txthidden.value=varid;
	window.document.frmreport.action = "pro_CSODetails.asp";
	window.document.frmreport.submit();
}


function actpro(kind)
{
var intkey2 = "<%=intProjectKey2%>";
var answer;

	answer=true;
	if (kind=="del")
		answer = confirm("Are you sure you want to detele this project?")
	else if (kind=="save")
	{
		answer=checkdata();
		if ( answer == true)
			if (intkey2!="" && intkey2 != document.frmreport.lbkey2.value)
				answer = confirm("Are you sure you want to change the second field of project key?")
		
	}
	if (answer==true)
	{
		document.frmreport.action = "project_register.asp?act=" + kind;
		document.frmreport.submit();
	}	
}
function getStats(fName){

	fullName = fName;
	shortName = fullName.match(/[^\/\\]+$/);
	document.forms.frmreport.txtCSOFileName.value = shortName
	
	if (document.all)
		divCSOFileNameDisplay.innerHTML=shortName
	else if (document.getElementById)
		document.getElementById("divCSOFileNameDisplay").innerHTML=shortName
	else if (document.layers){
	
		document.divCSOFileNameDisplay.document.write(shortName)
		document.divCSOFileNameDisplay.document.close()
		}
	
}

<%if fgStatus = "" then %>
var objNewWindow;
function addcontact() { //v2.0
var intheight;
//thao
	var tmp = alltrim(document.frmreport.txtClientCode.value);
	document.frmreport.txtClientCode.value = tmp;
	if (tmp==""){
		alert("Please enter the client prefix.");
		document.frmreport.txtClientCode.focus();
	}
	else{
		intheight = 215
				
		window.status = "";
		strFeatures = "top="+(screen.height/2-180)+",left="+(screen.width/2-150)+",width=302,height=" + intheight + ",toolbar=no," 
		              + "menubar=no,location=no,directories=no,resizable=no";
		if((objNewWindow) && (!objNewWindow.closed))
			objNewWindow.focus();	
		else {
			objNewWindow = window.open("addcompany.asp?charcode=" + "<%=Request.Form("txtClientCode")%>", "MyNewWindow", strFeatures);
		}
		window.status = "Opened a new browser window.";
	}
}

function window_onunload() {
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.close();
}
<%end if%>
//-->
</script>

    <style type="text/css">
        .style1
        {
            font-size: 8pt;
            color: #003399;
            font-family: arial, verdana;
            text-decoration: none;
        }
    </style>

</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" <%if fgStatus = "" then %>onUnload="return window_onunload();"<%end if%>>
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
		  <tr> 
			<td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr bgcolor=<%if strError="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
				  <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
		        </tr>
                <tr align="middle"> 
                  <td class="blue" height="10" align="left" width="50%"> &nbsp;&nbsp;
					<A href="javascript:back_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">Project List</a>
<%					If fgStatus <> "" and fgStatus <> "New" then%>
					 | <A href="javascript:SearchAgain();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">Search again</a>  
<%					End if%>					
                  </td>
                  <td class="blue" height="30" align="right" width="77%">&nbsp;</td>
                </tr>
                <tr align="middle"> 
                  <td class="title" height="50" align="middle" colspan="2">Project Information</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="HEIGHT: 79%" height="365">
                <tr> 
                  <td bgcolor="#ffffff" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td class="blue-normal" width="12%">&nbsp;</td>
                        <td class="blue-normal" width="19%"><b><u>APK *</u><b></td>
                        <td width="69%" valign="center"class="blue" >
                          <%=strFullProjectID%>
                        </td>
                      </tr>
<%
	
	Response.Write ProjectKeyHTMLForNew(fgStatus,fgApproval,strProjectIDkey)
	
%>                                                              
					<tr> 
                        <td class="style1"></td>
                        <td class="style1"><b><u>Project Information</u><b></td>
                        <td valign="center"></td>
                      </tr>
                    <tr> 
                        <td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal">Project Description*</td>
                        <td > 
                          <input name="txtproject" size="32" style="WIDTH: 228px; HEIGHT: 20px" class="blue-normal" value="<%=strProjectName%>" maxlength="120">
                        </td>
                      </tr>  
                      <tr> 
						<td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal"></td>
                        <td class="blue-normal"> 
							<input type="checkbox" name="chkUtilised" value="1" <%if strProjectkey2=1 then%>checked<%end if%> />Utilised&nbsp;
							<input type="checkbox" name="chkBIM" value="1" <%if blnisBIM then %>checked<%end if%>/>BIM &nbsp;
							<input type="checkbox" name="chkDesigned" value="1" <%if blnDesigned then %>checked<%end if%>/>Design &nbsp;<p>
                        </td>
                      </tr>
                      <tr> 
                        <td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal">Date Tranfer*</td>
                        <td > 
<%
Response.Write arrlstDay(1)
Response.Write arrlstDay(0)
Response.Write arrlstDay(2)
%>                             
                        </td>
                      </tr>      
                      <tr> 
                        <td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal">Department * </td>
                        <td>
<%
		strOutDepart=""
		strOutDepart= strOutDepart & "<select class='blue-normal' name='lbDepart' style='WIDTH: 228px; HEIGHT: 20px'><option value='0'>&nbsp;</option>"
		for ii=0 to UBound(varDepartment,2)
			strSelect=""
			if cint(varDepartment(0,ii))=cint(intDepartID) then strSelect="selected"
			strOutDepart= strOutDepart & "<option value='" & varDepartment(0,ii) & "'" & strSelect & ">" & showlabel(varDepartment(1,ii)) & "</option>"
		next
		strOutDepart = strOutDepart & "</select>"
		Response.Write strOutDepart
		
%>
						</td>
                      </tr> 
					  
					  
					  <tr> 
                        <td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal">Project BDM </td>
                        <td>

						<select class='blue-normal' name='lbBDM' style='WIDTH: 228px; HEIGHT: 20px'>
								<option value="" <%if dblBDM="" then%>selected<%end if%>>&nbsp;</option>
							<%do while not rsBDM.EOF%>
							
								<option value="<%=rsBDM("BDMID")%>" <%if dblBDM=rsBDM("BDMID") then%>selected<%end if%>><%=rsBDM("Fullname")%></option>
							<% rsBDM.Movenext
							loop%>
							</select>
						</td>
                      </tr> 
					  
					  
					  
<%		If fgStatus <> ""  then%>
                      <tr> 
                        <td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal">Manager</td>
                        <td><%Response.Write strOut%></td>
                      </tr>    
<%		end if%>                                              
<%	

		If fgStatus = "Issued" And (fgApproval OR fgInvoice) Then
%>                      
                      <tr> 
                        <td>&nbsp;</td>
<td class="blue-normal">Project Country</td>
                        <td>
							<select class='blue-normal' name='lbClientCountryCode' style='WIDTH: 228px; HEIGHT: 20px'>
								<option value="" <%if strClientCountry="" then%>selected<%end if%>>&nbsp;</option>
							<%do while not rsClientCountry.EOF%>
							
								<option value="<%=rsClientCountry("CountryCode")%>" <%if strClientCountry=rsClientCountry("CountryCode") then%>selected<%end if%>><%=rsClientCountry("CountryCode")%> - <%=rsClientCountry("CountryName")%></option>
							<% rsClientCountry.Movenext
							loop%>
							</select>
                        </td>
                      </tr>
			<input type="hidden" name="txthour" value="<%=dblHourTransfer%>">
<%		End If%>                                           
                      <tr> 
                        <td colspan="3">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">Company name*</td>
                        <td class="blue-normal">
							<table width="250px" border="0" cellspacing="0" cellpadding="0">
							 <tr>
								<td>
<%

Response.Write Replace(TableField(strCompanyName),"100%","228px")
if strServerPath<>"" then Response.Write Replace(TableField(strServerPath),"100%","300px")
%> 
								</td>
								<td>
						<%if fgStatus = "" and strCompanyName="" then %><a href="javascript:addcontact()">Add... </a>
<%end if%>
								</td>
								</tr>
							</table>
                        </td>
						
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">Details</td>
                        <td class="blue-normal" valign="center">                  
                          <TEXTAREA class="blue-normal" name="txtdescription" rows="4" cols="56"><%=strDescription%></TEXTAREA>
                        </td>
                      </tr>
<%If fgStatus <> "" and fgStatus <> "New" then%>
                      <tr> 
						<td>&nbsp;</td>
                        <td colspan="2" class="blue-normal"><b><u>CSO Information</u></b></td>
                      </tr>
                      
	<%If fgStatus = "Issued" And (fgApproval OR fgInvoice) Then
							'strSql="SELECT * FROM ATC_Currency WHERE fgActivate=1"
							strSql="SELECT Currency as CurrencyCode, RateToUSD FROM rp_CSOExchangeRate ORDER BY Currency" 
							Call GetRecordset(strSql,rsCurrency)
							
							'strCurrHtml=PopulateDataToList("lstCurrency",rsCurrency,"CurrencyCode"," ",strCurrency)
	%>

		<%If strProjectIDkey(3)<>"T" THEN%>
					
						<tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">Project value</td>
                        <td class="blue-normal">
                          <input name="txtCost" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblValue%>">&nbsp;
		<%elseif strProjectIDkey(3)="T" then%>                    
						<tr> 
                        <td >&nbsp;</td>
                        <td class="blue-normal">Daily Rate</td>
                        <td class="blue-normal">
                          <input name="txtDailyRate" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblDailyRate%>">&nbsp;
		<%end if%>
						<select class='blue-normal' name='lbCurrency' style='WIDTH: 72px; HEIGHT: 20px'>
								<option value="" <%if strCurrency="" then%>selected<%end if%>>&nbsp;</option>
		<%do while not rsCurrency.EOF%>
							
								<option value="<%=rsCurrency("CurrencyCode")%>" <%if strCurrency=rsCurrency("CurrencyCode") then%>selected<%end if%>><%=rsCurrency("CurrencyCode")%></option>
		<% rsCurrency.Movenext
		loop%>
							</select>
						</td>
                      </tr>  
                      <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">Exchange Rate</td>
                        <td class="blue-normal">
                       
                          <input name="txtExRate" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblExRate%>">&nbsp;&nbsp;&nbsp; 
                         </td>
                      </tr>
                      
					   <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">CSO Hours</td>
                        <td class="blue-normal">
                          <input name="txtCSOMainHours" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblCSOMainHours%>">&nbsp;			
                          <i><b><a href="javascript:AddCSODetails(&quot;<%=strFullProjectID%>&quot;)">CSO detail... </a></b></i>
                        </td>
                      </tr>
                       <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">Estimated Remaining</td>
                        <td class="blue-normal">
                          <input name="txtEstRemain" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblEstRemain%>"/>&nbsp;&nbsp;&nbsp; 
                         (hrs)</td>
                      </tr>
                      
                      <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">Contingency</td>
                        <td class="blue-normal">
                          <input name="txtContingency" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblContingency%>"/>&nbsp;&nbsp;&nbsp; 
                         %</td>
                      </tr>
					   <!--
					   <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">CWF</td>
                        <td class="blue-normal">
                          <input name="txtCWF" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%'=dblCWFValue%>">&nbsp;
                        </td>
                      </tr>                      
					   <tr> 
                        <td >&nbsp;</td>
                        <td class="blue-normal"></td>
                        <td class="blue-normal">
                          <input type="checkbox" name="chkCSOCompleted" value="1" <%if blnCSOCompleted then%>checked<%end if%>>CSO Completed&nbsp; 
                          <input type="checkbox" name="chkCSOApproval" value="1" <%if blnCSOApproval then%>checked<%end if%>>FD Approval&nbsp; 
                          <input type="checkbox" name="chkBillable" value="1" <%if blnBillable then%>checked<%end if%>>Billable&nbsp;
                        </td>
                      </tr>  -->        

					  <tr> 
                        <td >&nbsp;</td>
                        <td  class="blue-normal">CSO filename</td>
                        <td  class="blue-normal" valign="center">                  
                          <input name="txtCSOFileName" type="hidden" value="<%=strCSOFilename%>">
		<% if strCSOFilename <>"" then %> 
                          <span id="divCSOFileNameDisplay" name="divCSOFileNameDisplay" style="background-color:#CCCCCC; width:72px;HEIGHT: 20px" ><%=strCSOFilename%></span>
		<%end if %>
                          <br>
                          <input type="file" name="fileCSO" class="blue-normal" id="file" onChange="getStats(this.value)"/>
                          
                        </td>
                      </tr>

                      <tr>
						<td >&nbsp;</td>
                        <td class="blue-normal">CSO</td>
                        <td class="blue-normal"> &nbsp;
                          <input type="radio" name="radContract" value="0" <%if cint(intSignContract)=0 then%>checked<%end if%>>No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          <input type="radio" name="radContract" value="1" <%if cint(intSignContract)=1 then%>checked<%end if%>>Yes, not signed&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          <input type="radio" name="radContract" value="2" <%if cint(intSignContract)=2 then%>checked<%end if%>>Signed</td>
                      </tr>                                            
<%	else%>
						<tr> 
                        <td >&nbsp;</td>
                        <td class="blue-normal">CSO Hours</td>
                        <td class="blue-normal">
                          <input name="txthourtemp" size="10" style="WIDTH: 72px; HEIGHT: 20px" class="blue-normal" value="<%=dblCSOMainHours%>" disabled >&nbsp;&nbsp;&nbsp;
                          <input name="txtCSOMainHours" type="hidden"  value="<%=dblCSOMainHours%>"> 
						  <input name="lbClientCountryCode" type="hidden"  value="<%=strClientCountry%>"> 
						  <input name="txtCSOFileName" type="hidden"  value="<%=strCSOFilename%>"> 
						  
						  <input name="lbCurrency" type="hidden"  value="<%=strCurrency%>">
<!--							<input type="checkbox" name="chkCSOCompletedTemp" value="1" <%if blnCSOCompleted then%>checked<%end if%> disabled>CSO Completed&nbsp;
							<input type="checkbox" name="chkCSOApprovalTemp" value="1" <%if blnCSOApproval then%>checked<%end if%> disabled>FD Approval&nbsp;
							<input type="checkbox" name="chkBillableTemp" value="1" <%if blnBillable then%>checked<%end if%> disabled>Billable
							<input name="chkCSOCompleted" type="hidden" value="<%if blnCSOCompleted then%>1<%else%>0<%end if%>">
							<input name="chkCSOApproval" type="hidden" value="<%if blnCSOApproval then%>1<%else%>0<%end if%>">                          
							<input name="chkBillable" type="hidden" value="<%if blnBillable then%>1<%else%>0<%end if%>">-->
                        </td>
                      </tr>
                      <tr>
						<td>&nbsp;</td>
                        <td class="blue-normal">&nbsp;</td>
                        <td class="blue-normal">							
							<input name="radContract" type="hidden" value="<%=intSignContract%>">
                        </td>
                      </tr>
		<%if strCSOComment<>"" then%>
					  <tr> 
                        <td>&nbsp;</td>
                        <td class="blue-normal">CSO Comments</td>
                        <td class="blue-normal" valign="center"><%=strCSOComment%>
                        </td>
                      </tr>						
		<%End if%>
                        <input name="txtCSOComment" type="hidden" value="<%=strCSOComment%>">
	<%End if%>
						
<%End if%>
                      <tr> 
                        <td colspan="3" height="40">
<%
'Response.Write fgApproval & "_" & fgStatus & "_" & fgRegister & "_" & fgUpdate
'Response.Write (fgViewAll OR intOldManagerID = Session("UserID")) AND fgUpdate
fgUpdate=fgUpdate AND (fgViewAll OR cint(Session("UserID"))=cint(intOldManagerID))
'Response.Write fgUpdate & "-" & fgApproval
Response.Write (NavigationHTML(fgApproval,fgStatus,fgRegister,fgUpdate))
%>                                 
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
			</td>
		  </tr>
		</table>
<%		
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
	Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(2))    
%>
<input type="hidden" name="txthidden" value="<%=Request.Form("txthidden")%>"> 
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
<input type="hidden" name="txtpreviouspage" value="<%=Request.Form("txtpreviouspage")%>">
<input type="hidden" name="txtstaff" value="<%=intOwnerID%>">	
<%If fgStatus <> "" and fgStatus <> "New" then%>
	<input type="hidden" name="txtsearch" value="<%=Request.Form("txtsearch")%>">
	<input type="hidden" name="lbSeachType" value="<%=Request.Form("lbSeachType")%>">
	<input type="hidden" name="lbBill" value="<%=Request.Form("lbBill")%>">
	<input type="hidden" name="lbBooked" value="<%=Request.Form("lbBooked")%>">
<%end if%>
</form>
</body>
</html>
