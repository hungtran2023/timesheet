<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim intRegisterID,strRegisterName,intTypeID,intNumberOfLicence,strVendor
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon

'***************************************************************
'
'***************************************************************
function OutBody(rsSrc,intPage,PageSize)
	
	dim intStart,intFinish
	dim strOut
	dim i
	strOut = ""

	if not rsSrc.EOF then
		
		rsSrc.AbsolutePage = intPage
		intStart = rsSrc.AbsolutePosition
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsData.PageSize - 1)
		End if
	
		For i = intStart to intFinish
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
					
			strOut=strOut & "<tr bgcolor='" & strColor & "'>" 
			strOut=strOut & "<td valign='top' class='blue-normal'>&nbsp;" & i & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("SoftwareName") & "</td>" 
			strOut=strOut & "<td valign='top' class='blue-normal' align='center'><input type='checkbox' name='chkRemove' value='" & rsSrc("LicenceInstallSWID") & "'></td>" 		
			strOut=strOut & "</tr>"
			
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		next
	end if
	
	OutBody=strOut
End Function

'***************************************************************
'
'***************************************************************

function GetTypeOfListBox(rsSrc,intTypeID)
	dim strOut
	
	strOut=""
	
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
									
			strSelect=""		
			if isnull(intSoftTypeID) then intSoftTypeID=0
			if cint(rsSrc("LicenceTypeID")) =cint(intTypeID) then strSelect="selected"			
			
			strOut=strOut & "<option value='" & rsSrc("LicenceTypeID") & "' "  & strselect & " >" & rsSrc("TypeName")  & "</option>"
			rsSrc.MoveNext
		loop
		
	end if

	GetTypeOfListBox=strOut
end function
'***************************************************************
'
'***************************************************************
function ExecuteSQL(strSql)

	dim objDatabase
	dim strCnn
	dim blnReturn
	
	blnReturn=false	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	If objDatabase.dbConnect(strCnn) then		
		blnReturn= (objDatabase.runActionQuery(strSql))	
		strError="Update successfull."
		if not blnReturn then strError=objDatabase.strMessage		
	else
		strError=objDatabase.strMessage
	end if
	
	Set objDatabase = nothing
	ExecuteSQL=strError
	
end function

'***************************************************************
'
'***************************************************************
function AddRegister(strLicenceSoftware,intTypeOfLicence,strVendor,intNumberOfLicence,strNote)
		
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "InsertALicenceSoftware"

		Set myParam = myCmd.CreateParameter("LicenceSoftware", adVarChar,adParamInput,100)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("TypeOfLicence",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("NumberOfLicence", adInteger,adParamInput)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("Vendor", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam			
		Set myParam = myCmd.CreateParameter("LincenceNote", adVarChar,adParamInput,2000)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("LicenceRegisterID", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("intErrorCode", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("LicenceSoftware")	= strLicenceSoftware
		myCmd("TypeOfLicence")		= intTypeOfLicence
		myCmd("NumberOfLicence")    = intNumberOfLicence
		myCmd("Vendor")             =strVendor
		myCmd("LincenceNote")		= strNote

		myCmd.Execute

		If Err.number > 0 Then
			strError= Err.Description
		Else
			if myCmd("intErrorCode")>0 then
				strError="The licence software name is already existed."
			else
				strError = "New licence software adding successfull"
				intLicenceRegisterID=myCmd("LicenceRegisterID")
			end if
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	AddRegister=strError
	
end function


'***************************************************************
'
'***************************************************************
function UpdateRegister(intLicenceRegisterID,strLicenceSoftware,intTypeOfLicence,strVendor,intNumberOfLicence,strNote)
		
'	Response.Write intLicenceRegisterID & "#" & strLicenceSoftware & "#" & intTypeOfLicence & "#" & intNumberOfLicence & "#" & strNote
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "UpdateALicenceSoftware"

		Set myParam = myCmd.CreateParameter("LicenceSoftware", adVarChar,adParamInput,100)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("TypeOfLicence",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("NumberOfLicence", adInteger,adParamInput)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("Vendor", adVarChar,adParamInput,50)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("LincenceNote", adVarChar,adParamInput,2000)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("LicenceRegisterID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("intErrorCode", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("LicenceSoftware")	= strLicenceSoftware
		myCmd("TypeOfLicence")		= intTypeOfLicence
		myCmd("NumberOfLicence")    = intNumberOfLicence
		myCmd("Vendor")             =strVendor
		myCmd("LincenceNote")		= strNote
		myCmd("LicenceRegisterID")	= intLicenceRegisterID

		myCmd.Execute
		
'Response.Write 	"<br>" & myCmd("intErrorCode")

		If Err.number > 0 Then
			strError= Err.Description
		Else
			if myCmd("intErrorCode")>0 then
				strError="The Licence Software is already existed."
			else
				strError = "Update successfull"
				
			end if
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
		
	UpdateRegister=strError	
end function

'***************************************************************
'
'***************************************************************
function DelRegister(intLicenceRegisterID)
		
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 
	If objDatabase.dbConnect(strCnn) Then
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "DeleteALicenceSoftware"
	
		Set myParam = myCmd.CreateParameter("LicenceRegisterID", adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("count", adInteger,adParamOutput)
		myCmd.Parameters.Append myParam

		myCmd("LicenceRegisterID")	= intLicenceRegisterID
		myCmd.Execute
		
		intCountTest=myCmd("Count")
	
		If Err.number > 0 Then
			strError= Err.Description
		Else
			IF cint(myCmd("Count"))>0 then
				strError="Please remove all softwares from this Licence."
			else
				Response.Redirect("SoftwareRegister.asp")
			end if
		End If
		Err.Clear
	
		set myCmd=nothing
	else
		strError=objDatabase.strMessage
	end if
	set objDatabase=nothing	
	
	DelRegister=strError
	
end function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then		Response.Redirect("../../message.htm")
	
	intUserID = session("USERID")

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	intRegisterID = Request.Form("txtID")
	fgDel=Request.Form("fgstatus")

	if Request.QueryString("act") = "save" then
		
		strRegisterName=Request.Form("txtRegisterName")	
		
		intSoftTypeID=Request.Form("lstCategory")
		if intSoftTypeID="" then intSoftTypeID=null
		
		strVendor =Request.Form("txtVendor")
		if strVendor="" then strVendor=null
		
		intNumberOfLicence=Request.Form("txtNoLicence")
		if intNumberOfLicence="" then intNumberOfLicence=0
		
		strNote=Request.Form("txtNote")
		if strNote="" then strNote=null
		
		intRegisterID=Request.Form("txtID")
 	
		if fgDel<>"D" then
			
			if Cint(intRegisterID)=-1 then
				'Add new				
				strError=AddRegister(strRegisterName,intSoftTypeID,strVendor,intNumberOfLicence,strNote)
			else			
				strError=UpdateRegister(intRegisterID,strRegisterName,intSoftTypeID,strVendor,intNumberOfLicence,strNote)			

			end if
		else		
			strError=DelRegister(intRegisterID)		
				
		end if
	elseif Request.QueryString("act") = "remove" then
	
		arrComputer=Request.Form("chkRemove")
		
		if trim(arrComputer)<>"" then
			strSql="DELETE FROM ATC_LicenceInstallSoftwares WHERE LicenceInstallSWID IN (" & arrComputer & ")"
			strError= ExecuteSQL(strSql)
		end if
		
	else
	
		strRegisterName=""
		strVendor=""
		strCategory=""
		strVendor=""
		intNumberOfLicence=""
		intSoftTypeID=0
	End If
'--------------------------------------------------
' 
'--------------------------------------------------
	
	strSql="SELECT [LicenceRegisterID],[LicenceSoftware],[TypeOfLicence],[NumberOfLicence],[Vendor],[LincenceNote] FROM ATC_LicenceSoftwares WHERE LicenceRegisterID=" & intRegisterID
	
	Call GetRecordset(strSql,rsData)

'if Request.Form("txtID")="" then Response.End
	
	if rsData.RecordCount>0 then
		
		strRegisterName=rsData("LicenceSoftware")
		intTypeID=rsData("TypeOfLicence")		
		intNumberOfLicence=rsData("NumberOfLicence")
		strVendor=rsData("Vendor")
		strNote=rsData("LincenceNote")
	

	end if

	strSql="SELECT * FROM ATC_LicenceTypes WHERE fgActivated=1 ORDER BY TypeName "
	Call GetRecordset(strSql,rsData)
	
	strSoftTypeList= GetTypeOfListBox(rsData,intTypeID)
			
	strSql="SELECT a.*,b.SoftwareName FROM ATC_LicenceInstallSoftwares a " & _
	            "INNER JOIN ATC_Softwares b ON a.SoftwareID=b.SoftwareID " & _
	        "WHERE LicenceRegisterID=" & intRegisterID & " ORDER BY SoftwareName"
			
	Call GetRecordset(strSql,rsData)

'--------------------------------------------------
'Start Paging
'--------------------------------------------------

' Set the PageSize, CacheSize and populate the intPageCount
	rsData.PageSize=20
' The Cachesize property sets the number of records that will be cached locally in memory	
	rsData.CacheSize=rsData.PageSize	
	intPageCount=rsData.PageCount
	intRecordCount=rsData.RecordCount
	
' Checking to make sure that we are not before the start or beyond end of the recordset
' If we are beyond the end, set the current page equal the last page of the recordset.
' If we are before the start, set the current page equal the start of the recordset
	
	intPage=Request.QueryString("Navi")

	if intPage="" then intPage=1
	
	if cint(intPage)>Cint(intPageCount) then intPage=intPageCount
	if cint(intPage)<=0 then intPage=1

'--------------------------------------------------
'End Paging	
'--------------------------------------------------	
		
	strLast=OutBody(rsData,intPage,PageSize)
	
	'intNumberOfLincences = CountLicences(rsData)

'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)					
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
	If strChoseMenu = "" Then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

	arrlstFrom(0) = selectmonth("lstmonthF",month(dateBuying) , -1)
	arrlstFrom(1) = selectday("lstdayF", day(dateBuying), -1)
	arrlstFrom(2) = selectyear("lstyearF", year(dateBuying), 1999, year(date())+2, 0)
	
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intPageCount)	
End if

%>	

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

function next() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage < numpage) {
	
		curpage=<%=intPage+1%>
		document.frmreport.action = "RegisterDetail.asp?navi=" + curpage;
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function prev() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	if (curpage > 1) {
		curpage=<%=intPage-1%>
		document.frmreport.action = "RegisterDetail.asp?navi=" + curpage;
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function go() {
var curpage = <%=intPage%>
var numpage = <%=intPageCount%>
	var intpage = document.frmreport.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmreport.action = "RegisterDetail.asp?navi=" + intpage;
		document.frmreport.target = "_self";
		document.frmreport.submit();		
	}
}

function AtlasNetwork()
{
	//window.document.frmreport.action = "AtlasComputer.asp"			
	//window.document.frmreport.submit();
}
	
function savedata()
{
	if (checkdata())
	{
		window.document.frmreport.action = "RegisterDetail.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "RegisterDetail.asp?act=save"			
	window.document.frmreport.submit();
}

function checkdata()
{
	if (window.document.frmreport.txtRegisterName.value=="")
	{
		alert("Please enter Register name.");
		document.frmreport.txtRegisterName.focus();
		return false	
	}
	if (isNaN(document.frmreport.txtNoLicence.value)==true) {
	    alert("Please enter a number.");
		document.frmreport.txtNoLicence.focus();
		return false;
	}
	else if (document.frmreport.txtNoLicence.value<0) {
		alert("This number must be greater than 0.");
		document.frmreport.txtNoLicence.focus();
		return false;			
	}
	
	return true	
}

function checkedAll (own) {

	var aa= document.getElementById('frmreport');
	var chkName
	
	chkName="chkRemove"
		
	for (var i =0; i < aa.elements.length; i++) 
	{
		strName=String(aa.elements[i].name)
		
		if (aa.elements[i].type == "checkbox" && strName.indexOf(chkName)>-1)
			aa.elements[i].checked = own.checked;
	}
}

function SetOfSoftware()
{	
	window.document.frmreport.action = "SoftwareforLicences.asp"
	window.document.frmreport.submit();
}

function removeSoftware()
{
	window.document.frmreport.action = "RegisterDetail.asp?act=remove"			
	window.document.frmreport.submit();
}


function UpdateLincence(id)
{
	window.document.frmreport.txtPCRegisterID.value=id
	//window.document.frmreport.action = "LicenceSoftware.asp?fr=sof"
	//window.document.frmreport.submit();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
<%		If strError <> "" Then%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red" colspan="2">&nbsp;<b><%=strError%></b></td>
				</tr>
<%		End If%>				
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="SoftwareRegister.asp" onMouseOver="self.status='';return true">Register List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
             
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Register Detail</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign=top> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="2%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%">Register name *</td>
                              <td valign="middle" width="82%" class="blue">
								<input type="text" name="txtRegisterName" maxlength="100" class="blue-normal" style="width:50%" value="<%=strRegisterName%>"></td>
                              <td valign="top" width="1%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Category</td>
                              <td valign="middle" >
								<select name="lstCategory" class="blue-normal" style="width:50%">
								<%=strSoftTypeList%>
								</select>
                              </td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr> 
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Vendor</td>
                              <td valign="middle">
								<input type="text" name="txtVendor" maxlength="50" class="blue-normal" size="20" style="width:50%" value="<%=strVendor%>"></td>
                              <td valign="top" class="blue-normal" >&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">No.Of Licences</td>
                              <td valign="middle">
								<input type="text" name="txtNoLicence" maxlength="20" class="blue-normal" size="20" style="width:30%" value="<%=intNumberOfLicence%>"></td>
                              <td valign="top" class="blue-normal" >&nbsp;</td>
                            </tr>
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" >
                              		<textarea rows="10" name="txtNote" style="width:100%" class="blue-normal"><%=strNote%></textarea>
                              		<!--<input type="text" name="txtNote" maxlength="500" class="blue-normal" style="width:95%" value="<%=strNote%>">--></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                                                                   
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="55">
                                      <div align="center" class="blue"><a href="javascript:savedata()"  class="b">Save</a></div>
                                    </td>
<%if cint(intRegisterID)>0 then%>                                    
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="55">
                                      <div align="center" class="blue"><a href="javascript:deletedata()"  class="b">Delete</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="100">
                                      <div align="center" class="blue"><a href="javascript:SetOfSoftware()"  class="b">Add Softwares</a></div>
                                    </td>
<%end if%>                                    
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="10%">No.</td>                 
                              <td class="blue" align="center" width="80%">Software name</td>
                              <td class="blue" align="center" width="10%"><input type='checkbox' name='chkAll' value='1' onclick='checkedAll(this);' ></td> 
                            </tr>
<%Response.Write strLast%>
                          </table>
<%if strLast<>"" then%>                          
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue" align="right"><a href="javascript:removeSoftware()">Remove</a>&nbsp;&nbsp;</td>
                            </tr>
                          </table>
                          


        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
			  <tr> 
			    <td align="right" bgcolor="#E7EBF5"> 
			      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
			        <tr class="black-normal"> 
			          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
			          </td>
			          <td align="center" valign="middle" width="13%" class="blue-normal"> 
			            <input type="text" name="txtpage" class="blue-normal" value="<%=intPage%>" size="2" style="width:50">
			          </td>
			          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();"  onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
			          </td>
			          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=intPage%>/<%=intPageCount%>&nbsp;&nbsp;</td>
			          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();"  
			          onMouseOver="self.status='Previous page'; return true;" onMouseOut="self.status='';">Previous</a> /
			          <a href="javascript:next();"  onMouseOver="self.status='Next page'; return true;" onMouseOut="self.status='';"> Next</a>&nbsp;&nbsp;&nbsp;</td>
			        </tr>
			      </table>
			    </td>
			  </tr>
			</table> 
<%end if%>			                         
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
%>

<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))    
%>
<input type="hidden" name="fgstatus" value="<%=fgDel%>">
<input type="hidden" name="txtID" value="<%=intRegisterID%>">
<input type="hidden" name="txtPCRegisterID" value="">
<input type="hidden" name="txtName" value="<%=strRegisterName%>">

</form>

</body>
</html>
