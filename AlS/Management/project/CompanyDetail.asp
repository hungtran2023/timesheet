<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim intID,strClientName,strServerPath, strClientCode,strType
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon

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
	intID = Request.Form("txtID")
	fgDel=Request.Form("fgstatus")
	
	if Request.QueryString("act") = "save" then
		
		strClientName=Request.Form("txtClientName")
		strServerPath=Request.Form("txtPath")
		strWebsite=Request.Form("txtWebsite")
		strEmail=Request.Form("txtEmail")
		strNote=Request.Form("txtNote")
		strType=Request.Form("chkType")
		if strType="" then strType=1
		
		if fgDel<>"D" then
					
			strSql="UPDATE  ATC_Companies " & _
                    "SET  CompanyName  = '" & replace(strClientName,"'","''") & "'"& _
					", EmailAddress  = " & IIF(strEmail="", "NULL", "'" & replace(strEmail,"'","''") & "'") & _
                    ", Website  = " & IIF(strWebSite="", "NULL", "'" & replace(strWebsite,"'","''") & "'") & _
                    ", Note  = " & IIF(strNote="","NULL", "'" & Replace(strNote,"'","''") & "'") & _
                    ", SeverPath  =" & IIF(strServerPath="","NULL", "'" & Replace(strServerPath,"'","''") & "'") & _
                    ", [type]=" & strType & _
                " WHERE CompanyID=" & intID
				
		end if
'response.write 		strSql
	    strError= ExecuteSQL(strSql)
	
	End If
'--------------------------------------------------
' 
'--------------------------------------------------
	
	strSql="SELECT CompanyID,CompanyName,EmailAddress,CharCode,SeverPath,Note,Website,[Type]  FROM ATC_Companies WHERE CompanyID=" & intID

	Call GetRecordset(strSql,rsData)	

	if rsData.RecordCount>0 then
	
		strClientCode=rsData("CharCode")
		strClientName=rsData("CompanyName")
		strServerPath=rsData("SeverPath")
		strWebsite=rsData("Website")
		strEmail=rsData("EmailAddress")
		strNote=rsData("Note")
		strType=rsData("Type")

	end if


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
	If strChoseMenu = "" Then strChoseMenu = "AC"
	
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
var strURL="CompanyDetail.asp";

function savedata()
{
	if (checkdata())
	{
		window.document.frmreport.action = strURL + "?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = strURL + "?act=save"			
	window.document.frmreport.submit();
}

function checkdata()
{
	if (window.document.frmreport.txtClientName.value=="")
	{
		alert("Please enter Client name.");
		document.frmreport.txtClientName.focus();
		return false	
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
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="listofcompany.asp" onMouseOver="self.status='';return true">Client List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
             
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Client Information</td>
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
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%">Client Code </td>
                              <td valign="middle" width="35%" class="blue">
								<%=strClientCode%>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%">Client name *</td>
                              <td valign="middle" width="35%" class="blue">
								<input type="text" name="txtClientName" maxlength="100" class="blue-normal" style="width:95%" value="<%=strClientName %>"></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Website</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtWebsite" maxlength="50" class="blue-normal" style="width:95%" value="<%=strWebsite%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Email domain</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtEmail" maxlength="50" class="blue-normal" style="width:95%" value="<%=strEmail%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Path</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtPath" maxlength="50" class="blue-normal" style="width:95%" value="<%=strServerPath%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" >
                              		<textarea rows="2" name="txtNote" style="width:95%" class="blue-normal"><%=strNote%></textarea>
                              		<!--<input type="text" name="txtNote" maxlength="500" class="blue-normal" style="width:95%" value="<%=strNote%>">--></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">TP</td>
                              <td valign="middle" >
                              		<input type="checkbox" name="chkType" id="chkType" value="2" <%if cint(strType)=2 then%>checked<%end if%>>
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
<%if cint(intID)>0 then%>                                    
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="55">
                                      <div align="center" class="blue"><a href="#"  class="b">Delete</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="70">
                                      <div align="center" class="blue"><a href="#"  class="b">Contacts</a></div>
                                    </td>
<%end if%>                                    
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                 
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
<input type="hidden" name="txtID" value="<%=intID%>">
<input type="hidden" name="txtPCSoftwareID" value="">
<input type="hidden" name="txtName" value="<%=strSoftwareName%>">

</form>

</body>
</html>
