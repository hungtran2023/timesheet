<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim intPCSoftwareID,strSoftwareName,strComputerName,strPCUserName,strOriginalLicence,	strCurrentLicence,dateInstallDate,strNote
	Dim objEmployee, objDatabase, strError,rsData,strPre
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

	If Not checkSession(session("USERID")) Then Response.Redirect("../../message.htm")
	
	intUserID = session("USERID")

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	intPCSoftwareID = Request.Form("txtPCSoftwareID")
'Response.Write intPCSoftwareID

	fgDel=Request.Form("fgstatus")
	
	strPre=Request.QueryString("fr")

	if Request.QueryString("act") = "save" then

		strPre=Request.Form("txtPre")
	
		strOriginalLicence=Request.Form("txtOriginalLicence")	
		strCurrentLicence=Request.Form("txtCurrentLicence")	
			
		'dateInstallDate=Request.Form("")		
		strNote=Request.Form("txtNote")
		
		strSQl="UPDATE ATC_PCSoftware SET OriginalLicence = " & IIF(strOriginalLicence="","NULL","'" & strOriginalLicence & "'") & _
											",CurrentLicence = " & IIF(strCurrentLicence="","NULL","'" & strCurrentLicence & "'") & _
											",Note = " & IIF(strNote="","NULL","'" & strNote & "'") & _
				" WHERE PCSoftwareID =" &  intPCSoftwareID
				
		strError= ExecuteSQL(strSQL)
				
	End If

'--------------------------------------------------
' 
'--------------------------------------------------
	
	strSql= "SELECT b.SoftwareName, c.ComputerName, ISNULL(d.UserName, c.PublicName) AS UserName, a.OriginalLicence, a.CurrentLicence, a.InstallDate, a.Note " & _
				"FROM ATC_PCSoftware AS a INNER JOIN " & _
                      "ATC_Softwares AS b ON a.SoftwareID = b.SoftwareID INNER JOIN " & _
                      "ATC_AtlasPC AS c ON a.AtlasPCID = c.AtlasPCID LEFT OUTER JOIN " & _
                      "ATC_Users AS d ON c.UserID = d.UserID " & _
			"WHERE     a.PCSoftwareID =" &  intPCSoftwareID
			
	Call GetRecordset(strSql,rsData)
	
	if rsData.RecordCount>0 then
	
		strSoftwareName=rsData("SoftwareName")
		strComputerName=rsData("ComputerName")
		strPCUserName=rsData("UserName")
		
		strOriginalLicence=rsData("OriginalLicence")
		strCurrentLicence=rsData("CurrentLicence")
		
		dateInstallDate=rsData("InstallDate")
		
		strNote=rsData("Note")

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

	
function savedata()
{
	window.document.frmreport.action = "LicenceSoftware.asp?act=save"			
	window.document.frmreport.submit();
}

function canceldata()
{
	var strPre="<%=strpre%>"
	var strLink="SoftwareDetail.asp";
	if (strPre=="com")
	{
		strLink="AtlasComputer.asp";
	}
	
	window.document.frmreport.action = strLink
	window.document.frmreport.submit();
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
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="SoftwareList.asp" onMouseOver="self.status='';return true">Software List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
             
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2"><span class="red"><%=strSoftwareName%></span> licence of <span class="blue"><%=strComputerName%></span></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                                                                                   
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%">UserName</td>
                              <td valign="middle" width="35%" class="blue">
								<%=strPCUserName%></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>

                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Original Licence</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtOriginalLicence" maxlength="50" class="blue-normal" size="20" style="width:95%" value="<%=strOriginalLicence%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Current Licence</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtCurrentLicence" maxlength="50" class="blue-normal" size="20" style="width:95%" value="<%=strCurrentLicence%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" >
                              		<textarea rows="2" name="txtNote" style="width:95%" class="blue-normal"><%=strNote%></textarea>
                              </td>
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
                                    
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="55">
                                      <div align="center" class="blue"><a href="javascript:canceldata()"  class="b">Cancel</a></div>
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
<input type="hidden" name="txtID" value="<%=Request.Form("txtID")%>">
<input type="hidden" name="txtPCSoftwareID" value="<%=intPCSoftwareID%>">
<input type="hidden" name="txtAtlasPCID" value="<%=Request.Form("txtAtlasPCID")%>">
<input type="hidden" name="txtName" value="<%=Request.Form("txtName")%>">
<input type="hidden" name="txtPre" value="<%=strPre%>">

</form>

<% if request.QueryString("act") = "save" and strError<>"" then%>

	<%if strPre="sof" then%> 

<script language="javascript">
<!--
	window.document.frmreport.action = "SoftwareDetail.asp"			
	window.document.frmreport.submit();
-->
</script>

	<%else%>
	
	<script language="javascript">
<!--
	window.document.frmreport.action = "AtlasComputer.asp"			
	window.document.frmreport.submit();
-->
</script>

<%	end if
end if%>
</body>
</html>
