<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->  
<%
	Dim intUser, varUser, strConnect, strSQL, strTo,strToStaffID,strFrom, rsReportTo
    Dim objMail, objDatabase
    
'--------------------------------------------------
' 
'--------------------------------------------------    
function GetInformation(byval staffID, byval strField,byval rs)
	dim strReturn
	
	strReturn=""
	
	rs.Filter="StaffID=" & staffID
	if not rs.EOF then		
		strReturn=rs(strField)
	end if
	
	rs.Filter=""
	GetInformation=strReturn
end function
'--------------------------------------------------
' 
'--------------------------------------------------   
sub RecordEmailForStaff(intUserID,strContent)

	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	
	If objDatabase.dbConnect(strConnect) Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "RecordEmailNotification"
		
		Set myParam = myCmd.CreateParameter("SenderID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("EmailContent",adLongVarChar,adParamInput,10000)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("Recipient",adVarChar,adParamInput,2000)
		myCmd.Parameters.Append myParam
				
		myCmd("SenderID") = intUserID
		myCmd("EmailContent") = strContent
		myCmd("Recipient") = strToStaffID
		
		myCmd.Execute		
	end if
end sub
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID	= session("USERID")
	
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------

	varUser = session("varInvalidUser")
	intUser = Ubound(varUser,2)

	If intUser > 0 Then
		strTo = ""
		strToStaffID=""
		For ii = 1 To intUser
			If strTo = "" Then
				strTo = varUser(0,ii) 
				strToStaffID=varUser(1,ii)
			Else	
				strToStaffID=strToStaffID & "#" & varUser(1,ii)
				
				if right(strTo,1)<>">" then	
					strTo = strTo & ";" & varUser(0,ii)
				else
					strTo = strTo & varUser(0,ii)
				end if
			End If	
			if (ii mod 5 =0) then strTo = strTo & "<br>"
		Next
	End If
	
	strSQL = "SELECT personID as staffID,EmailAddress_ex as EmailAddress ,FirstName FROM ATC_PersonalInfo"
	call GetRecordset(strSQL,rsPersonInfo)
	strFrom=GetInformation(intUserID,"EmailAddress",rsPersonInfo)

	If Request.QueryString("act") <> "" Then
		
		strSQL = "SELECT StaffID,ISNULL(b.EmailAddress_ex,'') as EmailAddress FROM ATC_Employees a LEFT JOIN ATC_PersonalInfo b ON a.DirectLeaderID=b.PersonID"	
		call GetRecordset(strSQL,rsReportTo)	
		
			
	    	strFrom = Request.Form("email")
	    
		strSubject = Request.Form("txtsubject")
		strContent = Request.Form("txtcontent")
		
		varUser = session("varInvalidUser")

		For ii = 1 To Ubound(varUser,2)
						
			strTo1 = Trim(varUser(0,ii))

			
			Set cdoMessage = CreateObject("CDO.Message")  
			With cdoMessage 
				Set .Configuration = getCDOConfiguration()  
				.From = GetInformation(intUserID,"EmailAddress",rsPersonInfo)
				.To = strTo1 
				.Bcc="uyenchi.nguyentai@atlasindustries.com"
				.Subject = strSubject
				.TextBody = replace(strContent,"#Name#",GetInformation(varUser(1,ii),"FirstName",rsPersonInfo))
	
				.Send 
			End With

			Set cdoMessage = Nothing  
			Set cdoConfig = Nothing

		Next
		if Request.Form("chkRecord")<>"" then Call RecordEmailForStaff(intUserID,strContent)
		fgStatus = 1
	End If

	Set objDatabase = Nothing
%>

<html>
<head>
<title>Atlas Industries - Timesheet System</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/library.js"></script>
<script language="javascript">
<!--
	if ("<%=fgStatus%>" != "")
		window.close();
		
	function sendmail()
	{
		window.document.frmsend.action = "rpt_send_mail.asp?act=y"
		window.document.frmsend.submit();
	}
	
//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" language="javascript" onFocus="document.frmsend.txtcontent.focus()">
<form name="frmsend" method="post">
<table width="90%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="blue-normal" width="10%" height="27" align="right">From:&nbsp;</td>
    <td height="27" width="90%" class="blue">&nbsp;<%=strFrom%></td>
  </tr>
  <tr> 
    <td class="blue-normal" width="10%" height="27" align="right">To:&nbsp;</td>
    <td height="27" width="90%" class="blue">&nbsp;<%=strTo%></td>
  </tr>
  <tr> 
    <td class="blue-normal" width="10%" height="27" align="right">Subject:&nbsp;</td>
    <td height="27" width="90%"> 
      <input type="text" name="txtsubject" class="blue-normal" size="58" tabindex="1" value="Invalid Timesheet" style="width:100%">
    </td>
  </tr>
  <tr height=100%> 
    <td class="blue-normal" width="10%" align="right" valign="top">Message:&nbsp;</td>
    <td class="blue-normal" width="90%"> 
      <textarea name="txtcontent" class="blue-normal" rows="10" tabindex="2" style="width:100%;height=100%"></textarea>
    </td>
  </tr>
  <tr> 
    <td class="blue-normal" width="20%">&nbsp; </td>
    <td>
		<table width="100%">
			<tr>
				<td class="blue-normal"><input type="checkbox" name="chkRecord" value="1" checked>Recorded</td>
				<td class="blue-normal" align="right"><b>#Name#</b> will be replace by Firstname of recipient</td>
			</tr>
		</table>
	</td>
  </tr>
  <tr> 
    <td class="blue-normal" colspan="2"> 
      <table width="60" border="0" cellspacing="2" cellpadding="0" height="20" align="center">
        <tr> 
          <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" width="59" height="20" > 
            <div align="center" class="blue"><a href="javascript:sendmail();" class="b" onMouseOver="self.status='Clich here to send mail';return true" onMouseOut="self.status='';return true">Send</a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<input type="hidden" name="email" value="<%=strFrom%>">
</form>
</body>
</html>