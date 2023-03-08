<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file="../class/clsSHA-1.asp" -->
<!-- #include file = "../inc/library.asp"-->
<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
--> 

<%

'****************************************************************
'Get CDO Configuratio
'****************************************************************
Function getCDOConfiguration365()
    dim cdoConfig 
    Set cdoConfig = CreateObject("CDO.Configuration")  
    With cdoConfig.Fields  
        .Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="10.179.120.9"
        .Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 0
        .Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPconnectiontimeout

        ' Google apps mail servers require outgoing authentication. Use a valid email address and password registered with Google Apps.
        .Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="no-reply@atlasindustries.com"
        .Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="time7*sheed"

        .Update   
    
    End With  

	
	'Const SMTPsendusing=2
	'Const SMTPserver ="smtp.office365.com"
	'Const SMTPserverport = 587
	'Const SMTPusessl = 1
	'Const SMTPconnectiontimeout = 60
	'Const SMTPauthenticate = 1
	'Const SMTPsendusername ="atlas.ais.noreply@gmail.com"
	'Const SMTPsendpassword ="time7*sheed"

    set getCDOConfiguration365=cdoConfig
    
End Function

'-----------------------------------
' Analyse query string
'-----------------------------------
strAct = Request.QueryString("act")
gMessage=""
if strAct = "SAVE" then
	strusername = Request.Form("txtusername")
	'stremail = Request.Form("txtemail")
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "Select a.UserID , EmailID + '@atlasindustries.com' as EmailAddress, FirstName FROM ATC_Users a INNER JOIN ATC_PersonalInfo b ON a.UserID=b.PersonID INNER JOIN ATC_Employees c ON a.UserID=c.StaffID " &_
				"Where a.UserName = '" & strusername & "'"
	
	  ret = objDb.runQuery(strQuery)
	  if ret then
		if not objDb.noRecord then
		  if recCount(objDb.rsElement) = 1 then
			strName = objDb.rsElement("FirstName")
			stremail = objDb.rsElement("EmailAddress")
			intStaffID=objDb.rsElement("UserID")
' Call procedure that will send retrieved password to user
			
			
			objDb.cnDatabase.BeginTrans
			'strQuery = "UPDATE ATC_Users SET IDPassword = '" & strDigest & "' WHERE UserName = '" & strusername & "'"
			strQuery = "UPDATE ATC_Users SET DateRequest = Getdate(),IPRequest='" & Request.ServerVariables("REMOTE_ADDR") & "' WHERE UserName = '" & strusername & "'"
			
			if objDb.runActionQuery(strQuery) then
				objDb.cnDatabase.CommitTrans
			else
				gMessage = objDb.strMessage
	  			objDb.cnDatabase.RollbackTrans
			end if

			if gMessage = "" then
				If IsEmpty(Session("strHTTP")) Then
					Call MakeHTTP
				End if			
				strSubject = "Reset AIS Password"
				strTextBody="<p>Dear " & strName & ",</p> " & _
				                "<p>A request has been made to reset your password. For security reasons, Atlas Information System does not store and is not able to retrieve your existing password. Click the link below to set a new password which will allow you to log in.</p>" & _
				                "<p><a href='" & Session("strHTTP") & "tools/confirm.asp?id=" & intStaffID & "'>Reset your password now </a></p> " & _						
						        "<p>Trouble with the above link? You can copy and paste the following URL into your web browser: <b><i>" & _
						        Session("strHTTP") & "tools/confirm.asp?id=" & intStaffID & "</i></b></p>" &_
						
						        "<p>If you encounter any further problems, please contact your IT department. </p>" & _
						        "Regards,"
			
				Set cdoMessage = CreateObject("CDO.Message")  
		        With cdoMessage 
			        Set .Configuration = getCDOConfiguration()
			        .From = stremail
			        .To = stremail
			        .Subject = strSubject
			        .HTMLBody  = strTextBody

			        .Send 

		        End With

		        Set cdoMessage = Nothing  



				gMessage = "An email with password reset instructions has been sent to your address.<br> " & _
                            "Please click on the link in the email to reset your password. <br><br>"
				Session("strHTTP") = empty
			end if
		  else
		    gMessage = "The username you entered could not be found or email is invalid.<br><br>"
		  end if
		else
		  gMessage = "The username you entered could not be found or email is invalid.<br><br>"
		end if
	  else
	    gMessage = objDb.strMessage
	  end if
	  objDb.dbdisConnect
	Else
	  gMessage = objDb.strMessage
	End if
	set objDb = nothing
end if
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_menu.htm")


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
<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script LANGUAGE="JavaScript">
function checkdata() {
	document.frmpass.txtusername.value = alltrim(document.frmpass.txtusername.value);
	if (document.frmpass.txtusername.value=="") {
		alert("Please enter your UserName.");
		document.frmpass.txtusername.focus();
		return false;
	}
	
	return true;
}

function frm_submit() {
	if(checkdata()==true) {
		document.frmpass.action = "retrievepassword.asp?act=SAVE";
		document.frmpass.target = "_self";
		document.frmpass.submit();
	}
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6"  height="100%">&nbsp;</td>
    <td valign="middle" height="100%">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td width="7" background="images/l-03-3b.gif">&nbsp;</td>
          <td width="100%" valign="middle" height="100%" align="center"> 
 <%if gMessage<>"" then%>
            <table width="400" style="border-style:solid; border-width:1; border-color:#003399;" cellspacing="0" cellpadding="0"> 
                  <tr bgcolor="8CA0D1"> 
                       <td class="blue" height="30"> 
                                &nbsp;Reset Password Request Sent</td>
                       </tr>
                       
                          <tr bgcolor="C0CAE6"> 
                            <td class="blue-normal" style="text-align:center"> 
                              <p>&nbsp;</p><%=gMessage%>
                            </td>
                          </tr>
                          <tr bgcolor="C0CAE6"> 
                            <td height="40" > 
                              <table border='0' cellspacing='5' cellpadding='0' align='center' height='20' name='aa'>
                                <tr>
                                    <td bgcolor='#8CA0D1' width='100' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;'height='20' valign='middle'>
                                        <a class='b' href='../initial.asp';' onMouseOver="self.status='Login'; return true;" onMouseOut="self.status=''">Back to login</a></td>
                                 </tr>
                               </table>
                            </td>
                          </tr>
                  
              </table>
			        
<%else%>
            <table width="302" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" height="122" bgcolor="#003399">
              <tr> 
                <td> 
                  <table width="300" border="0" cellspacing="0" cellpadding="0" align="center" height="0">
                    <form name="frmpass" method="post">
                      <tr bgcolor="8CA0D1"> 
                        <td colspan="2" class="blue" height="30"> 
                            &nbsp;Reset Password</td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td colspan="2" height="20"></td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td width="30%" class="blue-normal" height="30" align="right"> 
                          User Name&nbsp;
                        </td>
                        <td width="70%" height="30" bgcolor="C0CAE6"> 
                          <input name="txtusername" maxlength="20" class="blue-normal" size="17" height="18px" value="<%=showvalue(strusername)%>" style="width:180px;height=21px; background-color: #ffffff; border-style :1px; border: thin #8CA0D1 solid">
                        </td>
                      </tr>
                      
                      <tr bgcolor="C0CAE6"> 
                        <td height="40" colspan="2"> 
                         
                          <table border='0' cellspacing='5' cellpadding='0' align='center' height='20' name='aa'>
                            <tr>
                                <td bgcolor='#8CA0D1' width='60' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;' height='20' valign='middle'>
                                    <a class='b' href="javascript:frm_submit();" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Submit</a></td>
                                <td bgcolor='#8CA0D1' width='100' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;'height='20' valign='middle'>
                                    <a class='b' href='../initial.asp';' onMouseOver="self.status='Login'; return true;" onMouseOut="self.status=''">Back to login</a></td>
                             </tr>
                           </table>
                        </td>
                      </tr>
                    </form>
                  </table>
                </td>
              </tr>
            </table>
            <p></p>
           <span class="blue">Enter your user name and we will send you a link 
                    <br>which will allow you to reset your password.</span>
<%end if%>
          </td>
          <td width="3" background="images/l-03-2b.gif">&nbsp;</td>
        </tr>
      </table>  
    </td>
    <td width="2"  height="100%">&nbsp;</td>
  </tr>
</table>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2)) %>
</form>
</body>
</html>