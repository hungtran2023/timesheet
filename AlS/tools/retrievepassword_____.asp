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
'-----------------------------------
' Analyse query string
'-----------------------------------
strAct = Request.QueryString("act")
gMessage=""
if strAct = "SAVE" then
	strusername = Request.Form("txtusername")
	stremail = Request.Form("txtemail")
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "Select b.EmailAddress, (b.FirstName + ' ' + ISNULL(b.MiddleName,'') + ' ' + b.LastName) AS FullName From ATC_Users a INNER JOIN ATC_PersonalInfo b ON a.UserID = b. PersonID " &_
				"Where a.UserName = '" & strusername & "'"
	  ret = objDb.runQuery(strQuery)
	  if ret then
		if not objDb.noRecord then
		  if recCount(objDb.rsElement) = 1 then
			strName = objDb.rsElement("FullName") 
' Call procedure that will send retrieved password to user
			Randomize
			str = Chr(Int((24 * Rnd) + 1)+96) & Chr(Int((24 * Rnd) + 1)+96) & Chr(Int((24 * Rnd) + 1)+96) &_
					Chr(Int((24 * Rnd) + 1)+96) & Chr(Int((24 * Rnd) + 1)+96) & Chr(Int((24 * Rnd) + 1)+96)
			Set objSHA1 = New clsSHA1
			strDigest = ObjSHA1.SecureHash(str)
			Set ObjSHA1 = Nothing
			
			objDb.cnDatabase.BeginTrans
			'strQuery = "UPDATE ATC_Users SET IDPassword = '" & strDigest & "' WHERE UserName = '" & strusername & "'"
			strQuery = "UPDATE ATC_Users SET IDPassword = '" & strDigest & "',IPRequest='" & Request.ServerVariables("REMOTE_ADDR") & "' WHERE UserName = '" & strusername & "'"
		
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
				
				strSubject = "Retrieved password"
				strTextBody = "Dear " & strName & ", " & chr(13) & chr(13) & _
						"A request has been made to reset your password." & chr(13)& " For security reasons, AIS does not store and is not able to retrieve your existing password." & chr(13) & "Click the link below to set a new password which will allow you to log in." & chr(13) & chr(13) & _
						"Click on the following URL " & Session("strHTTP") & "tools/confirm.asp and input the ID Password. " & chr(13) & chr(13) & _
						"ID Password: " & str & chr(13) & chr(13) & _
						"If you encounter any further problems, please contact your IT Department. " & chr(13) & chr(13) & _
						"Regards, "
						
				Set cdoConfig = CreateObject("CDO.Configuration")  
		        With cdoConfig.Fields  
			        .Item(cdoSendUsingMethod) = cdoSendUsingPort  
			        .Item(cdoSMTPServer) = AtlasSMTPServer  
			        .Update  
		        End With

		        Set cdoMessage = CreateObject("CDO.Message")  
		        With cdoMessage 
			        Set .Configuration = cdoConfig 
			        .From = stremail
			        .To = stremail 
			        .Subject = strSubject
			        .TextBody = strTextBody
			        '.Send 
		        End With

		        Set cdoMessage = Nothing  
		        Set cdoConfig = Nothing
		        
				gMessage = "An email with password has been sent to your address.<br>" & _
                            "Please click on the link in the email to confirm your new password.<br>"
				
				
				Session("strHTTP") = empty
			end if
		  else
		    gMessage = "Invalid Email address."
		  end if
		else
		  gMessage = "Invalid Email address."
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
    if (document.frmpass.txtusername.value == "") {
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
<table width="780" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6" background="../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
    <td valign="middle" height="100%">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td width="7" background="images/l-03-3b.gif">&nbsp;</td>
          <td width="100%" valign="middle" height="100%" align="center"> 
            <table width="302" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" height="122" bgcolor="#003399">
              <tr> 
                <td> 
                  <table width="300" border="0" cellspacing="0" cellpadding="0" align="center" height="0">
                    <form name="frmpass" method="post">
                      <tr bgcolor="8CA0D1"> 
                        <td colspan="2" class="blue" height="30"> &nbsp;Retrieve Password</td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td colspan="2" height="20"></td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td width="30%" class="blue-normal" height="30" align="right"> 
                          User Name&nbsp;
                        </td>
                        <td width="70%" height="30" bgcolor="C0CAE6"> 
                          <input name="txtusername" maxlength="20" class="blue-normal" size="17" height="18px" value="<%=showvalue(strusername)%>" style="width:130px;height=21px; background-color: #ffffff; border-style :1px; border: thin #8CA0D1 solid">
                        </td>
                      </tr>
                      
                      <tr bgcolor="C0CAE6"> 
                        <td height="40" colspan="2"> 
                          <table width="62" border="0" cellspacing="0" cellpadding="0" align="center" height="20">
                            <tr> 
                              <td bgcolor="8CA0D1" align="center" class="blue" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" width="62" height="20">
                                <a href="javascript:frm_submit();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Submit</a>
                                <a>Back to Login</a>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </form>
                  </table>
                </td>
              </tr>
            </table>
            <br>
            <span class="blue">Enter your user name and we will send you a link<br /> which will allow you to reset your password.  </span>
<%if gMessage<>"" then%>
			<br><br><span class="red"><b><%=gMessage%></b></span>
<%end if%>
          </td>
          <td width="3" background="images/l-03-2b.gif">&nbsp;</td>
        </tr>
      </table>  
    </td>
    <td width="2" background="../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
  </tr>
</table>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>
</body>
</html>