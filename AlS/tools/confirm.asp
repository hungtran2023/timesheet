<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file="../class/clsSHA-1.asp" -->
<!-- #include file="../inc/cdovbs.inc" -->
<%
'-----------------------------------
' Analyse query string
'-----------------------------------
intStaffID= Request.QueryString("id")


strAct = Request.QueryString("act")
gMessage=""
fgSucc = false

strConnect = Application("g_strConnect") 
Set objDb = New clsDatabase

if intStaffID<>"" then
    If objDb.dbConnect(strConnect) then
	  strQuery = "SELECT ISNULL(DATEDIFF(day,DateRequest,getdate()) ,-1) as checkrequest,fgChangePass FROM dbo.ATC_Users WHERE UserID=" & intStaffID
	  ret = objDb.runQuery(strQuery)	  
	  if ret then
	    if cint(objDb.rsElement("checkrequest"))<>0 And objDb.rsElement("fgChangePass") =1 then
	        gMessage="The link you are trying to access is no longer available"
	    end if
	  end if
    else
        gMessage = objDb.strMessage
	End if
end if

if strAct = "SAVE" then
	strpass = Request.Form("txtNew")
	intStaffID=Request.Form("txtstaff")


	Set objSHA1 = New clsSHA1
	strDigest = ObjSHA1.SecureHash(strpass)
	Set ObjSHA1 = Nothing
	
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
        
        objDb.cnDatabase.BeginTrans
		strQuery = "UPDATE ATC_Users SET Password = '" & strDigest & "', DateRequest = NULL, fgChangePass=1 WHERE UserID = " & intStaffID 
	
		if objDb.runActionQuery(strQuery) then
			objDb.cnDatabase.CommitTrans
		else
			gMessage = objDb.strMessage
	  		objDb.cnDatabase.RollbackTrans
		end if

		if gMessage = "" then
			fgSucc = true
			gMessage = "Your password has been reset. <br>" 
		end if
	end if

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
<link href="../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>
<link href="../jQuery/atlasJquery.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../jQuery/jquery-ui.min.js"></script>

<script language="javascript" src="../library/library.js"></script>

<script type="text/javascript">

    $(document).ready(function() {

        //$("#pswd_info").hide();


        $("#txtNew").focus();
        $("#submit").click(function(e) {

           

            //var errMsg = "<b>Password must meet the following requirements:</b><ul>";
            var errMsg = ""
            var newPass = $("input#txtNew").val();
            var retypePass = $("input#txtRetype").val();
            

            if (newPass.length < 8)
                errMsg = errMsg + "<li>Be at least <strong>8 characters</strong></li>";
                
            if (newPass.match(/[a-zA-Z]/) == null)
                errMsg = errMsg + "<li>At least <strong>one English uppercase or lowercase </strong></li>";
                
            if (newPass.match(/\d/) == null)
                errMsg = errMsg + "<li>At least <strong>one number</strong></li>";

            if (newPass.match(/\W/) == null)
                errMsg = errMsg + "<li>At least <strong>one nonalphanumeric character such as [@#$%^&*...]</strong></li>";
                
            if (errMsg != "")
                errMsg = " <h5>Password must meet the following requirements:</h5><ul>" + errMsg + "</ul>";

            if ((errMsg == "") && (retypePass != newPass))
                errMsg = errMsg + "Password does not match the confirm password.";

            if (errMsg == "")
                save();
            else
                $("div#errMsg").html(errMsg);

        })

    });  
    
</script>



<script LANGUAGE="JavaScript">

    function save() {
            document.frmdetail.action = "confirm.asp?act=SAVE";
            document.frmdetail.target = "_self";
            document.frmdetail.submit();
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
<%if gMessage<>"" then  %>          
<table width="332" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" height="117" bgcolor="#003399">
              <tr> 
                <td> 
                  <table width="330" border="0" cellspacing="0" cellpadding="0" align="center" height="0">
                    <form name="frmpass" method="post">
                      <tr bgcolor="8CA0D1"> 
                        <td colspan="2" class="blue" height="25"> <%if fgSucc then %>Password Reset<%else %>Access Denied<%end if %></td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td height="20">&nbsp;</td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td align="center" height="30" bgcolor="C0CAE6" class="blue-normal"> 
                          <%=gMessage %>
                        </td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td height="40"> 
                          <table width="100" border="0" cellspacing="0" cellpadding="0" align="center" height="20">
                            <tr> 
                              <td bgcolor="8CA0D1" align="center" class="blue" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" width="100%" height="20">
                                <a href="../initial.asp" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Back to Login</a>
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
<%else %>            
            <table width="352" border="0" cellspacing="0" cellpadding="0" align="center" height="152" bordercolor="#003399" bgcolor="#003399">
              <tr> 
                <td> 
                  <table width="350" border="0" cellspacing="0" cellpadding="0" align="center" height="150">
                    <form name="frmdetail" method="post">
                      <tr bgcolor="8CA0D1"> 
                        <td colspan="2" class="blue" height="25"> &nbsp;Reset Password</td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td colspan="2" height="20">&nbsp;</td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td valign="middle" class="blue-normal" width="42%">&nbsp;&nbsp;&nbsp;&nbsp;New Password</td>
                        <td valign="middle" width="58%" class="blue-normal"> 
                          <input type="password" name="txtnew" id="txtNew" class="blue-normal" size="25" style="width:180">
                        </td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td valign="middle" class="blue-normal" >&nbsp;&nbsp;&nbsp;&nbsp;Re-enter New Password</td>
                        <td valign="middle"  class="blue-normal"> 
                          <input type="password" id="txtRetype" name="txtRetype" class="blue-normal" size="25" style="width:180">
                        </td>
                      </tr>
                      <tr bgcolor="C0CAE6"> 
                        <td colspan="2" height="40"> 
                          <table border='0' cellspacing='5' cellpadding='0' align='center' height='20' name='aa'>
                            <tr>
                                <td bgcolor='#8CA0D1' width='60' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;' height='20' valign='middle'>
                                    <a class='b' id="submit" href='#'>Submit</a></td>
                                <td bgcolor='#8CA0D1' width='100' align='center' class='blue' onMouseOver='this.style.backgroundColor=&quot;#7791D1&quot;;' onMouseOut='this.style.backgroundColor=&quot;#8CA0D1&quot;;'height='20' valign='middle'>
                                    <a class='b' href='../initial.asp' onMouseOver="self.status='Login'; return true;" onMouseOut="self.status=''">Back to login</a></td>
                             </tr>
                           </table>
                        </td>
                      </tr>
                      <input type="hidden" name="txtstaff" value="<%=intStaffID%>">			
                    </form>
                  </table>
                </td>
              </tr>
            </table>
            <p></p>
          
            <div id="errMsg" class="red"><%=gMessage%></div> 
<% End if%>
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


</body>
</html>