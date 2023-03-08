<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file="../class/clsSHA-1.asp" -->
<!-- #include file = "../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("Inhouse")) = False Then
		Response.Redirect("message.htm")
	End If					

strtmp1 = Replace(logoff, "XX", session("strHTTP") & "admin/")
strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
			"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"

strAct = Request.QueryString("act")
if stract="" then
	Call freeAdmininput
	Call freeRole
	Call freeRoleAss
	Call freelistRole
end if
gMessage=""

if strAct = "SAVE" then
	strOld = Request.Form("txtold")
	strNew = Request.Form("txtnew")
	strCon = Request.Form("txtconfirm")
	Set objSHA1 = New clsSHA1	
	strDigest = ObjSHA1.SecureHash(strOld)
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "Select count(*) as mysum From ATC_Admin Where Password = '" & strDigest & "'"
	  ret = objDb.runQuery(strQuery)
	  if ret then
		if objDb.rsElement("mysum")=1 then '--------------starting update
			objDb.cnDatabase.BeginTrans
			strDigestOld = strDigest
			strDigest = ObjSHA1.SecureHash(strNew)
			strQuery = "UPDATE ATC_Admin SET Password = '" & strDigest & "' WHERE Password = '" & strDigestOld & "'"
			ret = objDb.runActionQuery(strQuery)
			if ret=false then				
				objDb.cnDatabase.RollbackTrans
				gMessage = objDb.strMessage
			else
				objDb.cnDatabase.CommitTrans
				gMessage = "Changed successfully."
				objDb.closerec
			end if
		else
		  gMessage = "Invalid Old PassWord !"
		end if
	  else
	    gMessage = objDb.strMessage
	  end if
	Else
	  gMessage = objDb.strMessage
	End if
	objDb.dbdisConnect
	set objDb = nothing
	Set ObjSHA1 = Nothing
end if
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_admin.htm")
curpage = 8
If arrPageTemplate(1)<>"" then
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(0) = Replace(arrTmp(0),"@@function", strfunction)
	for i = 1 to NumOfAdminMenu
		if i <> curpage then
			arrTmp(0) = Replace(arrTmp(0),"@@markin"&cstr(i)&"@@", "")
			arrTmp(0) = Replace(arrTmp(0),"@@markout"&cstr(i)&"@@", "")
		else
			arrTmp(0) = Replace(arrTmp(0),"@@markin"&cstr(i)&"@@", "<font color='#CD0000'>")
			arrTmp(0) = Replace(arrTmp(0),"@@markout"&cstr(i)&"@@", "</font>")
		end if
	next
End if
%>	

<html>
<head>
<title>Atlas Industries - Time Sheet System</title>
<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script LANGUAGE="JavaScript">
function checkdata() {
	if (alltrim(document.frmdetail.txtnew.value)=="") {
		alert("Please enter your new password.");
		document.frmdetail.txtnew.focus();
		return false;
	}
	if (alltrim(document.frmdetail.txtconfirm.value)=="") {
		alert("Please re-enter your new password.");
		document.frmdetail.txtconfirm.focus();
		return false;
	}
var strtmp1 = alltrim(document.frmdetail.txtnew.value);
var strtmp2 = alltrim(document.frmdetail.txtconfirm.value);
	if ((strtmp1!="")&&(strtmp2!="")&&(strtmp1!=strtmp2)) {
		alert("New Password and Confirmation are not consistent!");
		document.frmdetail.txtconfirm.value="";
		document.frmdetail.txtconfirm.focus();
		return false;
	}				
	return true;
}


function save() {
	if(checkdata()==true) {
		document.frmdetail.action = "changepass.asp?act=SAVE";
		document.frmdetail.target = "_self";
		document.frmdetail.submit();
	}
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmdetail" method="post">
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			%>		
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                 <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
                  <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
                </tr>
                <tr align="center"> 
                  <td class="blue" align="left" width="23%"> &nbsp;&nbsp;</td>
                  <td class="blue" align="right" width="77%">&nbsp;</td>
                </tr>
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Change password</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td bgcolor="#617DC0"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr bgcolor="#FFFFFF"> 
                        <td valign="top" width="27%" class="blue">&nbsp;</td>
                        <td valign="middle" class="blue-normal" width="22%"> Old 
                          Password</td>
                        <td valign="middle" width="51%" class="blue-normal"> 
                          <input type="password" name="txtold" class="blue-normal" size="15" style="width:130">
                        </td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td valign="top" width="27%" class="blue">&nbsp;</td>
                        <td valign="middle" class="blue-normal" width="22%">New 
                          Password</td>
                        <td valign="middle" width="51%" class="blue-normal"> 
                          <input type="password" name="txtnew" class="blue-normal" size="15" style="width:130">
                        </td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td valign="top" width="27%" class="blue">&nbsp;</td>
                        <td valign="middle" class="blue-normal" width="22%">Re-enter 
                          New Password</td>
                        <td valign="middle" width="51%" class="blue-normal"> 
                          <input type="password" name="txtconfirm" class="blue-normal" size="15" style="width:130">
                        </td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td valign="top" width="27%" class="blue">&nbsp;</td>
                        <td valign="middle" class="blue-normal" width="22%">&nbsp;</td>
                        <td valign="middle" width="51%" class="blue-normal">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td valign="top" width="27%" class="blue">&nbsp;</td>
                        <td valign="middle" class="blue-normal" width="22%">&nbsp;</td>
                        <td valign="middle" width="51%" class="blue-normal">&nbsp;</td>
                      </tr>
                    </table>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                      <tr> 
                        <td height="50"> 
                          <table width="60" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                            <tr> 
                              <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                                  <a href="javascript:save();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Change</a> 
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
</form>
</body>
</html>