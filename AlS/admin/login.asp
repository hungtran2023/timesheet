<%@ Language=VBScript %>
<!-- #include file = "../class/CDatabase.asp"-->
<!-- #include file = "../class/clsSHA-1.asp" -->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->

<%
Response.Expires = - 1441
Response.Buffer = true

If Request.Form("txtusername") <> "" Then
  strUserName = Request.Form("txtusername")
  If trim(strUserName) = "administrator" then
   	Dim objDatabase, objSHA1
	Dim strConnect, strDigest, strPassword, strError
	
	strConnect = Application("g_strConnect")
			
' Connect to SQL database 
	Set objDatabase = New clsDatabase 

	If objDatabase.dbConnect(strConnect) Then
		strPassword = Request.Form("txtpwd")
   
		Set objSHA1 = New clsSHA1	
		strDigest = ObjSHA1.SecureHash(strPassword)
'		Response.Write strDigest
'		Response.End
		Set ObjSHA1 = Nothing
		
		If (objDatabase.runQuery("SELECT count(*) FROM ATC_Admin WHERE Password = '" & strDigest & "'")) Then
			If objDatabase.rsElement(0)=1 Then
				'get companyID
				session("InHouse") = 0
				If (objDatabase.runQuery("SELECT CompanyID FROM ATC_CompanyProfile")) Then
					If not objDatabase.noRecord then
						session("InHouse") = objDatabase.getColumn_by_name("CompanyID")
					end if
				End if
				objDatabase.dbDisConnect
				Set objDatabase = Nothing

				Call MakeHTTP

				Response.Clear
				Response.Redirect("profile.asp")
			Else
				strError = "Invalid password!"
			End If
		Else
			strError = objDatabase.strMessage
		End If	
	Else
		strError = objDatabase.strMessage
		objDatabase.dbDisConnect
		Set objDatabase = Nothing
	End If
  Else
	strError = "Invalid Username"
  End if
End If

%>

<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../timesheet.css" type="text/css">
<script language="javascript" src="../library/library.js"></script>
<script language="Javascript">
function checkin() {
	if (isempty(window.document.frmlogin.txtusername.value)) 
	{		
		alert("Please enter user name.");
		window.document.frmlogin.txtusername.focus();
	}
	else 
	{
		window.document.frmlogin.action ="login.asp"; 
		window.document.frmlogin.target = "_self";
		window.document.frmlogin.submit();
	}
}

function window_onload()
{
var strError = "<%=strError%>";
	if (strError == "Invalid password!")
		window.document.frmlogin.txtpwd.focus();
	else{
		window.document.frmlogin.txtusername.focus();}
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" language="javascript" onload="window_onload()">
<form name="frmlogin" method="post" action="">
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="85" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="525"><img src="../images/title.gif" width="470" height="59"></td>
          <td width="255" class="red" align="right" valign="bottom">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
        <tr> 
          <td height="24" width="1" valign="top"><img src="../images/l-02.gif" width="1" height="24"></td>
          <td width="778" height="24" background="../images/l-01.gif" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="334"><img src="../images/line-l.gif" width="334" height="24"></td>
                <td width="437" align="right" class="blue">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td width="1" height="24" valign="top"><img src="../images/l-02.gif" width="1" height="24"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="10" background="../images/l-03-3a.gif"><img src="../images/l-03-3a.gif" width="1" height="7"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6" background="../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
    <td valign="top" height="100%" width="772">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="height:79%" height="365">
        <tr>
          <td bgcolor="#FFFFFF" valign="middle">&nbsp;</td>
        </tr>  
		<tr>
		  <td bgcolor="#FFFFFF" valign="middle">
			<table width="252" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" height="157" bgcolor="#003399">
			  <tr> 
			    <td> 
			      <table width="250" border="0" cellspacing="0" cellpadding="0" align="center" height="155">
                    <tr bgcolor="#C0CAE6"> 
                      <td colspan="2" height="20" class="red">&nbsp;</td>
                    </tr>
				    <tr bgcolor="#C0CAE6"> 
			          <td width="120" class="blue-normal" height="35"> 
						  <div align="right" class="c">User Name&nbsp; </div>
			          </td>
			          <td width="193" height="35" class="text-blue01" bgcolor="#C0CAE6"> 
			            <input id="txtusername" name="txtusername" tabindex="1" <%If strError = "Invalid password!" Then%> value="<%=Request.Form("txtusername")%>" <%End If%> size="17" height="18px" style="width:130px;height=21px; background-color: #ffffff; border-style :1px; border: thin #8CA0D1 solid">
			          </td>
			        </tr>
			        <tr bgcolor="#C0CAE6"> 
			          <td width="120" class="blue-normal" height="35"> 
			            <div align="right">Password&nbsp; </div>
			          </td>
			          <td width="193" height="35" class="text-blue01" bgcolor="#C0CAE6"> 
			            <input type="password" id="txtpwd" name="txtpwd" tabindex="2" size="17" height="18px" style="width:130px;height=21px; background-color: #ffffff; border-style :1px; border: thin #8CA0D1 solid">
			          </td>
			        </tr>
			        <tr bgcolor="#C0CAE6"> 
			          <td height="40" colspan="2"> 
			            <table width="60" border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa">
			              <tr> 
			                <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
			                  <div align="center" class="blue"><a href="javascript:checkin()" class="b" onMouseOver="self.status='Login to timesheet system.';return true" onMouseOut="self.status='';return true">Login</a></div>
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
		<tr>
		  <td>&nbsp;</td>
		</tr>  
        <tr>
		  <td>
		    <table width="250" border="0" cellspacing="0" cellpadding="0" align="center" height="0">
			  <tr>
				<td>
				  <div align="center"><font face="Arial" color="#FF0000"><b><%=strError%></b></font></div>
				</td>
			  </tr>
			</table>
		  </td>	  	
		</tr>
  	  </table>
    </td>  
    <td width="2" background="../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="../images/dot-01.gif" colspan="3"><img src="../images/dot-01.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td bgcolor="#8CA0D1" width="1"><img src="../images/l-02.gif" width="1" height="24"></td>
    <td bgcolor="#8CA0D1" class="blue-normal" width="778"> 
      <div align="center"> Copyright &copy; 2000-2006 Atlas Industries Limited. 
        All Rights Reserved </div>
    </td>
    <td bgcolor="#8CA0D1" width="1" align="right"><img src="../images/l-02.gif" width="1" height="24"></td>
  </tr>
  <tr> 
    <td background="../images/dot-01.gif" colspan="3"><img src="../images/dot-01.gif" width="1" height="1"></td>
  </tr>
</table>
</form>
<SCRIPT language=JavaScript1.2>
var hotkey=13
if (document.layers)
document.captureEvents(Event.KEYPRESS)
function backhome(e){
	if (document.layers){
		if (e.which==hotkey)
			checkin()
	}
	else if (document.all){
		if (event.keyCode==hotkey){
			event.keyCode = 0;
			checkin()
			}
		}
}
document.onkeypress=backhome
</SCRIPT>
</body>
</html>