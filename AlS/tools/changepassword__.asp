<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file="../class/clsSHA-1.asp" -->
<!-- #include file = "../inc/library.asp"-->
<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	Dim gMessage
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'----------------------------------
' Get Full Name and Job Title
'----------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
'----------------------------------	
' Make list of menu
'----------------------------------
	If isEmpty(session("Menu")) then 
		getRes = getarrMenu(session("USERID"))
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End if	
	
	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "C"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../")
'-----------------------------------
' Analyse query string
'-----------------------------------

if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
end if
strAct = Request.QueryString("act")
gMessage=""
if fgExecute then
	if strAct = "SAVE" then
		strOld = Request.Form("txtold")
		strNew = Request.Form("txtnew")
		strCon = Request.Form("txtconfirm")
		Set objSHA1 = New clsSHA1	
		strDigest = ObjSHA1.SecureHash(strOld)
		strConnect = Application("g_strConnect") 
		Set objDb = New clsDatabase
		If objDb.dbConnect(strConnect) then
		  strQuery = "Select count(*) as mysum From ATC_Users Where UserID = " & session("USERID") & "and Password = '" & strDigest & "'"
		  ret = objDb.runQuery(strQuery)
		  if ret then
			if objDb.rsElement("mysum")=1 then '--------------starting update
				objDb.cnDatabase.BeginTrans
				strDigest = ObjSHA1.SecureHash(strNew)
				strQuery = "UPDATE ATC_Users SET Password = '" & strDigest & "' WHERE UserID = " & session("USERID")
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
	Elseif strAct = "LIST" then
	    
	    Set objSHA1 = New clsSHA1	
		strDigest = ObjSHA1.SecureHash(strOld)
		strConnect = Application("g_strConnect") 
		Set objDb = New clsDatabase
		
		If objDb.dbConnect(strConnect) then
		  strQuery = "SELECT * FROM ATC_Users WHERE NewPassword IS NOT NULL"
		  ret = objDb.runQuery(strQuery)
		  if ret then
	        if not objDb.noRecord then
	        
	            strDigest = ObjSHA1.SecureHash(strNew)
		        arrlistUsers = objDb.rsElement.GetRows
		        
		        strDigest = ObjSHA1.SecureHash(strOld)
		        
		        for i=0 to UBound(arrlistUsers,2)
		            strDigest = ObjSHA1.SecureHash(arrlistUsers(7,i))
		            
		            strSql="UPDATE ATC_Users SET EncrypNewPassword = '" & strDigest & "' WHERE UserID = " & arrlistUsers(0,i)
		            Response.Write strSql & "<br>"
		            
		        next
		            
		        objDb.CloseRec
	        else
		        arrlistUsers = ""
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
        if (document.frmdetail.txtnew.value == "") {
            alert("Please enter your new password.");
            document.frmdetail.txtnew.focus();
            return false;
        }
        if (document.frmdetail.txtconfirm.value == "") {
            alert("Please re-enter your new password.");
            document.frmdetail.txtconfirm.focus();
            return false;
        }
        var strtmp1 = document.frmdetail.txtnew.value;
        var strtmp2 = document.frmdetail.txtconfirm.value;
        if ((strtmp1 != "") && (strtmp2 != "") && (strtmp1 != strtmp2)) {
            alert("New Password and Confirmation are not consistent!");
            document.frmdetail.txtconfirm.value = "";
            document.frmdetail.txtconfirm.focus();
            return false;
        }
        return true;
    }

    function save() {
        if (checkdata() == true) {
            document.frmdetail.action = "changepassword.asp?act=SAVE";
            document.frmdetail.target = "_self";
            document.frmdetail.submit();
        }
    }

    function List() {

        document.frmdetail.action = "changepassword.asp?act=LIST";
        document.frmdetail.target = "_self";
        document.frmdetail.submit();

    }
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmdetail" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>

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
                <div id="container"> <!--Begin Container-->
                <!--<form>
                    <ul>
                        <li>
                            <label>Old Password</label>
                            <span><input id="txtOld" name="txtOld" type="text" /> </span>
                        </li>
                    </ul>
                </form>-->
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
                        <td valign="middle" class="blue-normal" width="22%">Re-enter New Password</td>
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
                          <table border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                            <tr> 
                              <td width="60"  bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                                  <a href="javascript:save();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Change</a> 
                              </td>
                              <%if session("USERID")=252 then %>
                              <td width="60"  bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                                  <a href="javascript:List();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">List</a> 
                              </td>
                              <%end if %>
                            </tr>
                          </table>
                          
                          </div> <!--End Container-->
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
</form>
</body>
</html>