<!--#INCLUDE FILE="fileupload_class.inc"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
'****************************************
' function: makelistbox
' Description: 
' Parameters: 
'			  
' Return value: string have HTML format
' Author: 
' Date: 
' Note:
'****************************************
function makelistbox(Byval strName, ByRef rsSrc, ByVal intSel)
	strOut = "<select name='" & strName & "' class='blue-normal' style='HEIGHT: 22px; WIDTH: 160px'>"
	rsSrc.MoveFirst
	Do Until rsSrc.EOF
	  if rsSrc(0)=int(intSel) then strSel=" selected " else strSel="" end if
	  strOut = strOut & "<option value='" & rsSrc(0) & "'" & strSel & ">" & showlabel(rsSrc(1)) & "</option>"
	  rsSrc.MoveNext
	Loop
	strOut = strOut & "</select>"
	makelistbox = strOut
end function
'****************************************
' function: tasksave
' Description: 
' Parameters: 
'			  
' Return value: string have HTML format
' Author: 
' Date: 
' Note:
'****************************************
function tasksave()
	intCompanyID = Request.Form("txthidden")
	strName = trim(Request.Form("txtcompany"))
	strName = "'" & Replace(strName, "'", "''") & "'"
	strAddress = trim(Request.Form("txtaddress"))
	if strAddress = "" then
		strAddress = "NULL"
	else
		strAddress = "'" & Replace(strAddress, "'", "''") & "'"
	end if
	strCity = trim(Request.Form("txtcity"))
	if strCity = "" then
		strCity = "NULL"
	else
		strCity = "'" & Replace(strCity, "'", "''") & "'"
	end if
	
	strState = trim(Request.Form("txtstate"))
	if strState = "" then
		strState = "NULL"
	else
		strState = "'" & Replace(strState, "'", "''") & "'"
	end if
	
	intCountryID = Request.Form("lstCountry")
	
	strPostal = trim(Request.Form("txtpostal"))
	if strPostal = "" then
		strPostal = "NULL"
	else
		strPostal = "'" & Replace(strPostal, "'", "''") & "'"
	end if
	
	strPhone = trim(Request.Form("txtphone"))
	if strPhone = "" then
		strPhone = "NULL"
	else
		strPhone = "'" & Replace(strPhone, "'", "''") & "'"
	end if
	
	strFax = trim(Request.Form("txtfax"))
	if strFax = "" then
		strFax = "NULL"
	else
		strFax = "'" & Replace(strFax, "'", "''") & "'"
	end if
	
	strEmail = trim(Request.Form("txtemail"))
	if strEmail = "" then
		strEmail = "NULL"
	else
		strEmail = "'" & Replace(strEmail, "'", "''") & "'"
	end if
	
	strWeb = trim(Request.Form("txtweb"))
	if strWeb = "" then
		strWeb = "NULL"
	else
		strWeb = "'" & Replace(strWeb, "'", "''") & "'"
	end if
	
	gMessage = ""
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	ret = objDb.dbConnect(strConnect)
	If ret Then
		objDb.cnDatabase.BeginTrans
		if int(intCompanyID) = 0 then 'insert
		  strQuery = "INSERT INTO ATC_Companies(CountryID, CompanyName, Address, City, State, PostalCode, Phone, Fax, " &_
					"EmailAddress, Website) VALUES(" & intCountryID & "," & strName & "," & strAddress & "," & strCity & "," &_
					strState & "," & strPostal & "," & strPhone & "," & strFax & "," & strEmail & "," & strWeb & ")"
		  if not objDb.runActionQuery(strQuery) then
			gMessage = objDb.strMessage
		  else
			strQuery = "Select @@IDENTITY as myid"
			if objDb.runQuery(strQuery) then
			  if not objDb.noRecord then
				intCompanyID = objDb.rsElement("myid")
				strQuery = "INSERT INTO ATC_CompanyProfile(CompanyID) VALUES(" & intCompanyID & ")"
				if not objDb.runActionQuery(strQuery) then 
					gMessage = objDb.strMessage
				else
					session("Inhouse") = Cint(intCompanyID)
				end if
			  else
				gMessage = "Error in INSERT statement."
			  end if
			else
			  gMessage = objDb.strMessage
			end if
		  end if
		else 'update
		  strQuery = "UPDATE ATC_Companies SET CountryID=" & intCountryID & ", CompanyName=" & strName & ", Address=" &_
					strAddress & ", City=" & strCity & ", State=" & strState & ", PostalCode=" & strPostal & ", Phone=" &_
					strPhone & ", Fax=" & strFax & ", EmailAddress=" & strEmail & ", Website=" & strWeb & _
					" WHERE CompanyID = " & intCompanyID
		  if not objDb.runActionQuery(strQuery) then
			gMessage = objDb.strMessage
		  end if
		end if
		if gMessage = "" then 'successfull
			objDb.cnDatabase.CommitTrans
			fgComplete = true
			gMessage = "Updates successfully."
		else
			fgComplete = false
			objDb.cnDatabase.RollbackTrans
		end if
		objDb.dbDisconnect
	else
		gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
	tasksave = fgComplete
end function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("Inhouse")) = False Then
		Response.Redirect("message.htm")
	End If
'-------------------------------------------------
Dim gMessage
Call freeAdmininput
Call freeRole
Call freeRoleAss
Call freelistRole

If IsEmpty(Session("strHTTP")) Then
	Call MakeHTTP
End if
strtmp1 = Replace(logoff, "XX", session("strHTTP")&"admin/")
strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
			"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"

if Request.QueryString("act") = "SAVE" then
	ret = tasksave
end if

intCompanyID = 0
strName = ""
strAddress = ""
strCity = ""
strState = ""
intCountryID = 0
strPostal = ""
strPhone = ""
strFax = ""
strEmail = ""
strWeb = ""
			
strConnect = Application("g_strConnect")
Set objDb = New clsDatabase
ret = objDb.dbConnect(strConnect)
If ret Then
	If session("Inhouse")<>0 then
		strQuery = "select a.CompanyID, CompanyName, ISNULL(CountryID, 0) CountryID, ISNULL(Address, '') Address, ISNULL(City, '') City," &_
					"ISNULL(State, '') State, ISNULL(PostalCode, '') PostalCode, ISNULL(Phone, '') Phone, ISNULL(Fax, '') Fax," &_
					"ISNULL(EmailAddress, '') Email, ISNULL(Website, '') Website, ISNULL(logo, '') Logo " &_
					"FROM ATC_CompanyProfile a INNER JOIN ATC_Companies b ON a.CompanyID = b.CompanyID "
		if objDb.runQuery(strQuery) then
			If not objDb.noRecord then
				intCompanyID = objDb.rsElement("CompanyID")
				strName = objDb.rsElement("CompanyName")
				strAddress = objDb.rsElement("Address")
				strCity = objDb.rsElement("City")
				strState = objDb.rsElement("State")
				intCountryID = objDb.rsElement("CountryID")
				strPostal = objDb.rsElement("PostalCode")
				strPhone = objDb.rsElement("Phone")
				strFax = objDb.rsElement("Fax")
				strEmail = objDb.rsElement("Email")
				strWeb = objDb.rsElement("Website")
				strlogo = objDb.rsElement("Logo")
			end if
		else
			gMessage = objDb.strMessage
		end if
	End if

	'Get list of Country
	strQuery = "Select CountryID, CountryName from ATC_Countries"
	if objDb.runQuery(strQuery) then
		If not objDb.noRecord then
			strListCountry = makelistbox("lstCountry", objDb.rsElement, intCountryID)
			objDb.CloseRec
		else
			strListCountry = ""
		end if
	else
		gMessage = objDb.strMessage
	end if
	objDb.dbDisconnect
Else
  gMessage = objDb.strMessage
End if
Set objDb = Nothing
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_admin.htm")
curpage = 1
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
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css" type="text/css">
<script language="javascript" src="../library/library.js"></script>
<script>
var objNewWindow;
function _upload() { //v2.0
  window.status = "";
  strFeatures = "top="+(screen.height/2-78)+",left="+(screen.width/2-132)+",width=265,height=158,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no";
  if((objNewWindow) && (!objNewWindow.closed))
	objNewWindow.focus();	
  else {
	objNewWindow = window.open("upload.asp", "MyNewWindow", strFeatures);
  }
  window.status = "Opened a new browser window.";  
}

function window_onunload() {
	if((objNewWindow)&&(!objNewWindow.closed))
		objNewWindow.close();
}

function _reset(){
	for(i=0;i<document.frminput.length;i++){
		document.frminput.elements[i].value = "" ;
	}
}

function checkdata() {
	var tmp = document.frminput.txtemail.value;
	document.frminput.txtemail.value = alltrim(tmp);
	if(document.frminput.txtemail.value!="") {
		if(isemail(document.frminput.txtemail.value)==false) {
			alert("Invalid value email address \nValid format is: 'NickName@domain.com'");
			document.frminput.txtemail.focus();
			return false
		}
	}
	tmp = document.frminput.txtweb.value;
	document.frminput.txtweb.value = alltrim(tmp);
	if(document.frminput.txtweb.value!="") {
		if(iswebsite(document.frminput.txtweb.value)==false) {
			alert("Invalid website value")
			document.frminput.txtweb.focus();
			return false
		}
	}
	tmp = document.frminput.txtcompany.value;
	document.frminput.txtcompany.value = alltrim(tmp);
	if(document.frminput.txtcompany.value=="") {
		alert("Please enter a value")
		document.frminput.txtcompany.focus();
		return false
	}

	return true;
}

function _save(){
	if(checkdata()==true) {
  		document.frminput.action = "profile.asp?act=SAVE"
		document.frminput.target = "_self";
		document.frminput.submit();
	}
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frminput" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			%>
  
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" LANGUAGE="javascript" onUnload="return window_onunload();">
  <tr> 
    <td align="center"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
			<td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
<%if session("Inhouse")<>0 then%>
        <tr> 
          <td class="blue" align="right" height="30">&nbsp;
			<%if strlogo<>"" then%><a href="javascript:_upload();" onMouseOver="self.status='Update logo'; return true;" onMouseOut="self.status='';"><img src="../images/<%=strlogo%>" border="0"></a>&nbsp;&nbsp;
			<%else%><a href="javascript:_upload();" onMouseOver="self.status='Update logo'; return true;" onMouseOut="self.status='';">Logo</a>
			<%end if%>&nbsp;&nbsp;</td>
        </tr>
<%else%>
        <tr> 
          <td class="blue" align="right" height="30">&nbsp;</td>
        </tr>
<%end if%>
        <tr> 
          <td class="title" height="50" align="center"> Company Profile 
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr> 
                <td class="blue" align="center" width="150">&nbsp;</td>
                <td class="blue" align="left" colspan="2">&nbsp;</td>
                <td class="blue" align="center" width="150">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal" width="100">Company Name</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtcompany" class="blue-normal" size="17" style="width:200" value="<%=strName%>" maxlength="40">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">Address</td>
                <td class="blue-normal"> 
                  <textarea name="txtaddress" class="blue-normal" cols="17" style="width:200" rows="2"><%=strAddress%></textarea>
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">City</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtcity" class="blue-normal" size="17" style="width:200" value="<%=strCity%>" maxlength="20">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">State</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtstate" class="blue-normal" size="17" style="width:200" maxlength="20" value="<%=strState%>">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">Country</td>
                <td class="blue-normal"> 
<%Response.Write strlistCountry%>
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">Postal Code</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtpostal" class="blue-normal" size="17" style="width:200" maxlength="10" value="<%=strPostal%>">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">Telephone</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtphone" class="blue-normal" size="17" style="width:200" value="<%=strphone%>" maxlength="50">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">Fax</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtfax" class="blue-normal" size="17" style="width:200" value="<%=strFax%>" maxlength="50">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">E-mail Address</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtemail" class="blue-normal" size="17" style="width:200" value="<%=stremail%>" maxlength="60">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
              <tr> 
                <td class="blue">&nbsp;</td>
                <td class="blue-normal">Web Site</td>
                <td class="blue-normal"> 
                  <input type="text" name="txtweb" class="blue-normal" size="17" style="width:200" value="<%=strweb%>" maxlength="60">
                </td>
                <td class="blue-normal">&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
              <tr> 
                <td height="50"> 
                  <table width="180" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                    <tr> 
                      <td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" width="90"> 
                          <a href="javascript:_save();" class="b">Save Change</a>
                      </td>
                      <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center" width="90">
						<a href="javascript:_reset();" class="b">Reset</a></td>
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
    <td>&nbsp; </td>
  </tr>
</table>
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(1))
			%>
<input type="hidden" name="txthidden" value="<%=intCompanyID%>">
</form>
</body>
</html>
