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
	
'-----------------------------------
'Check MANAGER right
'-----------------------------------
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		fgRight = false
		getRight = session("Righton")
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = "receive report" then
				fgRight=true
				exit for
			end if
		next
	end if
	
'----------------------------------
' Get Full Name and Job Title
'----------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
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
Call freeListpro
Call freeProInfo
Call freeAssignment
Call freeAssignRight
Call freeShort
Call freeListEmp
Call freeSinglepro
Call freeSumpro

strAct = Request.QueryString("act")
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
	strAct = ""
else
	fgExecute = true
end if


gMessage=""
if strAct = "SAVE" then
	strURL = Request.Form("lstpage")
	if fgRight then
	  introw = int(Request.Form("txtrow"))
	  strfilter = ""	'Request.Form("chkset")
	else
	  introw = 0
	  strfilter = ""
	end if

	if strURL="" then strURL_I = "Null" else strURL_I = "'" & strURL & "'"
	if introw = 0 then introw_I = "Null" else introw_I = introw
	if strfilter<>"" then
		if isEmpty(session("filter")) then
			varPro_I = "Null"
		else
			varPro_I = "'" & session("filter") & "'"
		end if
		if isEmpty(session("filteremp")) then
			varEmp_I = "Null"
		else
			varEmp_I = "'" & session("filteremp") & "'"
		end if
	else
		varPro_I = "Null"
		varEmp_I = "Null"
	end if
		
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "Select count(*) as mysum From ATC_Preferences Where StaffID = " & session("USERID")
	  ret = objDb.runQuery(strQuery)
	  if ret then
		if objDb.rsElement("mysum")=1 then '--------------starting update
			objDb.closerec
			strQuery = "UPDATE ATC_preferences SET FavoriteUrl = " & strURL_I & ", NumofRows =" & introw_I & ", ProCriteria = " &_
				varPro_I & ", EmpCriteria = " & varEmp_I & " WHERE StaffID = " & session("USERID")
		else
			strQuery = "INSERT INTO ATC_preferences(StaffID, FavoriteUrl, NumofRows, ProCriteria, EmpCriteria) VALUES(" &_
						session("USERID") & ", " & strURL_I & ", " & introw_I & ", " & varPro_I & ", " & varEmp_I & ")"
		end if
		objDb.cnDatabase.BeginTrans
		ret = objDb.runActionQuery(strQuery)
		if ret=false then				
			objDb.cnDatabase.RollbackTrans
			gMessage = objDb.strMessage
		else
			objDb.cnDatabase.CommitTrans
			gMessage = "Updated successfully."
			session("Preferences") = empty
			If isEmpty(session("Preferences")) then
				getPre = getarrPreference(session("USERID"))
				if isArray(getPre) then session("Preferences") = getPre
			End if
		end if
	  else
	    gMessage = objDb.strMessage
	  end if
	  objDb.dbdisConnect
	Else
	  gMessage = objDb.strMessage
	End if
	set objDb = nothing
else
'------------------------------
' get data
'------------------------------
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	If objDb.dbConnect(strConnect) then
	  strQuery = "Select isnull(FavoriteUrl, '') as FavoriteUrl, isnull(NumofRows, 0) as NumofRows, isnull(ProCriteria, '') as ProCriteria, " &_
				"isnull(EmpCriteria, '') as EmpCriteria From ATC_Preferences Where StaffID = " & session("USERID")
	  ret = objDb.runQuery(strQuery)
	  if ret then
		if not objDb.noRecord then
			strURL = objDb.rsElement("FavoriteUrl")
			introw = objDb.rsElement("NumofRows")
			strfilter = objDb.rsElement("ProCriteria")
			if strfilter="" then strfilter = objDb.rsElement("EmpCriteria")
		else
		  strURL = ""
		  introw = 0
		  strfilter = ""
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


strlistpage = "<Select name='lstpage' size='1' class='blue-normal' style='width:160'>"
if strURL="" then strSel = " selected " else strSel=""
strlistpage = strlistpage & "<option value=''" & strSel & ">None</option>" & chr(13)

if cint(Session("GroupManager"))=-1 then
	if strURL="tms/timesheet.asp" then strSel = " selected " else strSel=""
	strlistpage = strlistpage & "<option value='tms/timesheet.asp'" & strSel & ">Complete Timesheet</option>" & chr(13)
	strlistpage = strlistpage & "<option value='aisnet/DashBoard/DashBoard'" & strSel & ">Dashboard</option>" & chr(13)
	if strURL="assignedproject.asp" then strSel = " selected " else strSel=""
	strlistpage = strlistpage & "<option value='assignedproject.asp'" & strSel & ">Assigned Projects</option>" & chr(13)
end if
for ii=0 to ubound(getRes, 2)
  if getRes(0, ii) = 0 then
	if strURL=getRes(2, ii) then strSel = " selected " else strSel=""
	strlistpage = strlistpage & "<option value='" & getRes(2, ii) & "'" & strSel & ">" & getRes(1, ii) & "</option>" & chr(13)
  end if
next
strlistpage = strlistpage & "</select"

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
varpage = "<%=strURL%>";
varrow = "<%=introw%>";
varfilter = "<%=strfilter%>";

  tmp = document.frmdetail.lstpage.options[document.frmdetail.lstpage.selectedIndex].value;
  if(tmp!=varpage) return true;
	
  if(document.frmdetail.txtrow) {
	document.frmdetail.txtrow.value = alltrim(document.frmdetail.txtrow.value);
	if ((document.frmdetail.txtrow.value!="")&&(document.frmdetail.txtrow.value!=varrow)) {
	  if ((!isNaN(document.frmdetail.txtrow.value))&&(document.frmdetail.txtrow.value>=5)) {return true;}
	  else {
		alert("Please enter a number larger than 4.");
		document.frmdetail.txtrow.focus();
		return false;
	  }
	}
	//if (((varfilter=="")&&(document.frmdetail.chkset.checked))||((varfilter!="")&&(!document.frmdetail.chkset.checked))) return true;
  }
  return false;

}


function save() {
	if(checkdata()==true) {
		document.frmdetail.action = "preferences.asp?act=SAVE";
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
                  <td class="title" height="50" align="center" colspan="2">Preferences Settings </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top" align="center">
				<table width="100%" border="0" cellspacing="0" cellpadding="5">
				  <tr> 
				    <td width="10%" class="blue">&nbsp;Options </td>
				    <td width="90%" class="blue-normal">&nbsp;</td>
				  </tr>
				  <tr> 
				    <td width="10%" class="blue-normal" align="right" height="40">&nbsp;</td>
				    <td width="90%" class="blue-normal" height="40">When logging 
				      in, go directly to 
<%Response.Write strlistpage%>
				    </td>
				  </tr>
<%if fgRight then%>
<!--				  <tr> 
				    <td width="10%" class="blue-normal" align="right" height="40"> 
				      <input type="checkbox" name="chkset" value="1" <%if strfilter<>"" then%>checked<%end if%>>
				    </td>
				    <td width="90%" class="blue-normal" height="40">Use current 
				      filter criteria as default</td>
				  </tr> -->
				  <tr> 
				    <td width="10%">&nbsp;</td>
				    <td width="90%" class="blue-normal" nowrap>Number of rows on each page 
				      <input type="text" name="txtrow" style="width:30" size="3" maxlength="3" value="<%=introw%>">
				    </td>
				  </tr>
<!--				  <tr> 
				    <td width="10%" class="blue">&nbsp;Shortlist</td>
				    <td width="90%" class="blue" align="right">&nbsp;</td>
				  </tr>
				  <tr> 
				    <td width="10%" class="blue">&nbsp;</td>
				    <td width="90%" class="blue-normal" align="left"><a href="shortlist.asp" onMouseOver="self.status='Shortlist'; return true;" onMouseOut="self.status=''">Click 
				      here</a> to view or update Shortlist</td>
				  </tr>-->
<%end if%>
				</table>
                <table width="60" border="0" cellspacing="5" cellpadding="0" height="20">
                  <tr> 
                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
                    <a href="javascript:save();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Save</a></td>
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