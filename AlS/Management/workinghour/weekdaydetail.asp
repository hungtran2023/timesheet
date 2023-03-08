<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'****************************************
' function: task_save
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_save()
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
		objDb.cnDatabase.BeginTrans
		strQuery = "UPDATE ATC_Weekday SET fgDayOff = " & CStr(fgdayoff) & ", ratio = " & iRatio & " WHERE WeekdayID = " & itemID
		ret = objDb.runActionQuery(strQuery)
		if ret=false then				
		  objDb.cnDatabase.RollbackTrans
		  gMessage = objDb.strMessage
		else
		  objDb.cnDatabase.CommitTrans
		  gMessage = "Saved successfully."
	    end if
		objDb.dbdisConnect
	End if
	set objDb = nothing
End Sub


	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	Dim itemID, gMessage, fgdayoff

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check ACCESS right
'-----------------------------------
'	tmp = Request.ServerVariables("URL") 
'	while Instr(tmp, "/")<>0
'		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
'	Wend
	tmp =Request.Form("txtpreviouspage")
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if
	
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
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu = "" then strChoseMenu = "A"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")
'-----------------------------------
' Analyse query string
'-----------------------------------
if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
end if

strAct = Request.QueryString("act")
itemID = Request.Form("txthidden")
gMessage=""
if fgExecute then
	if strAct = "SAVE" and fgUpdate then
		iName = Request.Form("txtname")
		iRatio = Request.Form("txtratio")
		if Request.Form("optdayoff") = "1" then
			fgdayoff = 1
		else
			fgdayoff = 0
		end if
		Call task_save
	end if
else
	strAct = ""
end if
'--------------------------------------------------
' Get data from atc_weekday
'--------------------------------------------------
if strAct="" then
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "select * from ATC_Weekday where WeekDayID = " & itemID
	  If objDb.runQuery(strQuery) Then
	    iName = objDb.rsElement("Weekday")
	    iRatio = objDb.rsElement("ratio")
	    if objDb.rsElement("fgdayOff") then
			fgdayoff = 1
		else
			fgdayoff = 0
		end if
	    objDb.closerec()
	  else
		gMessage = objDb.strMessage
	  End if
	  objDb.dbdisConnect
    Else
      gMessage = objDb.strMessage      
    End if
	set objDb = nothing
end if
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


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
<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>
<script LANGUAGE="JavaScript">
function checkdata() {
	if (alltrim(document.frmdetail.txtratio.value)!="") {
		if (isNaN(document.frmdetail.txtratio.value)==true) {
			alert("Please enter a number.");
			document.frmdetail.txtratio.focus();
			return false;
		}
	}
	return true;
}


function save() {
	if(checkdata()==true) {
		document.frmdetail.action = "weekdaydetail.asp?act=SAVE";
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
                 <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
                  <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
                </tr>
                <tr> 
                  <td class="blue" align="left" width="23%"> &nbsp;&nbsp;<a href="workinghours.asp" onMouseOver="self.status='Return the previous page'; return true;" onMouseOut="self.status=''">Weekday 
                    List</a></td>
                  <td class="blue" align="right" width="77%">&nbsp;</td>
                </tr>
                <tr valign="middle">
                  <td class="title" height="50" align="center" colspan="2">Weekday</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td bgcolor="#617DC0"> 
                      <table width="100%" border="0" cellspacing="0" cellpadding="2">
                        <tr bgcolor="#FFFFFF"> 
                          <td valign="top" width="34%" class="blue">&nbsp;</td>
                          <td valign="middle" class="blue" width="13%"> 
                            <%=showlabel(iName)%></td>
                          <td valign="middle" width="53%" class="blue-normal">&nbsp;
                          </td>
                        </tr>                      
                        <tr bgcolor="#FFFFFF"> 
                          <td valign="top" width="34%" class="blue">&nbsp;</td>
                          <td valign="middle" class="blue-normal" width="13%"> 
                            Day off</td>
                          <td valign="middle" width="53%" class="blue-normal"> 
                            <input type="checkbox" name="optdayoff" value="1" <%if fgdayoff=1 then%>checked<%end if%>>
                          </td>
                        </tr>
                        <tr bgcolor="#FFFFFF"> 
                          <td valign="top" width="34%" class="blue">&nbsp;</td>
                          <td valign="middle" class="blue-normal" width="13%">Overtime 
                            Ratio</td>
                          <td valign="middle" width="53%" class="blue-normal"> 
                            <input type="text" name="txtratio" class="blue-normal" size="3" value="<%=showvalue(iRatio)%>">
                          </td>
                        </tr>
                        <tr bgcolor="#FFFFFF"> 
                          <td valign="top" width="34%" class="blue">&nbsp;</td>
                          <td valign="middle" class="blue-normal" width="13%">&nbsp;</td>
                          <td valign="middle" width="53%" class="blue-normal">&nbsp;</td>
                        </tr>
                      </table>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                        <tr> 
                          <td height="50"> 
                            <table width="60" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                              <tr> 
                                <td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" width="59" height="20"> 
<%if fgUpdate then%><a href="javascript:save();" class="b" onMouseOver="self.status='Update'; return true;" onMouseOut="self.status=''">Update</a> 
<%else%>update<%end if%>
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
<input type="hidden" value="<%=iName%>" name="txtname">
<input type="hidden" name="txthidden" value="<%=itemID%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>