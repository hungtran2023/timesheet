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
Function task_save()
	fgSuccessful = false
	if Request.Form("optday") ="1" then
		arrFrom = Array(CInt(Request.Form("lstmonf")), CInt(Request.Form("lstdayf")), CInt(Request.Form("lstyearf")))
		arrTo = Array(CInt(Request.Form("lstmont")), CInt(Request.Form("lstdayt")), CInt(Request.Form("lstyeart")))
	else
		arrDay = Array(CInt(Request.Form("lstmond")), CInt(Request.Form("lstdayd")), CInt(Request.Form("lstyeard")))
	end if
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	'---------------------------------
	' Check whether holiday existed if itemID<>""
	'---------------------------------
	  If itemID<>"" then
		strQuery = "select distinct Holiday from atc_holiday where Holiday not in (select Holiday From atc_holiday where HolidayID = " & itemID & ")"
	  Else
		strQuery = "select distinct Holiday from atc_holiday"
	  End if	
	  If objDb.runQuery(strQuery) Then
			gMessage = ""
			Do while not objDb.rsElement.EOF
				if objDb.rsElement("Holiday") = iName then
					gMessage = "This '" & iName & "' has already been inputted."
					Exit do
				end if
				objDb.rsElement.MoveNext
			Loop
	  Else
		  gMessage = objDb.strMessage
	  End if
	  if gMessage="" then '-----------starting update
		  '-------------------------------
		  ' generate array of date
		  '-------------------------------
		  Dim arrdate()
		  If isEmpty(arrDay) then 'from to
			if arrFrom(0) = arrTo(0) and arrFrom(2) = arrTo(2) then' same month, year
			  kk = -1
			  For ii = arrFrom(1) to arrTo(1)
				kk = kk + 1
				redim preserve arrdate(2, kk)
				arrdate(0, kk) = arrFrom(0)
				arrdate(1, kk) = ii
				arrdate(2, kk) = arrFrom(2)
			  Next
			Else
			  For curyear = arrFrom(2) to arrTo(2)
				if arrFrom(2) = arrTo(2) then ' same year
					beginmonth = arrFrom(0)
					endmonth = arrTo(0)
				else
					If curyear = arrFrom(2) then 'first year
					  beginmonth = arrFrom(0)
					  endmonth = 12
					elseif curyear = arrTo(2) then 'end year
					  beginmonth = 1
					  endmonth = arrTo(0)
					else
					  beginmonth = 1
					  endmonth = 12
					End if
				end if
				For curmonth = beginmonth to endmonth
				  if curmonth = arrFrom(0) then 'first
						curday = arrFrom(1)
						endday = GetDay(curmonth, curyear)
						kk = -1
					elseif curmonth = arrTo(0) then 'end
						curday = 1
						endday = arrTo(1)
						kk = Ubound(arrdate, 2)
					else 'middle
						curday = 1
						endday = GetDay(curmonth, curyear)
						kk = Ubound(arrdate, 2)
					end if
					For ii = curday to endday
						kk = kk + 1
						redim preserve arrdate(2, kk)
						arrdate(0, kk) = curmonth
						arrdate(1, kk) = ii
						arrdate(2, kk) = curyear
					Next
				Next
			  Next
			End if '-----------------same month
		  Else 'date
			redim preserve arrdate(2, 0)
		  	arrdate(0, 0) = arrDay(0)
			arrdate(1, 0) = arrDay(1)
			arrdate(2, 0) = arrDay(2)
		  End if

		  objDb.cnDatabase.BeginTrans
		  ret = false
		  if itemID<>"" then
			strQuery = "DELETE ATC_Holiday WHERE Holiday = (select Holiday from ATC_Holiday Where HolidayID = " & itemID & ")"
			ret = objDb.runActionQuery(strQuery)
		  End if
		  if ret=true or itemID="" then
			for ii = 0 to Ubound(arrdate, 2)
				strQuery = "INSERT INTO ATC_Holiday(Holiday, smonth, sday, syear, ratio) VALUES('" & iName & "', " &_
						arrdate(0, ii) & ", " & arrdate(1, ii) & ", " & arrdate(2, ii) & ", " & iRatio & ")"
				ret = objDb.runActionQuery(strQuery)
				if ret=false then				
				  objDb.cnDatabase.RollbackTrans
				  gMessage = objDb.strMessage
				  exit for
				end if
			next
		  else
			objDb.cnDatabase.RollbackTrans
			gMessage = objDb.strMessage
		  end if
	  end if'-------starting update
	  if gMessage="" then
		  ret = objDb.runQuery("SELECT @@IDENTITY as ID")
		  if ret=true then 
			itemID = objDb.rsElement("ID")
			objDb.cnDatabase.CommitTrans
			gMessage = "Saved successfully."
			fgSuccessful = true
		  end if
	  end if
	  objDb.dbdisConnect
	End if
	set objDb = nothing
	task_save = fgSuccessful
End function
'****************************************
' function: task_delete
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_delete
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  objDb.cnDatabase.BeginTrans
	  strQuery = "DELETE ATC_Holiday WHERE Holiday = (select Holiday from ATC_Holiday Where HolidayID = " & itemID & ")"
	  ret = objDb.runActionQuery(strQuery)
	  if ret=false then
	  	objDb.cnDatabase.RollbackTrans
		gMessage = objDb.strMessage
	  else
		objDb.cnDatabase.CommitTrans
		gMessage = "Deleted successfully."
	  end if
	  objDb.dbdisConnect
	end if
	set objDb = nothing
End sub
'****************************************
' function: task_add
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_add
	itemID=""
	iName=""
	iRatio=""
	strMode = "EDIT"
	fgfromto = false
	intyear = year(now())
	arrlstFrom(0) = selectmonth("lstmonf", 1, 0)
	arrlstFrom(1) = selectday("lstdayf", 1, 0)
	arrlstFrom(2) = selectyear("lstyearf", intyear, intyear, intyear+2, 0)
	arrlstTo(0) = selectmonth("lstmont", 1, 0)
	arrlstTo(1) = selectday("lstdayt", 1, 0)
	arrlstTo(2) = selectyear("lstyeart", intyear, intyear, intyear+2, 0)
	arrlstDay(0) = selectmonth("lstmond", 1, 0)
	arrlstDay(1) = selectday("lstdayd", 1, 0)
	arrlstDay(2) = selectyear("lstyeard", intyear, intyear, intyear+2, 0)
End sub
'****************************************
' function: task_show
' Description: 
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Sub task_show
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
	  strQuery = "select * from ATC_Holiday where Holiday  = (select Holiday from ATC_Holiday Where HolidayID = " & itemID & ") ORDER BY syear, smonth, sday"
	  If objDb.runQuery(strQuery) Then
	    iName = objDb.rsElement("Holiday")
	    iRatio = objDb.rsElement("ratio")
	    numrec = recCount(objDb.rsElement)
	    intyear = year(now())
	    objDb.rsElement.MoveFirst
	    if numrec>1 then
		  fgfromto = true
		  arrlstFrom(0) = selectmonth("lstmonf", objDb.rsElement("smonth"), 0)
	      arrlstFrom(1) = selectday("lstdayf", objDb.rsElement("sday"), 0)
	      arrlstFrom(2) = selectyear("lstyearf", objDb.rsElement("syear"), intyear, intyear+2, 0)
	      objDb.rsElement.Move numrec-1
	      arrlstTo(0) = selectmonth("lstmont", objDb.rsElement("smonth"), 0)
	      arrlstTo(1) = selectday("lstdayt", objDb.rsElement("sday"), 0)
	      arrlstTo(2) = selectyear("lstyeart", objDb.rsElement("syear"), intyear, intyear+2, 0)
	      arrlstDay(0) = selectmonth("lstmond", 1, 0)
	      arrlstDay(1) = selectday("lstdayd", 1, 0)
	      arrlstDay(2) = selectyear("lstyeard", intyear, intyear, intyear+2, 0)
	    else
	      fgfromto = false
	      arrlstFrom(0) = selectmonth("lstmonf", 1, 0)
	      arrlstFrom(1) = selectday("lstdayf", 1, 0)
	      arrlstFrom(2) = selectyear("lstyearf", intyear, intyear, intyear+2, 0)
	      arrlstTo(0) = selectmonth("lstmont", 1, 0)
	      arrlstTo(1) = selectday("lstdayt", 1, 0)
	      arrlstTo(2) = selectyear("lstyeart", intyear, intyear, intyear+2, 0)
	      arrlstDay(0) = selectmonth("lstmond", objDb.rsElement("smonth"), 0)
	      arrlstDay(1) = selectday("lstdayd", objDb.rsElement("sday"), 0)
	      arrlstDay(2) = selectyear("lstyeard", objDb.rsElement("syear"), intyear, intyear+2, 0)
	    end if
	    objDb.closerec()
	  else
		gMessage = objDb.strMessage
	  End if
    Else
      gMessage = objDb.strMessage
    End if
    objDb.dbdisConnect
	set objDb = nothing
end Sub
'-----------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	Dim itemID, arrFrom, arrTo, arrDay, gMessage, fgfromto, strMode, iName, iRatio
	dim arrlstFrom(2), arrlstTo(2), arrlstDay(2)

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
if fgExecute then
	if itemID="" and strAct = "" then strAct = "ADD"
	if not fgUpdate and (strAct="EDIT" or strAct="DELETE" or strAct="ADD" or strAct="SAVE") then strAct = "RESTORE"
else
	strAct = "RESTORE"
end if
gMessage=""

select case strAct
case "RESTORE"
	strMode = Request.Form("txtmode")
	if itemID = "" then
		Call task_add
	else
		Call task_show
	end if
case "EDIT"
	strMode = "EDIT"
	Call task_show
case "DELETE"
	Call task_delete
	Call task_add
case "ADD"
	Call task_add
case "SAVE"
	iName = Request.Form("txtname")
	iRatio = Request.Form("txtratio")
	if task_save()=true then
		strMode = "VIEW"
	else
		strMode = "EDIT"
	End if
	intyear = year(now())
	if isEmpty(arrDay) then
	  fgfromto = true	  
	  arrlstFrom(0) = selectmonth("lstmonf", arrFrom(0), 0)
	  arrlstFrom(1) = selectday("lstdayf", arrFrom(1), 0)
	  arrlstFrom(2) = selectyear("lstyearf", arrFrom(2), intyear, intyear+2, 0)
	  arrlstTo(0) = selectmonth("lstmont", arrTo(0), 0)
	  arrlstTo(1) = selectday("lstdayt", arrTo(1), 0)
	  arrlstTo(2) = selectyear("lstyeart", arrTo(2), intyear, intyear+2, 0)
	  arrlstDay(0) = selectmonth("lstmond", 1, 0)
	  arrlstDay(1) = selectday("lstdayd", 1, 0)
	  arrlstDay(2) = selectyear("lstyeard", intyear, intyear, intyear+2, 0)
	else
	  fgfromto = false
	  arrlstFrom(0) = selectmonth("lstmonf", 1, 0)
	  arrlstFrom(1) = selectday("lstdayf", 1, 0)
	  arrlstFrom(2) = selectyear("lstyearf", intyear, intyear, intyear+2, 0)
	  arrlstTo(0) = selectmonth("lstmont", 1, 0)
	  arrlstTo(1) = selectday("lstdayt", 1, 0)
	  arrlstTo(2) = selectyear("lstyeart", intyear, intyear, intyear+2, 0)
	  arrlstDay(0) = selectmonth("lstmond", arrDay(0), 0)
	  arrlstDay(1) = selectday("lstdayd", arrDay(1), 0)
	  arrlstDay(2) = selectyear("lstyeard", arrDay(2), intyear, intyear+2, 0)
	end if
	if not isEmpty(arrDay) then
		arrDay = empty
	else
		arrFrom = empty
		arrTo = empty
	end if
case ""
	strMode = "VIEW"
	Call task_show
end select
'--------------------------------------------------
' Get data from atc_weekday
'--------------------------------------------------
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
function CheckMode(field){
var varMode="<%=strMode%>";
    if (varMode!="EDIT"){
	  field.blur();
      if (field.type!="text") frmdetail.txtname.focus();
    }
    else {
		var strtmp1 = field.name
		var strtmp2 = strtmp1.substr(strtmp1.length - 1 , 1);
		if((strtmp2=="f")||strtmp2=="t")
			document.frmdetail.optday[0].checked = true;
		else
			document.frmdetail.optday[1].checked = true;
    }
}

function edit() {
var varMode="<%=strMode%>";
	if(varMode!="EDIT") {
		document.frmdetail.action = "holidaydetail.asp?act=EDIT";
		document.frmdetail.target = "_self";
		document.frmdetail.submit();
	}
}

function checkday(k) {
	varM1 = "a04a06a09a11a";
	if (k==1) {
	  varmonth1 = document.frmdetail.lstmonf.options[document.frmdetail.lstmonf.selectedIndex].value;
	  varday1 = document.frmdetail.lstdayf.options[document.frmdetail.lstdayf.selectedIndex].value;
	  varyear1 = document.frmdetail.lstyearf.options[document.frmdetail.lstyearf.selectedIndex].value;
	  tmp = "a"+varmonth1+"a";
	  if ((varM1.indexOf(tmp)!=-1) && (varday1>30)) return false;
	  if ((varmonth1==2)&&(varday1>28)) return false;
	  varmonth2 = document.frmdetail.lstmont.options[document.frmdetail.lstmont.selectedIndex].value;
	  varday2 = document.frmdetail.lstdayt.options[document.frmdetail.lstdayt.selectedIndex].value;
	  varyear2 = document.frmdetail.lstyeart.options[document.frmdetail.lstyeart.selectedIndex].value;
	  tmp = "a"+varmonth2+"a";
	  if ((varM1.indexOf(tmp)!=-1) && (varday2>30)) return false;
	  if ((varmonth2==2)&&(varday2>28)) return false;
	  if (varyear1>varyear2) return false;
	  if (((varyear1==varyear2)&&(varmonth1>varmonth2))||((varmonth1==varmonth2)&&(varday1>varday2))) return false;
	}
	else {
	  varmonth1 = document.frmdetail.lstmond.options[document.frmdetail.lstmond.selectedIndex].value;
	  varday1 = document.frmdetail.lstdayd.options[document.frmdetail.lstdayd.selectedIndex].value;
	  tmp = "a"+varmonth1+"a";
	  if ((varM1.indexOf(tmp)!=-1)&&(varday1>30)) return false;
	  if ((varmonth1==2)&&(varday1>28)) return false;
	}
	return true;
}

function checkdata() {
	if (alltrim(document.frmdetail.txtratio.value)!="") {
		if (isNaN(document.frmdetail.txtratio.value)==true) {
			alert("Please enter a number.");
			document.frmdetail.txtratio.focus();
			return false;
		}
	}
	else {
			alert("Please enter a number.");
			document.frmdetail.txtratio.focus();
			return false;
	}
	
	if (alltrim(document.frmdetail.txtname.value)=="") {
		alert("This field can't be empty.");
		document.frmdetail.txtname.focus();
		return false;
	}
	
	if (document.frmdetail.optday[0].checked==true) {
		if (checkday(1)==false) {
			alert("Invalid data.");
			document.frmdetail.lstmond.focus();
			return false;
		}
	}
	else {
		if (checkday(2)==false) {
			alert("Invalid data.");
			document.frmdetail.lstmonf.focus();
			return false;
		}
	}
	
	return true;
}

function mydelete() {
	if(document.frmdetail.txthidden.value!="") {
		document.frmdetail.action = "holidaydetail.asp?act=DELETE";
		document.frmdetail.target = "_self";
		document.frmdetail.submit();
	}
}

function add() {
	document.frmdetail.action = "holidaydetail.asp?act=ADD"
	document.frmdetail.target = "_self";
	document.frmdetail.submit();
}
function save() {
	if(checkdata()==true) {
		document.frmdetail.action = "holidaydetail.asp?act=SAVE";
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
                <tr align="center"> 
                  <td class="blue" align="left" width="23%"> &nbsp;&nbsp;<a href="workinghours.asp" onMouseOver="self.status='Return the previous page'; return true;" onMouseOut="self.status=''">Holiday 
                    List</a></td>
                  <td class="blue" align="right" width="77%">&nbsp;</td>
                </tr>
                <tr valign="middle"> 
                  <td class="title" height="50" align="center" colspan="2">Holiday Detail</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%">
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="18%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="17%"> 
                                Name</td>
                              <td valign="middle" width="65%" class="blue-normal"> 
                                <input type="text" name="txtname" maxlength="30" class="blue-normal" size="20" value="<%=showvalue(iName)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="18%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="17%">Overtime 
                                Ratio</td>
                              <td valign="middle" width="65%" class="blue-normal"> 
                                <input type="text" name="txtratio" class="blue-normal" size="5" value="<%=showvalue(iRatio)%>" <%if strMode<>"EDIT" then%>onFocus="CheckMode(this)"<%end if%>>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="middle" width="18%" class="blue" align="right"> 
                                <input type="radio" name="optday" value="1" <%if fgfromto=true then%>checked<%end if%>>
                              </td>
                              <td valign="middle" class="blue-normal" width="17%">From</td>
                              <td valign="middle" width="65%" class="blue-normal"> 
<%Response.Write arrlstFrom(1)
  Response.Write arrlstFrom(0)
  Response.Write arrlstFrom(2)
%>
                                To 
<%Response.Write arrlstTo(1)
  Response.Write arrlstTo(0)
  Response.Write arrlstTo(2)
%>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="middle" width="25%" class="blue" align="right"> 
                                <input type="radio" name="optday" value="2" <%if fgfromto=false then%>checked<%end if%>>
                              </td>
                              <td valign="middle" class="blue-normal" width="14%">Date</td>
                              <td valign="middle" width="61%" class="blue-normal"> 
<%Response.Write arrlstDay(1)
  Response.Write arrlstDay(0)
  Response.Write arrlstDay(2)
%>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="240" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr>
<%if fgUpdate then%>
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue">
									 <a href="javascript:add();" class="b" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add</a></td>
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
<%if itemID<>"" then%>				<a href="javascript:edit();" class="b" onMouseOver="self.status='Edit'; return true;" onMouseOut="self.status=''">Edit</a>
<%else%>Edit<%end if%></td>
									<td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
<%if strMode<>"VIEW" then%><a href="javascript:save();" class="b" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status=''">Save</a>
<%else%>Save<%end if%></td>
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
<%if itemID<>"" then%>				<a href="javascript:mydelete();" class="b" onMouseOver="self.status='Delete'; return true;" onMouseOut="self.status=''">Delete</a>
<%else%>Delete<%end if%></td>
<%else%>
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue">
									 Add</td>
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
									Edit</td>
									<td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
									Save</td>
                                    <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
									Delete</td>
<%end if%>
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
<input type="hidden" name="txthidden" value="<%=itemID%>">
<input type="hidden" name="txtmode" value="<%=strMode%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">
</form>
</body>
</html>