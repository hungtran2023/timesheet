<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
'****************************************
' Function: outbody
' Description: 
' Parameters: array data, page size, which page
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
Function OutBody(ByRef arrSrc, ByVal PSize, ByVal Whichpage)
	strOut = ""
	i = (Whichpage - 1)*PSize	
	lastU = Ubound(arrSrc, 2)
	cnt = 0
	Do Until i>lastU
		cnt = cnt + 1
		if i < lastU then
			strColor = "#FFFFFF"
			strTmp = "<td valign='top' width='21%' class='blue-normal'>&nbsp;" & showlabel(arrSrc(0, i)) & "</td>" &_
					"<td valign='top' width='43%' class='blue-normal'>&nbsp;" & showlabel(arrSrc(1, i)) & "</td>" &_
					"<td valign='top' width='12%' class='blue-normal' align='right'>" & FormatNumber(arrSrc(2, i), 2) & "</td>" &_
					"<td valign='top' width='12%' class='blue-normal' align='right'>" & FormatNumber(arrSrc(3, i), 2) & "</td>" &_
					"<td valign='top' width='12%' class='blue-normal' align='right'>" & FormatNumber(CSng(arrSrc(2, i)) + CSng(arrSrc(3, i)), 2) & "</td>"
		else
			strColor = "#FFF2F2"
			strTmp = "<td valign='top' colspan='2' class='blue' align='right'>" & arrSrc(1, i) & "</td>" &_
					"<td valign='top' width='12%' class='blue' align='right'>" & FormatNumber(arrSrc(2, i), 2) & "</td>" &_
					"<td valign='top' width='12%' class='blue' align='right'>" & FormatNumber(arrSrc(3, i), 2) & "</td>" &_
					"<td valign='top' width='12%' class='blue' align='right'>" & FormatNumber(CSng(arrSrc(2, i))+CSng(arrSrc(3, i)), 2) & "</td>"
		end if
		strOut = strOut & "<tr bgcolor='" & strColor & "'>" & strTmp & "</tr>"
		i = i + 1
		if cnt = pSize then exit do
	Loop
	Outbody = strOut
End Function
'****************************************
' Function: selectfullmonth
' Description: 
' Parameters:
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectfullmonth(ByVal vname, Byval vselected)
	arrlongmon  = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	strlst = "<select name='" & vname & "' size='1' height='26px' width='75px' " &_
			"style='width:75px;height=24px;' class='blue-normal' onClick='document.frmreport.opttime[1].checked=true;'>"
	strTmp = ""
	For i = 1 to 12
		strTmp1 = CStr(i)
		if len(strTmp1) = 1 then strTmp1 = "0" & strTmp1 end if
		if i = vselected then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & strTmp1 & "' " & strSel & ">" & arrlongmon(i-1) & "</option>"
	Next
	strlst = strlst & strTmp & "</select>"
	set arrlongmon = nothing
	selectfullmonth = strlst
end function
'****************************************
' Function: selectyear
' Description: 
' Parameters:
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function selectlistyear(ByVal vname, Byval vselected, Byref arrSrc)
	strlst = "<select name='" & vname & "' size='1' height='26px' width='70px' " &_
			"style='width:70px;height=24px;' class='blue-normal' onClick='document.frmreport.opttime[1].checked=true;'>"
	strTmp = ""
	For i = 0 to ubound(arrSrc)
		if arrSrc(i) = int(vselected) then strSel = "selected" else strSel = "" end if
		strTmp = strTmp & "<option value='" & arrSrc(i) & "' " & strSel & ">" & arrSrc(i) & "</option>"
	Next
	strlst = strlst & strTmp & "</select>"
	selectlistyear = strlst
end function

	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If

'-------------------------------
' Calculate pagesize
'-------------------------------
	if not isEmpty(session("Preferences")) then
		arrPre = session("Preferences")
		if arrPre(1, 0)>0 then PageSize = arrPre(1, 0) else PageSize = PageSizeDefault
		set arrPre = nothing
	else
		PageSize = PageSizeDefault
	end if
	
'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 
	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../welcome.asp")
	end if
	
'----------------------------------
' Get Full Name and Job Title
'----------------------------------

	Set objEmployee = New clsEmployee
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'><a href='../welcome.asp?choose_menu=B' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				"<a class='c' href='javascript:_print();' onMouseOver='self.status=&quot;Print report&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Print</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Preparing data
'--------------------------------------------------
	Call freeListpro
	Call freeProInfo
	Call freeAssignment
	Call freeAssignRight
	Call freeListEmp
	Call freeShort
	Call freeSinglepro
if Request.TotalBytes=0 or Request.QueryString("outside")<>"" then
	Call freeSumpro
end if

stract = Request.QueryString("act")
if stract = "REFRESH" then session("READYSUMPRO") = false

stropt = Request.Form("opttime")
strfrom = ""
strto = ""
if stropt = "" then stropt = "1"
if stropt = "0" then 'from to
	strfrom = Request.Form("txtfrom")
	strto = Request.Form("txtto")
	arrTmp = split(strfrom, "/")	
	strfrom_I = arrTmp(1) & "/" & arrTmp(0) & "/" & arrTmp(2)
	arrTmp = split(strto, "/")	
	strto_I = arrTmp(1) & "/" & arrTmp(0) & "/" & arrTmp(2)
	
	byear = year(Cdate(strfrom_I))
	eyear = year(Cdate(strto_I))
		
	strwhere = ""
	gMessage = ""
	strConnect = Application("g_strConnect") 
  Set objDb = New clsDatabase
  If objDb.dbConnect(strConnect) then
	  strQuery = "select TMS_Table from ATC_Index where year(start_date)>=" & byear & " AND year(start_date)<=" & eyear & " order by TMS_Table"
      If objDb.runQuery(strQuery) Then
		If not objDb.noRecord then
		  if eyear>=year(now()) then 'have current year
			strtmp = objDb.rsElement(0)
			strwhere = "(select assignmentID, Overtime, Hours, Tdate from " & strtmp &_
					" where assignmentID>1 and (tdate>='" & strfrom_I & "' and tdate<='" & strto_I & "') "
			objDb.rsElement.MoveNext
			Do until objDb.rsElement.EOF
				strtmp = objDb.rsElement(0)
				strwhere = strwhere & "UNION select assignmentID, Overtime, Hours, Tdate from " & strtmp &_
							" where assignmentID>1 and (tdate>='" & strfrom_I & "' and tdate<='" & strto_I & "') "
				objDb.rsElement.MoveNext
			Loop
			strwhere = strwhere & "UNION select assignmentID, Overtime, Hours, Tdate from atc_timesheet " &_
					"where assignmentID>1 and (tdate>='" & strfrom_I & "' and tdate<='" & strto_I & "')) a "
		  else 'no current year
			strtmp = objDb.rsElement(0)
			strwhere = "(select assignmentID, Overtime, Hours, Tdate from " & strtmp &_
					" where assignmentID>1 and (tdate>='" & strfrom_I & "' and tdate<='" & strto_I & "') "
			objDb.rsElement.MoveNext
			Do until objDb.rsElement.EOF
				strtmp = objDb.rsElement(0)
				strwhere = strwhere & "UNION select assignmentID, Overtime, Hours, Tdate from " & strtmp &_
							" where assignmentID>1 and (tdate>='" & strfrom_I & "' and tdate<='" & strto_I & "') "
				objDb.rsElement.MoveNext
			Loop
			strwhere = strwhere & " ) a "
		  end if
		else 'no record from table ATC_index
			if eyear>=year(now()) then 'have current year
				strwhere = "(select assignmentID, Overtime, Hours, Tdate from ATC_Timesheet " &_
					" where assignmentID>1 and (tdate>='" & strfrom_I & "' and tdate<='" & strto_I & "')) a "
			end if
		end if
	  Else
		gMessage = objDb.strMessage
	  end if
	objDb.dbDisconnect
  Else
	gMessage = objDb.strMessage
  End if
  set objDb = nothing		
else 'month
	strmonth = Request.Form("lstmonth")
	if strmonth = "" then strmonth = month(now())
	arrlongmon  = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	strnameofmonth = arrlongmon(int(strmonth)-1)
	set arrlongmon = nothing
	stryear = Request.Form("lstyear")
	if stryear = "" then stryear = Cstr(year(now())) ' for first run
	strtmp = stryear	
	if int(stryear) = year(now()) then strtmp = ""
	strwhere = "(select assignmentID, Overtime, Hours, Tdate from atc_timesheet" & strtmp &" where " &_
				"assignmentID>1 and month(tdate) = " & strmonth & " ) a "
end if

if strwhere = "" then 'if from to have error then switch to month
	if gMessage="" then gMessage = "No data from " & strfrom & " to " & strto
	stropt = "1"
	strmonth = Request.Form("lstmonth")
	if strmonth = "" then strmonth = month(now())
	arrlongmon  = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	strnameofmonth = arrlongmon(int(strmonth)-1)
	set arrlongmon = nothing
	stryear = Request.Form("lstyear")
	if stryear = "" then stryear = Cstr(year(now())) ' for first run
	strtmp = stryear	
	if int(stryear) = year(now()) then strtmp = ""
	strwhere = "(select assignmentID, Overtime, Hours, Tdate from atc_timesheet" & strtmp &" where " &_
				"assignmentID>1 and month(tdate) = " & strmonth & " ) a "
end if

if stropt = "0" then 
	strfromto = "from " & strfrom & " to " & strto
else 
	strfromto = strnameofmonth & " - " & stryear
end if
strprintdate = FormatDateTime(Date, 1)	'day(date()) & "/" & month(date()) & "/" & year(date())
strlstmonth = selectfullmonth("lstmonth", int(strmonth))

If isEmpty(session("READYSUMPRO")) or session("READYSUMPRO")=false then
  strConnect = Application("g_strConnect") 
  Set objDb = New clsDatabase
  If objDb.dbConnect(strConnect) then 
	strQuery = "select d.ProjectID, d.projectname, sum(a.Hours) as hours, sum(a.Overtime) as overtime " &_
				"from " & strwhere &_
				"Left join ATC_Assignments b On a.AssignmentID = b.AssignmentID " &_
				"left join ATC_Tasks c On b.SubTaskID = c.SubTaskID " &_
				"left join ATC_Projects d On c.ProjectID = d.ProjectID " &_
				"where d.ProjectID is not null group by d.projectID, d.projectname " &_
				"order by d.projectID"
'Response.Write STRqUERY	
'Response.End			
    if objDb.runQuery(strQuery) then
		if not objDb.noRecord then 
			Dim arrData()
			cnt = -1
			OverallHour = 0
			OverallOver = 0
			Do Until objDb.rsElement.EOF
			  cnt = cnt + 1
			  Redim preserve arrData(3, cnt)
			  arrData(0, cnt) = objDb.rsElement("ProjectID")
			  arrData(1, cnt) = objDb.rsElement("projectname")
			  arrData(2, cnt) = objDb.rsElement("Hours")
			  arrData(3, cnt) = objDb.rsElement("Overtime")
			  OverallHour = OverallHour + CDbl(objDb.rsElement("Hours"))
			  OverallOver = OverallOver + CDbl(objDb.rsElement("Overtime"))
			  objDb.MoveNext
			Loop
			'row for overall total
			cnt = cnt + 1
			Redim preserve arrData(3, cnt)
			arrData(0, cnt) = ""
			arrData(1, cnt) = "Overall Total: "
			arrData(2, cnt) = OverallHour
			arrData(3, cnt) = OverallOver
			session("arrSumPro") = arrData
			session("NumPageSumPro") = PageCount(arrData, PageSize)
			session("CurpageSumPro") = 1
			objDb.closeRec
		else
			session("NumPageSumPro") = 0
			session("CurpageSumPro") = 0
			session("arrSumPro") = empty
		end if 'test have data
	else 'error on open query
	  gMessage = objDb.strMessage
	end if
	objDb.dbDisconnect
  else 'error on connection
	gMessage = objDb.strMessage
  end if
  session("READYSUMPRO") = true
  set objDb = nothing
end if

'--------------------
' get table ATC_Index
'--------------------
if isEmpty(session("arryearValid")) then
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	Dim arryearValid()
	cnt = -1
	If objDb.dbConnect(strConnect) then
		strQuery = "select start_date from ATC_Index"
		If objDb.runQuery(strQuery) Then
			If not objDb.noRecord then
				Do until objDb.rsElement.EOF
					cnt = cnt + 1
					Redim preserve arryearValid(cnt)
					arryearValid(cnt) = year(objDb.rsElement(0))
					objDb.MoveNext
				Loop
				objDb.closerec
			end if
		Else
		  gMessage = objDb.strMessage
		end if
		objDb.dbDisconnect
	Else
		gMessage = objDb.strMessage
	End if
	set objDb = nothing
	'For current year
	cnt = cnt + 1
	Redim preserve arryearValid(cnt)
	arryearValid(cnt) = year(now())
	session("arryearValid") = arryearValid
end if
arrtmp = session("arryearValid")
strlstyear = selectlistyear("lstyear", stryear, arrTmp)
set arrtmp = nothing

if isEmpty(session("arrInfoCompany")) then
	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
	If objDb.dbConnect(strConnect) then
		strQuery = "select a.CompanyName, isnull(Address,'') Address, isnull(City,'') City, isnull(b.CountryName,'') Country, " &_
					"isnull(Phone,'') Phone, isnull(Fax,'') Fax, isnull(c.Logo,'') Logo from ATC_Companies a " &_
					"left join ATC_Countries b On a.CountryID = b.CountryID " &_
					"left join ATC_CompanyProfile c ON a.CompanyID = c.CompanyID " &_
					"where a.CompanyID = " & session("Inhouse")
		If objDb.runQuery(strQuery) Then
			If not objDb.noRecord then
				arrInfoCompany = objDb.rsElement.getRows
				session("arrInfoCompany") = arrInfoCompany
				objDb.closerec
			end if
		Else
		  gMessage = objDb.strMessage
		end if
		objDb.dbDisconnect
	Else
		gMessage = objDb.strMessage
	End if
	set objDb = nothing
end if

varNavi = Request.QueryString("navi")
if varNavi <> "" then
	tmpi = session("CurPageSumPro")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPageSumPro") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPageSumPro")
			end if
	End select
	session("CurPageSumPro") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPageSumPro") = CInt(varGo)

if not isEmpty(session("arrSumPro")) then
  arrGet = session("arrSumPro")
  strLast = OutBody(arrGet, PageSize, session("CurpageSumPro"))
  set arrGet = nothing
end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_report.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
if not isEmpty(session("arrInfoCompany")) then
	arrTmp = session("arrInfoCompany")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", arrTmp(0, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", arrTmp(1, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", arrTmp(2, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", arrTmp(3, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", arrTmp(4, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", arrTmp(5, 0))
	if arrTmp(6, 0)<>"" then
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "<img src='../images/" & arrTmp(6, 0) & "' border='0'>" )
	else
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;" )
	end if
	set arrTmp = nothing
else
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;")
end if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link href="../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>
<script src="../jQuery/jquery.min.js"></script>
<script src="../jQuery/jquery-ui.min.js"></script>
<link href="../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<script>
	$(function() {
		var dates = $( "#txtfrom, #txtto" ).datepicker({
		dateFormat: "dd/mm/yy",
		minDate:"1/1/2000",
		maxDate: "31/12/<%=year(Date())%>",
        onSelect: function( selectedDate ) {
                                    var option = this.id == "txtfrom" ? "minDate" : "maxDate",
                                        instance = $( this ).data( "datepicker" ),
                                        date = $.datepicker.parseDate(
                                            instance.settings.dateFormat ||
                                            $.datepicker._defaults.dateFormat,
                                            selectedDate, instance.settings );
                                    dates.not( this ).datepicker( "option", option, date );
                                }
            });
	});
</script>

<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script>
var objWindowSumPro;

function _print() { //v2.0
var str1 = "<%=strfromto%>";
	str1 = escape(str1);
var str2 = "<%=strprintdate%>";
	str2 = escape(str2);
var fgprint = <%=session("NumPageSumPro")%>;
if (fgprint!=0) {
	window.status = "";
	strFeatures = "top="+(screen.height/2-275)+",left="+(screen.width/2-390)+",width=800,height=550,toolbar=no," 
	            + "menubar=yes,location=no,directories=no,scrollbars=yes,status=yes";
	if ((objWindowSumPro) && (!objWindowSumPro.closed)) {
		objWindowSumPro.focus();
	
	} else {
		objWindowSumPro = window.open("p_sumworkhours.asp?fromto=" + str1 + "&printdate=" + str2, "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";
  }
else
	alert("No data for your request.")
}

function window_onunload() {
	if((objWindowSumPro) && (!objWindowSumPro.closed))
		objWindowSumPro.close();
}

function checkdata()
{
	if (document.frmreport.opttime[0].checked)
	{
		if (isnull(document.frmreport.txtfrom.value)==true)
		{
			alert("Please enter startdate before click here.")
			document.frmreport.txtfrom.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtfrom.value)==false)
			{			
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtfrom.focus();
				return false;
			}
		}
		
		if (isnull(document.frmreport.txtto.value)==true)
		{
			alert("Please enter enddate before click here.")
			document.frmreport.txtto.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtto.value)==false)
			{
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtto.focus();
				return false;
			}
		}
		
		if (comparedate(document.frmreport.txtfrom.value,document.frmreport.txtto.value)==false)
		{
			alert("The startdate must be less than the finishdate.")
			document.frmreport.txtfrom.focus();
			return false;
		}
	}	
	return true;
}

function _submit() {
	if(checkdata()==true) {
		document.frmreport.action = "sumworkhours.asp?act=REFRESH";
		document.frmreport.target = "_self" ;
		document.frmreport.submit();
	}
}

function next() {
var curpage = <%=session("CurPageSumPro")%>;
var numpage = <%=session("NumPageSumPro")%>;
	if (curpage < numpage) {
		document.frmreport.action = "sumworkhours.asp?navi=NEXT"
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function prev() {
var curpage = <%=session("CurPageSumPro")%>;
var numpage = <%=session("NumPageSumPro")%>;
	if (curpage > 1) {
		document.frmreport.action = "sumworkhours.asp?navi=PREV";
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function go() {
	var numpage = <%=session("NumPageSumPro")%>;
	var curpage = <%=session("CurPageSumPro")%>;
	var intpage = document.frmreport.txtpage.value;
	intpage = parseInt(intpage, 10);
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmreport.action = "sumworkhours.asp?Go=" + intpage;
		document.frmreport.target = "_self";
		document.frmreport.submit();		
	}
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE=javascript onunload="return window_onunload()">
<form name="frmreport" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>
	
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%"  align="center" >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="36%">
              <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td valign="top" width="0%" class="blue" height="33">&nbsp;</td>
                  <td width="4%" class="blue" align="right" height="33"> 
                    <input type="radio" name="opttime" value="0" <%if stropt="0" then%>checked<%end if%>>
                  </td>
                  <td class="blue-normal" width="8%" height="33">From </td>
                  <td class="blue-normal" width="40%" height="33"> 
                    <input type="text" name="txtfrom" id="txtfrom" class="blue-normal"  size="5" style="width:60" value="<%=strfrom%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                  <td width="8%" class="blue-normal" height="33"> To </td>
                  <td class="blue-normal" height="33" width="40%"> 
                    <input type="text" name="txtto" id="txtto" class="blue-normal" size="5" style="width:60" value="<%=strto%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                </tr>
                <tr> 
                  <td valign="top" width="0%" class="blue">&nbsp;</td>
                  <td width="4%" class="blue" align="right"> 
                    <input type="radio" name="opttime" value="1" <%if stropt="1" then%>checked<%end if%>>
                  </td>
                  <td class="blue-normal" width="8%">Month </td>
                  <td class="blue-normal" width="40%"> 
<%Response.Write strlstmonth%>
                  </td>
                  <td width="8%" class="blue-normal"> Year </td>
                  <td class="blue-normal" width="40%">
<%Response.Write strlstyear%>
                  </td>
                </tr>
              </table>
            </td>
            <td width="64%">
              <table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
                <tr> 
                  <td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
                      <a href="javascript:_submit();" class="b">Submit</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
			<td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		  </tr>
          <tr> 
            <td bgcolor="8CA0D1"><img src="../IMAGES/DOT-01.GIF" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>&nbsp; </td>
          </tr>
        </table>
    		<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Hours")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strfromto)
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
			%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                <tr> 
                  <td class="blue" width="21%" bgcolor="#E7EBF5">&nbsp;Project ID </td>
                  <td class="blue" width="43%" bgcolor="#E7EBF5">&nbsp;Project Name </td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5">Hours</td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5">Overtime</td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5">Total</td>
                </tr>
<%Response.Write strLast%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<%if session("NumPageSumPro")>0 then%>
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="20" align="center">
    <tr> 
      <td align="right" bgcolor="#E7EBF5"> 
        <table width="70%" border="0" cellspacing="1" cellpadding="0" height="18">
          <tr> 
            <td align="right" valign="middle" width="37%" class="blue-normal">Page 
            </td>
            <td align="center" valign="middle" width="13%" class="blue-normal"> 
              <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageSumPro")%>" size="2" style="width:50">
            </td>
            <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();"  onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
            </td>
            <td align="right" valign="middle" width="15%" class="blue-normal">Pages 
               <%=session("CurpageSumPro")%>/<%=session("NumpageSumPro")%>&nbsp;&nbsp;</td>
            <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();"  onMouseOver="self.status='Previous page'; return true;" onMouseOut="self.status='';">Previous</a> 
              /<a href="javascript:next();" onMouseOver="self.status='Next page'; return true;" onMouseOut="self.status='';"> Next</a>&nbsp;&nbsp;&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
</table>
<%end if%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>
</body>
</html>