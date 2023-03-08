<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->

<%

'*********************************************************
'Get duration
'*********************************************************
sub Getperiod()
	dim arrQuater
	
	'Option=1: From --> to
	'Option=0: Quarter
	stropt=Request.Form("opttime")
	stropt=IIF(stropt="",1,stropt)
				
	if cint(stropt)=1 then
		strQ=Request.Form("lstquater")
		'Default is current quarter
		strQ=IIF(strQ="",arrFirstMonthOfQuater(datepart("q",Date())-1),strQ)
		strYear=Request.Form("lstyear")
		'Default is current year
		strYear=IIF(strYear="",year(date()),strYear)
		
		strfrom=cdate("1-" & strQ & "-" & strYear)
		strTo=DateAdd("q",1,strfrom)-1
	else
		strfrom=cdate(ConvertTommddyyyy(Request.Form("txtfrom")))
		strTo=cdate(ConvertTommddyyyy(Request.Form("txtto")))
	end if
		
End sub
'*********************************************************
'Generate report
'*********************************************************
Function ATSsql(byval strF,byval strT)
	dim strATS
	If year(strF)<> year(strT) then
		strATS="(SELECT AssignmentID,Hours,OverTime,Tdate FROM "
		For ii=year(strF) To year(strT)
			strATS=strATS & selectTable(ii)
			If ii<>	year(strT) then
				strATS=strATS & " UNION ALL SELECT AssignmentID,Hours,OverTime,Tdate FROM "
			else
				strATS=strATS & ")"
			end if
		Next
	else
		strATS=selectTable(year(strT))
	end if
	ATSsql=strATS
End function
'*********************************************************
'Generate report
'*********************************************************
Function GenerateReport(strF,strT)
	dim strSql, strATS,strReturn,arrlongmon
	dim rsPro,rsStaff,rsHour,rsExchange, strProID,strSID,ii
		
	arrlongmon  = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	arrContract  = Array("NC", "CP", "CS")

	strProID=""
	strSID=0
	strReturn=""
	'??????
	strATS=ATSsql(Cdate("1/1/2005"),Date())
	strQuery=""
	if not fgViewAll then strQuery = strQuery & " AND " & getWherePhase("d",session("USERID"))
	
	'For all project that have hours booking on Timesheet
	strSql= "SELECT Pro.ProjectID,Pro.ProjectName,ISNULL(Pro.[Value],0) as Value,CSOCompleted,CSOMainHours,ManagerID,signContract,ISNULL(Hours,0) as Hours, " & _
				"DailyRate,ExchangeRate, EstRemaining,CWFValue, isnull(ProInvoice.InvoiceValue,0) as InvoiceValue,isnull(ProInvoice.InvoiceValueUSD,0) as InvoiceValueUSD, Currency FROM ATC_Projects Pro INNER JOIN " &_
				"ATC_Countries ON ATC_Countries.CountryCode=SUBSTRING(pro.ProjectID,12,2) " & _
				"LEFT JOIN (SELECT SUM((a.Hours + a.OverTime)) as Hours,d.ProjectID FROM " & strATS & " a " &_
					"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " &_
					"INNER JOIN ATC_Tasks c ON b.SubtaskID=c.SubtaskID " &_
					"INNER JOIN ATC_Projects d ON c.ProjectID=d.ProjectID " &_
					"WHERE LEFT(d.ProjectID,3)<>'ATL' " & strQuery &_
					"GROUP BY d.ProjectID) Timesheet ON Pro.ProjectID=Timesheet.ProjectID LEFT JOIN " & _
					"(SELECT ISNULL(SUM(InvoiceValue),0) as InvoiceValue,ISNULL(SUM(InvoiceValue * ExchangeRate),0) as InvoiceValueUSD,projectID FROM ATC_ProjectInvoices GROUP BY projectID) ProInvoice ON Pro.ProjectID=ProInvoice.ProjectID " & _
					"ORDER BY Pro.ProjectID" 		
'Response.Write strSql						
	'if not fgViewAll then strQuery = strQuery & " AND " & getWherePhase("ATC_Projects",session("USERID"))
'Response.Write strQuery
	
	Call GetRecordset(strSql,rsPro)
	'For all project that have hours booking on Timesheet in period
	strATS=ATSsql(strF,strT)
	strSql= "SELECT ISNULL(hours,0) as Hours, isNULL(ATS.ProjectID,Invoice.ProjectID) as ProjectID FROM " & _
			"(SELECT SUM((a.Hours + a.OverTime)) as Hours,d.ProjectID FROM " & strATS & " a " &_
			"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " &_
					"INNER JOIN ATC_TAsks c ON b.SubtaskID=c.SubtaskID " &_
					"INNER JOIN ATC_Projects d ON c.ProjectID=d.ProjectID " &_
					"WHERE LEFT(d.ProjectID,3)<>'ATL' " & strQuery & " AND a.AssignmentID>1  AND Tdate Between '" & cdate(strF) & "' AND '" & strT & "' GROUP BY d.ProjectID) as ATS " &_
			"FULL OUTER JOIN (SELECT DISTINCT ProjectID FROM ATC_ProjectInvoices " & _
								"WHERE InvoiceDate BETWEEN '" & cdate(strF) & "' AND '" & strT & "') as Invoice " & _
			"ON Invoice.ProjectID=ATS.ProjectID " & _
					IIf(strSearchText<>""," WHERE ATS.ProjectID like '%" & trim(strSearchText) & "%' OR Invoice.ProjectID like '%" & trim(strSearchText) & "%'","") & _
					" ORDER BY ATS.ProjectID"	
	
'Response.Write "<br>" &  strSql	
'Response.End
	Call GetRecordset(strSql,rsHour)
'Response.Write "<br>" &  rsHour.RecordCount	

	strSql="SELECT UserID,UserName FROM ATC_Users a INNER JOIN ATC_PersonalInfo b ON a.UserID=b.PersonID WHERE fgDelete=0 AND UserType=1"	
	Call GetRecordset(strSql,rsStaff)
	
	strSql="SELECT Currency ,RateToUSD FROM ATC_ExchangeRate a WHERE DateApply =(SELECT MAX(DateApply) FROM ATC_ExchangeRate b WHERE a.Currency=b.Currency)"
	Call GetRecordset(strSql,rsExchange)
	
	strSql="SELECT ISNULL(SUM(InvoiceValue),0) as InvoiceValue,ISNULL(SUM(InvoiceValue * ExchangeRate),0) as InvoiceValueUSD,projectID FROM ATC_ProjectInvoices WHERE InvoiceDate BETWEEN '" & strF & "' AND '" & strT & "' GROUP BY projectID"
	Call GetRecordset(strSql,rsInvoice)



	if gMessage="" then
		if rsHour.RecordCount>0 then
				rsHour.MoveFirst
				
				Do while not rsHour.EOF
					rsPro.Filter="ProjectID='" & rsHour("ProjectID") & "'"

					rsStaff.Filter="UserID=" & rsPro("ManagerID")
					
					strBgcolor="#FFFFFF"
					if rsPro("CSOCompleted") then strBgcolor="#F1DADB"
					strReturn=strReturn & "<tr valign='top' bgcolor='" & strBgcolor & "' class='blue-normal'> "
					
					'For AM
					if not rsStaff.EOF Then
						strReturn=strReturn & "<td>" & rsStaff("UserName") & "</td>"
					else
						strReturn=strReturn & "<td>--</td>"
					end if
					
					'For Project APK
					strTemp=ParseAPK(rsHour("ProjectID"))

					strReturn=strReturn & "<td>" & strTemp(0) & strTemp(1) & "</td>"
					strReturn=strReturn & "<td>" & strTemp(2) & "</td>"					
					strReturn=strReturn & "<td>" & strTemp(3) & "</td>"
					
					'For Project Description
					strReturn=strReturn & "<td>" & rsPro("ProjectName") & "</td>"
					
					'For Daily Rate
					strReturn=strReturn & "<td>" & IIF(strTemp(3)="T",rsPro("DailyRate"),"--") & "</td>"					
					
					'For CSO
					strReturn=strReturn & "<td>" & IIF(isnull(rsPro("CSOMainHours")),"--",rsPro("CSOMainHours")) & "</td>"
					
					dblCSOValue=0
					dblCSORate=0
					dblExRate=1
					if not isnull(rsPro("ExchangeRate")) then dblExRate=cdbl(rsPro("ExchangeRate"))
					
					if strTemp(3)="T" then
						if not isnull(rsPro("DailyRate")) then 							
							dblCSOValue= FormatNumber((cdbl(rsPro("DailyRate")) * cdbl(rsPro("Hours")))/8,2)
							dblCSORate=FormatNumber(cdbl(rsPro("DailyRate"))/8,2)
							
						end if
					else
						if not isnull(rsPro("Value")) then 
							dblCSOValue=rsPro("Value")
							if not isnull(rsPro("CSOMainHours")) then dblCSORate=FormatNumber((cdbl(rsPro("Value")))/cdbl(rsPro("CSOMainHours")),2)
						end if
						
					end if
					strReturn=strReturn & "<td>" & IIF(cdbl(dblCSOValue)=0,"",dblCSOValue) & "</td>"														
					strReturn=strReturn & "<td>" & IIF(dblCSORate=0,"", dblCSORate) & "</td>"
					
					'For CWF Value
					strReturn=strReturn & "<td>" & rsPro("CWFValue") & "</td>"
					
					'For Contract Status
					strReturn=strReturn & "<td>" & arrContract(rsPro("signContract")) & "</td>"
					
					'For Actual Value
					strReturn=strReturn & "<td>" & rsPro("Hours") & "</td>"
					
					'For Period
					strReturn=strReturn & "<td>" & rsHour("Hours") & "</td>"
					
					'For Estimated Remaining
					strReturn=strReturn & "<td>" & rsPro("EstRemaining") & "</td>"
					
					'For To Completion
					dblEstRemain=IIF(isnull(rsPro("EstRemaining")),0, rsPro("EstRemaining"))
					strReturn=strReturn & "<td>" & formatnumber(cdbl(dblEstRemain) + cdbl(rsPro("Hours")),2) & "</td>"
					
					'For ATC/CSO
					dblActualOverCSO="--"
					if not isnull(rsPro("CSOMainHours")) then dblActualOverCSO=formatnumber((cdbl(rsPro("Hours")) - cdbl(rsPro("CSOMainHours"))) * 100 /cdbl(rsPro("CSOMainHours")),2)
					strReturn=strReturn & "<td>" & dblActualOverCSO & "</td>"
					
					'For Project Completion
					if cdbl(dblEstRemain) + cdbl(rsPro("Hours"))>0 then
						strReturn=strReturn & "<td>" & formatnumber(cdbl(rsPro("Hours")) * 100/(cdbl(dblEstRemain) + cdbl(rsPro("Hours"))),2) & "</td>"
					else
						strReturn=strReturn & "<td>--</td>"
					end if
					
					'For Actual Total
					'dblActuaRate=cdbl(rsPro("InvoiceValueUSD"))/cdbl(rsPro("Hours"))
					rsInvoice.Filter="ProjectID='" & rsHour("ProjectID") & "'"
					dblInvoicePer=""
					if rsInvoice.RecordCount>0 then
						dblInvoicePer=FormatNumber(cdbl(rsInvoice("InvoiceValueUSD")),2)
						if cdbl(rsPro("Hours"))>0 then
							dblActuaRate=cdbl(rsInvoice("InvoiceValueUSD"))/cdbl(rsPro("Hours"))
							strReturn=strReturn & "<td>" & FormatNumber(dblActuaRate,2) & "</td>"
						else
							strReturn=strReturn & "<td></td>"
						end if
					else
						strReturn=strReturn & "<td></td>"
					end if
					rsInvoice.Filter=""
					'For Ext. Earning
					dblExtEarning=""
					dblExchangeDefault=1
					
					rsExchange.filter="Currency='" & rsPro("Currency") & "'"
					if not rsExchange.EOF then	dblExchangeDefault=cdbl(rsExchange("RateToUSD"))					
					'
					if not isnull(rsPro("CWFValue")) and cdbl(dblEstRemain) + cdbl(rsPro("Hours"))>0 then 
						dblTemp=cdbl(rsPro("CWFValue")) - (cdbl(rsPro("InvoiceValue")) * dblExchangeDefault)
						dblExtEarning=formatnumber((cdbl(rsPro("InvoiceValueUSD")) + dblTemp)/(cdbl(dblEstRemain) + cdbl(rsPro("Hours"))),2)					
					end if
					strReturn=strReturn & "<td>" & dblExtEarning & "</td>"
					
					'for Pot. Earning
					dblEarning=0
					if dblCSOValue<>"" and cdbl(dblEstRemain) + cdbl(rsPro("Hours"))>0 then dblEarning=formatnumber((cdbl(dblCSOValue) * cdbl(rsPro("Hours")))/(cdbl(dblEstRemain) + cdbl(rsPro("Hours"))),0)
					strReturn=strReturn & "<td>" & IIF(dblEarning<>0,dblEarning,"") & "</td>"
					'Invoice Per.
					strReturn=strReturn & "<td>" & dblInvoicePer & "</td>"
					'for Invoice To-Date($)
					dblInvoice=0
					if cdbl(rsPro("InvoiceValueUSD"))<>0 then dblInvoice=FormatNumber(cdbl(rsPro("InvoiceValueUSD")),2)
					strReturn=strReturn & "<td>" & IIF(dblInvoice<>0,dblInvoice,"") & "</td>"
					
					'for Over/Under(%)
					If(dblEarning<>0) then
						strReturn=strReturn & "<td>" & formatnumber((dblInvoice-dblEarning)*100/dblEarning,2) & "</td>"
					else
						strReturn=strReturn & "<td>" & "" & "</td>"
					end if
					
					'for Invoice To-Date(%)				
					if cdbl(dblCSOValue)<>0 and cdbl(rsPro("InvoiceValueUSD"))<>0 then
						strReturn=strReturn & "<td>" & formatnumber(cdbl(rsPro("InvoiceValueUSD")) * 100/ cdbl(dblCSOValue),2)  & "</td>"
					else
						strReturn=strReturn & "<td></td>"
					end if

					rsStaff.Filter=""
					rsPro.Filter=""
					rsHour.MoveNext
					strReturn=strReturn & "</tr>"
				Loop		
		else
			gMessage = "There is no data for this period."
		end if
	end if
	GenerateReport=strReturn
End Function
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
'--------------------------------------------------
' Check VIEWALL project right
' User can update all project
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgViewAll = False
	Else
		varGetRight = session("RightOn")
		fgViewAll = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "View all projects" Then
				fgViewAll = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If		
'----------------------------------
' Get report
'----------------------------------	
dim arrFirstMonthOfQuater
Dim intDate,intMonth,intYear,arrlstDay(2),strprintdate,strfromto,strfrom,strTo
Dim gMessage,intnumMonths,stropt,strQ,strYear,strTypeStatus,strSearchText
arrFirstMonthOfQuater  = Array("Jan","Apr","Jun","Oct")

call Getperiod()
strSearchText=Request.Form("txtsearch")

strprintdate=FormatDateTime(cdate(month(date) & "/" & Day(date) & "/" & year(date)),1)
strfromto = "From " & ddmmyyyy(strfrom) & " to " & ddmmyyyy(strTo)

strLast = GenerateReport(strFrom,strTo)
if not isEmpty(session("rpt_forecast")) then session("rpt_forecast")=empty

session("rpt_KPI")=strLast
'----------------------------------
' Get Company Information
'----------------------------------
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

<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
<script>
var objWindowSumPro;

function _print() { //v2.0
var str1 = "<%=strfromto%>";
	str1 = escape(str1);
var str2 = "<%=strprintdate%>";
	str2 = escape(str2);
var fgprint = 10;
if (fgprint!=0) {
	window.status = "";
	strFeatures = "top="+(screen.height/2-275)+",left="+(screen.width/2-390)+",width=800,height=550,toolbar=no," 
	            + "menubar=yes,location=no,resizable=yes,directories=no,scrollbars=yes,status=yes";
	if ((objWindowSumPro) && (!objWindowSumPro.closed)) {
		objWindowSumPro.focus();
	
	} else {
		objWindowSumPro = window.open("p_kpisumtimesheet.asp?fromto=" + str1 + "&printdate=" + str2 + "&num=" + 7, "MyNewWindow", strFeatures);
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
			alert("Please enter end date before click here.")
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
		document.frmreport.action = "KPI_sumtimesheet.asp?act=REFRESH";
		document.frmreport.target = "_self" ;
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

		<table width="780" border="0" cellspacing="0" cellpadding="0" align="center" >
	<tr>
		<td>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="30%" valign="bottom"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td valign="top" width="0%" class="blue" height="33">&nbsp;</td>
                  <td width="4%" class="blue" align="right" height="33"> 
                    <input type="radio" name="opttime" value="0" <%=IIf(cint(stropt)=0,"checked","")%> >
                  </td>
                  <td class="blue-normal" width="8%" height="33">From </td>
                  <td class="blue-normal" width="40%" height="33"> 
                    <input type="text" name="txtfrom" class="blue-normal"  size="5" style="width:60" value="<%=iif(cint(stropt)=0,ddmmyyyy(strfrom),"")%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                  <td width="8%" class="blue-normal" height="33"> To </td>
                  <td class="blue-normal" height="33" width="40%"> 
                    <input type="text" name="txtto" class="blue-normal" size="5" style="width:60" value="<%=iif(cint(stropt)=0,ddmmyyyy(strTo),"")%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                </tr>
                <tr> 
                  <td valign="top" width="0%" class="blue">&nbsp;</td>
                  <td width="4%" class="blue" align="right"> 
                    <input type="radio" name="opttime" value="1" <%=IIf(cint(stropt)=1,"checked","")%>>
                  </td>
                  <td class="blue-normal" width="8%"></td>
                  <td class="blue-normal" width="40%"> 
						<select name='lstquater' size='1' height='26px' width='75px' style='width:75px;height=24px;' class='blue-normal' onClick='document.frmreport.opttime[1].checked=true;'>
						<%For ii=0 to 3 %>
							<option value='<%=arrFirstMonthOfQuater(ii)%>' <%=IIf(strQ=arrFirstMonthOfQuater(ii),"selected","")%>><%="Q" & (ii+1)%></option>					
						<%next%>
						</select>
                  </td>
                  <td width="8%" class="blue-normal"> Year </td>
                  <td class="blue-normal" width="40%" >
					<select name='lstyear' size='1' height='26px' width='70px' style='width:70px;height=24px;' class='blue-normal' onClick='document.frmreport.opttime[1].checked=true;'>
					<%For ii=Year(date()) to 2005 step -1%>
						<option value='<%=ii%>' <%=IIf(cint(strYear)=ii,"selected","")%>><%=ii%></option>					
					<%next%>
					</select>
                  </td>
                </tr>
              </table>
            </td>
            <td width="35%"  valign="bottom"> 
			<table width="100%" border="0" cellpadding="1" cellspacing="0">
                <tr> 
                  <td class="blue-normal" width="30%" height="33" valign="bottom">Search for</td>
                  <td class="blue-normal" width="70%" valign="bottom">
					<input type="text" name="txtsearch" class="blue-normal" size="10" style="width:95%;" value="">
				  </td>
                </tr>
                <tr> 
                  <td class="blue-normal" >Project Status</td>
                  <td class="blue-normal" valign="bottom">
					<select name="lsttypepro" style="width:95%" class="blue-normal">
						<option value="0" selected>&nbsp;</option>
						<option value="1" >Completed</option>
						<option value="2" >In Progress</option>
					</select> </td>
                  
                </tr>                
              </table>
             
            </td>
            <td width="35%" align="left" valign="bottom"> 
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
		</td>
	</tr>
	<tr>
		<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="FFFFFF">
			<td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=IIF(gMessage<>"",gMessage,"")%></b></td>
		  </tr>
          <tr> 
            <td bgcolor="8CA0D1"><img src="../IMAGES/DOT-01.GIF" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>&nbsp; </td>
          </tr>
        </table>
		</td>
	</tr>
	</table>

  <table width="90%" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%" align="center"  >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top">
      
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td ><img src="../IMAGES/dot1px.gif" width="1" height="10"></td>
          </tr>
        </table>	
        
  <%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			'arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "KPI Report")
			
			'arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strfromto)
			'arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			'Response.Write(arrPageTemplate(1))
			
%>


<table width="100%" border="0" cellspacing="0" cellpadding="0">
           <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="5">
				 <tr>
                  <td rowspan="2" align="center" valign="bottom" bgcolor="#E7EBF5" class="blue">AM</td>
                  <td colspan="3" align="center" bgcolor="#E7EBF5" class="blue">ProjectID</td>
                  <td rowspan="2" align="center" bgcolor="#E7EBF5" class="blue">Description</td>
                  <td rowspan="2" align="center" bgcolor="#E7EBF5" class="blue">Daily Rate<br>($/hrs)</td>
                  <td colspan="3" align="center" bgcolor="#E7EBF5" class="blue">CSO</td>
                  <td rowspan="2" align="center" bgcolor="#E7EBF5" class="blue">CWF Value (Orig Cur.)</td>
                  <td rowspan="2" align="center" bgcolor="#E7EBF5" class="blue">Contract Status</td>
                  <td colspan="4" align="center" bgcolor="#E7EBF5" class="blue">Actual </td>
                  <td colspan="8" align="center" bgcolor="#E7EBF5" class="blue">Calculated</td>
                  <td rowspan="2" align="center" bgcolor="#E7EBF5" class="blue">Invoiced To Date <br>(%)</td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#E7EBF5" class="blue">APK</td> 
                  <td align="center" bgcolor="#E7EBF5" class="blue">VO</td>
                  <td align="center" bgcolor="#E7EBF5" class="blue">Type<br>(TC<BR>/LS)</td>
                  <td align="center" bgcolor="#E7EBF5" class="blue">Hours</td>
                  <td align="center" bgcolor="#E7EBF5" class="blue"> Value <br>(USD)</td>

                  <td align="center" bgcolor="#E7EBF5" class="blue">Rate<br>    (USD/hrs)</td>
                  <td align="center" bgcolor="#E7EBF5" class="blue">To-date<br>    (hrs) </td>
                  <td align="center" bgcolor="#E7EBF5" class="blue">Period<br>    (hrs) </td>
                  <td align="center" bgcolor="#E7EBF5" class="blue">Estimated Remaining<br>    (hrs)</td>
                  <td align="center" bgcolor="#E7EBF5" class="blue">To Completion<br>    (hrs)</td>
				  <td align="center" bgcolor="#E7EBF5" class="blue">ATC/CSO<br>(%)</td>
				  <td align="center" bgcolor="#E7EBF5" class="blue">Project Completion<br>(%)</td>	
				  
				  <td align="center" bgcolor="#E7EBF5" class="blue">Actual Total<br>(USD/hrs)</td>	
				  <td align="center" bgcolor="#E7EBF5" class="blue">Expected. Earning<br>(USD/hrs)</td>
  				  <td align="center" bgcolor="#E7EBF5" class="blue">Pot. Earning<br>($)</td>
  				  <td align="center" bgcolor="#E7EBF5" class="blue">Invoice Per.<br>($)</td>
   				  <td align="center" bgcolor="#E7EBF5" class="blue">Invoice to-date<br>($)</td>
				  <td align="center" bgcolor="#E7EBF5" class="blue">Over/Under<br>(%)</td>
			    </tr>
			    
			    
<%=strLast%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <p>
<%			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
%>
</form>
</body>
</html>