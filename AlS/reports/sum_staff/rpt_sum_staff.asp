<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	Response.Buffer = True

	Dim intUserID, intMonth, intYear, intRow, intNum, intDepartmentID, intCompanyID, intRows, intPageSize, intCurPage, intCurRow, intCount, intTotalPage
	Dim varFullName, varFrom, varTo, varDepartment, varStaff 
	Dim strTitle, strFunction, strType, strTitle2, strFrom, strTo
	Dim rsJobtitle, rsIDnumber,rsReportTo


'--------------------------------------------------
'Get data from Timesheet
'--------------------------------------------------
Function GetSumWorkHoursRecordset(strFullname,dblDepartmentID,dateStart,dateEnd)

	dim strConnect,objDatabase
	dim rs
	
	strConnect = Application("g_strConnect")
	
	rs=null
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then
		objDatabase.cnDatabase.CursorLocation=adUseClient 
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "SumOfStaffHours"
		
		Set myParam = myCmd.CreateParameter("fullname",adVarChar,adParamInput,200)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("department",adInteger,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("start_date",adDate,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("finish_date",adDate,adParamInput)
		myCmd.Parameters.Append myParam


		myCmd("fullname") = strFullname
		myCmd("department")=dblDepartmentID
		myCmd("start_date")=dateStart
		myCmd("finish_date")=dateEnd

		set rs=myCmd.Execute		
	end if

	set GetSumWorkHoursRecordset = rs
end function
'--------------------------------------------------
'
'--------------------------------------------------
Function GetReportTo(intStaffID)
	dim strReport

	strsql="SELECT PersonID,FirstNameLeader + ' ' + LastnameLeader FROM HR_Employee"


	GetReportTo=strReport
end function
'--------------------------------------------------
'
'--------------------------------------------------
Sub ResetArrayByInt(byref arr)
	for i = 0 to UBound(arr)
		arr(i)=0
	next
end sub

'--------------------------------------------------
'
'--------------------------------------------------
Function GenerateSubtotalRow(strLable,dblTotal)
	dim strSubtotal
	
	if cdbl(dblTotal(UBound(dblTotal)-2))>0 then dblTotal(UBound(dblTotal))=cdbl(dblTotal(0)+dblTotal(2))/cdbl(dblTotal(UBound(dblTotal)-2))
	
	
	
	strSubtotal= "<tr bgcolor='#E7EBF5'>" & _
		            "<td valign='top' colspan='5' class='blue' align='right'>" & strLable & "</td>"
	for i=0 to UBound(dblTotal)-1
		strSubtotal= strSubtotal & "<td valign='top' class='blue' align='right'>" & FormatNumber(dblTotal(i),2)& "</td>" 
	next
	
	strSubtotal= strSubtotal & "<td valign='top' class='blue' align='right'>" & FormatPercent(dblTotal(UBound(dblTotal)),2) & "</td>" & _
						"</tr>"
	
	GenerateSubtotalRow=strSubtotal
end function
'--------------------------------------------------
'
'--------------------------------------------------
Function GenerateSumReport(rs)
	dim strReport,intNo,dblTotal,dblTotalAvailable
	dim dblSubtotal(15),dblGrandTotal(15)
	dim intIndirect
	strReport=""

	call ResetArrayByInt(dblSubtotal)
	call ResetArrayByInt(dblGrandTotal)
	'if not rs.EOF then
		rs.MoveFirst
		'rs.AbsolutePage=intcurPage
		intNo=1	'+ (rs.AbsolutePage - 1) * rs.PageSize
		intIndirect=0
		intI=0

	'Response.End
	
	
		do while not rs.EOF 'and intI < rs.PageSize
'Response.write rs("Fullname") & "<br>"
		
			if rs("fgIndirect")<>intIndirect then
				if strReport<>"" then strReport=strReport & GenerateSubtotalRow("Sub Total of Direct Staff",dblSubtotal)
				intIndirect=rs("fgIndirect")				
				call ResetArrayByInt(dblSubtotal)
			end if

			dblRate=0
			dblTotalAvailable=0
			dblTotal=0
			for i = 5 to rs.Fields.Count-1
			
				dblTotal = cdbl(dblTotal) + cdbl(rs.Fields(i))
				
				if (i<11 AND i<>7 ) then dblTotalAvailable=dblTotalAvailable + cdbl(rs.Fields(i))
				
				dblSubtotal(i-5)=cdbl(dblSubtotal(i-5)) + cdbl(rs.Fields(i))
				dblGrandTotal(i-5)=	dblGrandTotal(i-5) + cdbl(rs.Fields(i))
				
			next
			dblSubtotal(i-5)=dblSubtotal(i-5) + dblTotalAvailable
			dblGrandTotal(i-5) = dblGrandTotal(i-5) + dblTotalAvailable

			dblSubtotal(i-4)=dblSubtotal(i-4) + dblTotal
			dblGrandTotal(i-4) = dblGrandTotal(i-4) + dblTotal

			if dblTotalAvailable>0 then dblRate=(cdbl(rs("ClientHoursBill"))+cdbl(rs("ClientOTHours")))/dblTotalAvailable

			strJobTitle=""
			rsJobtitle.MoveFirst
			rsJobtitle.Find "StaffID=" & rs("StaffID")
			if not rsJobtitle.EOf then strJobTitle=rsJobtitle("JobTitle")

			strLeader=""
			rsReportTo.MoveFirst
			rsReportTo.Find "PersonID=" & rs("StaffID")
			if not rsReportTo.EOf then strLeader=rsReportTo("LeaderName")

			strClientHoursBill=	FormatNumber(rs("ClientHoursBill"),2)		
			if 	cdbl(rs("ClientHoursBill"))>0 then strClientHoursBill = "<a href='javascript:viewdetail(" & rs("StaffID") & ",0)'>" & FormatNumber(rs("ClientHoursBill"),2) & "</a>"
			
			strClientHoursNoBill=	FormatNumber(rs("ClientHoursNoBill"),2)		
			if 	cdbl(rs("ClientHoursNoBill"))>0 then strClientHoursNoBill = "<a href='javascript:viewdetail(" & rs("StaffID") & ",1)'>" & FormatNumber(rs("ClientHoursNoBill"),2) & "</a>"
			
			strClientOTHours=	FormatNumber(rs("ClientOTHours"),2)		
			if 	cdbl(rs("ClientOTHours"))>0 then strClientOTHours = "<a href='javascript:viewdetail(" & rs("StaffID") & ",2)'>" & FormatNumber(rs("ClientOTHours"),2) & "</a>"
			
			strATLHours=	FormatNumber(rs("ATLHours"),2)		
			if 	cdbl(rs("ATLHours"))>0 then strATLHours = "<a href='javascript:viewdetail(" & rs("StaffID") & ",3)'>" & FormatNumber(rs("ATLHours"),2) & "</a>"
			
			strDowntimeHours=	FormatNumber(rs("DowntimeHours"),2)		
			if 	cdbl(rs("DowntimeHours"))>0 then strDowntimeHours = "<a href='javascript:viewdetail(" & rs("StaffID") & ",4)'>" & FormatNumber(rs("DowntimeHours"),2) & "</a>"
			
			strReport= strReport & "<tr bgcolor='#FFFFFF'>" & _
		            "<td valign='top' class='blue' align='right'>" & intNo & "</td>" & _
		            "<td valign='top' class='blue' >" & rs("IDNumber") & "</td>" & _
		            "<td valign='top' class='blue-normal' >" & rs("Fullname") & "</td>" & _
		            "<td valign='top' class='blue-normal' >" & strJobTitle & "</td>" & _
		            "<td valign='top' class='blue-normal' >" & strLeader & "</td>" & _
		            "<td valign='top' bgcolor='#D2DAEC' class='blue-normal'align='right'>" & strClientHoursBill & "</td>" & _
		            "<td valign='top' bgcolor='#D2DAEC' class='blue-normal'align='right'>" & strClientHoursNoBill & "</td>" & _
		            "<td valign='top' bgcolor='#D2DAEC' class='blue-normal'align='right'>" & strClientOTHours & "</td>" & _
		            "<td valign='top' class='blue-normal'align='right'>" & strATLHours & "</td>" & _
		            "<td valign='top' class='blue-normal' align='right'>" & FormatNumber(rs("GA"),2)& "</td>" & _
		            "<td valign='top' bgcolor='#FFE1E1' class='blue-normal' align='right'>" & strDowntimeHours & "</td>" & _
		            "<td valign='top' class='blue-normal' align='right'>" & FormatNumber(rs("PD"),2)& "</td>" & _
		            "<td valign='top' bgcolor='#FFF2F2' class='blue-normal' align='right'>" & FormatNumber(rs("PH"),2)& "</td>" & _
		            "<td valign='top' bgcolor='#FFF2F2' class='blue-normal' align='right'>" & FormatNumber(rs("AH"),2)& "</td>" & _
		            "<td valign='top' bgcolor='#FFF2F2' class='blue-normal' align='right'>" & FormatNumber(rs("SL"),2)& "</td>" & _
		            "<td valign='top' bgcolor='#FFF2F2' class='blue-normal' align='right'>" & FormatNumber(rs("OL"),2)& "</td>" & _
					"<td valign='top' bgcolor='#FFF2F2' class='blue-normal' align='right'>" & FormatNumber(rs("OffOT"),2)& "</td>" & _
		            "<td valign='top' bgcolor='#FFF2F2' class='blue-normal' align='right'>" & FormatNumber(rs("UL"),2)& "</td>" & _
		            "<td valign='top' class='blue' align='right'>" & FormatNumber(dblTotalAvailable,2)& "</td>" & _
		            "<td valign='top' class='blue' align='right'>" & FormatNumber(dblTotal,2)& "</td>" & _
		            "<td valign='top' class='blue' align='right'>" & FormatPercent(dblRate,2) & "</td>" & _
				"</tr>"
			intNo=intNo+1
			rs.MoveNext
			intI=intI+1
		loop
		
		strReport=strReport & GenerateSubtotalRow("Sub Total of Indirect Staff",dblSubtotal)
		strReport=strReport & GenerateSubtotalRow("Grand Total ",dblGrandTotal)

	'end if
	GenerateSumReport=strReport
End Function

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	intNum = -1
	intRow = -1
	
	intCompanyID = session("InHouse")
	strName=""
	if Request.Form("txtname")<>"" then strName = Request.Form("txtname")
	intDepartmentID=0
	if Request.Form("lbdepartment")<>"" then intDepartmentID = Request.Form("lbdepartment")
	
	strType = "M"
	IF Request.Form("rdotype")<>"" then strType = Request.Form("rdotype")
		
	If strType = "M" Then
	
		intMonth = Month(Date)		
		If Request.Form("lbmonth") <> "" Then intMonth = CInt(Request.Form("lbmonth"))
				
		intYear = Year(Date)
		If Request.Form("lbyear") <> "" Then intYear	= Request.Form("lbyear")
		
		strTitle2	= SayMonth(intMonth) & " - " & intYear
		
		strFrom=Cdate(intMonth & "/1/" & intYear)	
		strTo=DateAdd("m",1,strFrom)-1
	Else
	
		strFrom1	= Request.Form("txtfrom")
		strTo1		= Request.Form("txtto")
		strTitle2	= "From " & strFrom1 & " To " & strTo1
	
		varFrom		= split(strFrom1,"/")
		strFrom		= CDate(varFrom(1) & "/" & varFrom(0) & "/" & varFrom(2))
		varTo		= split(strTo1,"/")
		strTo		= CDate(varTo(1) & "/" & varTo(0) & "/" & varTo(2))			
	End If
		

'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

		If Not isEmpty(session("Preferences")) Then
			arrPre = session("Preferences")
			If arrPre(1, 0) > 0 Then intPageSize = arrPre(1, 0) Else intPageSize = 9
			Set arrPre = Nothing
		Else
			intPageSize = 9
		End If			

'--------------------------------------------------
' Get current page
'--------------------------------------------------
	
	intCurPage = trim(Request.Form("P"))
	
	If intCurPage = "" Or Request.QueryString("act") = "vra" Then
		intCurPage = 1
	End If	

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID	= session("USERID")
'--------------------------------------------------
' Initialize department array
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	strSql="SELECT personID,IDNumber FROM ATC_PersonalInfo WHERE fgDelete=0 AND userType=1 ORDER BY personID"	
	If objDatabase.dbConnect(strConnect) Then	Call GetRecordset(strSql,rsIDNumber)

	strSql="SELECT StaffID,JobTitle FROM ATC_Employees A INNER JOIN ATC_Jobtitle B ON A.JobtitleID=B.JobTitleID ORDER BY StaffID"	
	strSql="SELECT StaffID,JobTitle FROM [HR_CurrentJobtitle] ORDER BY StaffID"	
	If objDatabase.dbConnect(strConnect) Then	Call GetRecordset(strSql,rsJobtitle)

	
	strsql="SELECT PersonID,(FirstNameLeader + ' ' + LastnameLeader) as LeaderName FROM HR_Employee  ORDER BY PersonID"
	If objDatabase.dbConnect(strConnect) Then	Call GetRecordset(strSql,rsReportTo)


	strSql="SELECT DepartmentID, Department, fgActivate FROM  ATC_Department WHERE  (fgActivate = 1) ORDER BY Department"	
	Call GetRecordset(strSql,rsDepart)
	strDepartment= PopulateDataToListWithoutSelectTag(rsDepart,"DepartmentID", "Department",-1)


'--------------------------------------------------
' End Of initializing department array
'--------------------------------------------------

'--------------------------------------------------
' Analyse query and prepare report
'--------------------------------------------------

	If Request.QueryString("act") = "" Or Request.QueryString("act") = "vra" Then
	
		set rsSumHours=	GetSumWorkHoursRecordset(strName,intDepartmentID,strFrom,strTo)
		if rsSumHours.recordcount>0 then
			
			rsSumHours.PageSize = intPageSize    ' So ban ghi tren mot trang
			intTotalPage = rsSumHours.PageCount       ' Tong so trang						
			
			If Not IsEmpty(session("varSumHours")) Then session("varSumHours") = Empty
			set session("varSumHours") = rsSumHours
		end if	
	Else
		
		set rsSumHours=session("varSumHours")
		
		If Request.QueryString("act") = "vpa1" Then
			intCurPage = Request.Form("txtpage")
		ElseIf Request.QueryString("act") = "vpa2" Then
			intCurPage = CInt(intCurPage) - 1
		ElseIf Request.QueryString("act") = "vpa3" Then	
			intCurPage = CInt(intCurPage) + 1
		End If	
	End If
	
'--------------------------------------------------
' Generate Report
'--------------------------------------------------	
	
	strLast=GenerateSumReport(rsSumHours)
	
	Session("StrLast")=strLast
		
	Set objDatabase = Nothing
'--------------------------------------------------
' End of preparing report
'--------------------------------------------------

'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))

	strFunction = "<a class='c' href='../../welcome.asp?choose_menu=B'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  strtmp1 & "&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage();'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' title='Log Out'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#'>Help</a>&nbsp;&nbsp;&nbsp;"


	If isEmpty(session("arrInfoCompany")) Then
		strConnect = Application("g_strConnect") 
		Set objDb = New clsDatabase
		If objDb.dbConnect(strConnect) Then
			strQuery = "select a.CompanyName, isnull(Address,'') Address, isnull(City,'') City, isnull(b.CountryName,'') Country, " &_
						"isnull(Phone,'') Phone, isnull(Fax,'') Fax, isnull(c.Logo,'') Logo from ATC_Companies a " &_
						"left join ATC_Countries b On a.CountryID = b.CountryID " &_
						"left join ATC_CompanyProfile c ON a.CompanyID = c.CompanyID " &_
						"where a.CompanyID = " & session("Inhouse")
			If objDb.runQuery(strQuery) Then
				If not objDb.noRecord Then
					arrInfoCompany = objDb.rsElement.getRows
					session("arrInfoCompany") = arrInfoCompany
					objDb.closerec
				End If
			Else
				strError = objDb.strMessage
			End If
			objDb.dbDisconnect
		Else
			strError = objDb.strMessage
		End If
		Set objDb = Nothing
	End If

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_report.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If Not isEmpty(session("arrInfoCompany")) Then
	arrTmp = session("arrInfoCompany")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", arrTmp(0, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", arrTmp(1, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", arrTmp(2, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", arrTmp(3, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", arrTmp(4, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", arrTmp(5, 0))
	If arrTmp(6, 0)<>"" Then
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "<img src='../../images/" & arrTmp(6, 0) & "' border='0'>" )
	Else
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;" )
	End If
	Set arrTmp = Nothing
Else
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;")
End If


%>

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>

<link href="../../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>
<script src="../../jQuery/jquery.min.js"></script>
<script src="../../jQuery/jquery-ui.min.js"></script>
<link href="../../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<script>
	$(function() {
		var dates = $( "#txtfrom, #txtto" ).datepicker({
		dateFormat: "dd/mm/yy",
		minDate:"1/1/<%=year(Date())-6%>",
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

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	var url;
	url = "../../logout.asp";
	if (ns)
		document.location = url;
	else
	{
		window.document.frmreport.action = url;
		window.document.frmreport.target = "_self";
		window.document.frmreport.submit();
	}	
}

function printpage() 
{ //v2.0
	if ("<%=strError%>" == "")
	{
		var objNewWindow;
		window.status = "";
	 
		strFeatures = "top=1,left="+(screen.width/2-380)+",width=1024,height=680,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";
	              
		if((objNewWindow) && (!objNewWindow.closed))
			objNewWindow.focus();	
		else 
		{
			objNewWindow = window.open('rpt_print_preview.asp?title=' + '<%=strTitle2%>', "MyNewWindow", strFeatures);
		}
		window.status = "Opened a new browser window.";  
	}	
}

function viewpage(kind)
{
	var intpage = parseInt(window.document.frmreport.txtpage.value,10);
	var curpage = "<%=CInt(intCurPage)%>";
	var pagetotal = "<%=CInt(intTotalPage)%>";
	
	if (kind == 1)
	{
		window.document.frmreport.txtpage.value = intpage
		if ((intpage > 0) & (intpage <= pagetotal) & (intpage != curpage)) 
		{
			document.frmreport.action = "rpt_sum_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "rpt_sum_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function checkdata()
{
	if (document.frmreport.rdotype[0].checked)
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


function submitform()
{
	if (checkdata())
	{
		document.frmreport.action = "rpt_sum_staff.asp?act=vra";
		document.frmreport.submit();			
	}		
}


function viewdetail(s,types)
{
	window.document.frmreport.txthidden.value = s;
	if (ns)
		document.location = "rpt_staff_detail.asp?t=" + types;
	else
	{
		document.frmreport.target = "_self"
		document.frmreport.action = "rpt_staff_detail.asp?t=" + types;
		document.frmreport.submit();
	}	

}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0" height="445" style="height:&quot;76%&quot;" align="center">
  <tr> 
    <td bgcolor="#FFFFFF" valign="top"> 
      <table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="28%"> 
            <table border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="0%">&nbsp;</td>
                <td width="35%" class="blue-normal">Name&nbsp;&nbsp;</td>
                <td width="62%"> 
                  <input type="text" name="txtname" class="blue-normal" size="16" style="width:153" value="<%=showvalue(strName)%>">
                </td>
                <td width="3%">&nbsp; </td>
              </tr>
              <tr> 
                <td width="0%">&nbsp;</td>
                <td width="35%" class="blue-normal">Department&nbsp;</td>
                <td width="62%"> 
                  <select id="lbdepartment" size="1" name="lbdepartment" class="blue-normal">
                    <option value="0" selected>&nbsp;</option>
<%=strDepartment%>	
                  </select>
                </td>
                <td width="3%">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td width="35%"> 
            <table width="50%" border="0" cellspacing="0" cellpadding="1">
              <tr> 
                <td valign="top" width="0%" class="blue" height="33">&nbsp;</td>
                <td width="5%" class="blue" align="right" height="33"> 
                  <input type="radio" name="rdotype" value="D" <%If strType = "D" Then %> checked <%End If%>language="javascript" onClick="document.frmreport.txtfrom.focus()">
                </td>
                <td class="blue-normal" width="7%" height="33">From </td>
                <td class="blue-normal" width="16%" height="33"> 
                  <input type="text" name="txtfrom" id="txtfrom" class="blue-normal" size="2" value="<%=strFrom1%>" style="width:50" language="javascript" onClick="document.frmreport.rdotype[0].checked=true">
                </td>
                <td width="5%" class="blue-normal" height="33">To </td>
                <td class="blue-normal" height="33"> 
                  <input type="text" name="txtto" id="txtto" class="blue-normal" size="2" value="<%=strTo1%>" style="width:50" language="javascript" onClick="document.frmreport.rdotype[0].checked=true">
                </td>
              </tr>
              <tr> 
                <td valign="top" width="0%" class="blue">&nbsp;</td>
                <td width="5%" class="blue" align="right"> 
                  <input type="radio" name="rdotype" value="M" <%If strType = "M" Or strType = "" Then %>checked <%End If%>language="javascript" onClick="document.frmreport.lbmonth.focus()">
                </td>
                <td class="blue-normal" width="7%">Month </td>
                <td class="blue-normal" width="16%"> 
				  <select name="lbmonth" size="1" class="blue-normal" language="javascript" onClick="document.frmreport.rdotype[1].checked=true">
					<%for iM=1 to 12%>
				    <option <%If CInt(intMonth)=iM Then%>selected<%End If%> value="<%=iM%>"><%=MonthName(iM,true)%></option>
				    <%next%>
				  </select>
                </td>
                <td width="5%" class="blue-normal"> Year </td>
                <td class="blue-normal"> 
				  <select name="lbyear" size="1" class="blue-normal" language="javascript" onClick="document.frmreport.rdotype[1].checked=true">
				<%For ii=Year(Date)-1 To Year(Date)%>
				    <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
				<%Next%>
				  </select>
                </td>
              </tr>
            </table>
          </td>
          <td width="37%"> 
            <table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
              <tr> 
                <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
                  <div align="center" class="blue"><a href="javascript:submitform();" class="b" onMouseOver="self.status='';return true">Submit</a></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor=<%If strError="" Then%>"FFFFFF"<%Else%>"#E7EBF5"<%End If%>>
		  <td class="red" height="20" align="left" width="100%"> &nbsp;<b><%=strError%></b></td>
	    </tr>
        <tr> 
          <td bgcolor="#8CA0D1"><img src="../../IMAGES/DOT-01.GIF" width="1" height="1"></td>
        </tr>
        <tr> 
          <td>&nbsp; </td>
        </tr>
      </table>
<%
'--------------------------------------------------
' Write the title of report page
'--------------------------------------------------

	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Staff Hours")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strTitle2)
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", formatdatetime(date,vbLongDate))
	Response.Write(arrPageTemplate(1))
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#617DC0"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="3">
              <tr > 
                <td class="blue" align="center" width="3%" bgcolor="#E7EBF5" rowspan="2">No.</td>
                <td class="blue" align="center" width="4%" bgcolor="#E7EBF5" rowspan="2">StaffID</td>
                <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">Employee Name </td>
                <td class="blue" align="center" width="11%" bgcolor="#E7EBF5" rowspan="2">Jobtitle </td>
                <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">Report To </td>
                <td class="blue" align="center" colspan="5" bgcolor="#E7EBF5">Worked Hours</td>
                <td class="blue" align="center" width="4%"valign="bottom" bgcolor="#E7EBF5" rowspan="2">Downtime<br>(4)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" rowspan="2">PD<br>(5)</td>
                <td class="blue" align="center" colspan="6" bgcolor="#E7EBF5">Off Hours </td>
                <td class="blue" align="center" width="7%" bgcolor="#E7EBF5" rowspan="2">Total<br>Available-hours<br>(1a)+(1b)+(2)+(3)+(4)</td>
                <td class="blue" align="center" width="7%" bgcolor="#E7EBF5" rowspan="2">Total hours<br>(1abc)+(2)+(3)+(4)+(5)+(6)+(7)+(8)+(9)+(10)</td>
                <td class="blue" align="center" width="5%" bgcolor="#E7EBF5" rowspan="2">Client<br> Hours(%)</td>
              </tr>
              <tr> 
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >Client Billable Hrs<br>(1a)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >Client Non-Billable Hrs<br>(1b)</td>
                 <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >OT Hrs<br>(1c)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5">ATL<br>(2)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5">GA<br>(3)</td>
				
				<td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >PH<br>(6)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >AH<br>(7)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >SL<br>(8)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >OL<br>(9)</td>
				<td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >Time<br>InLieu(10)</td>
                <td class="blue" align="center" width="3%" valign="bottom" bgcolor="#E7EBF5" >UL<br>(11)</td>
				
              </tr>
<%Response.Write strLast%>		      
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%if false then%>
<table width="90%" border="0" cellspacing="0" cellpadding="0" height="20" align="center">
  <tr> 
    <td align="right" bgcolor="#E7EBF5"> 
      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
        <tr class="black-normal"> 
          <td align="right" valign="middle" width="37%" class="blue-normal">Page</td>
          <td align="center" valign="middle" width="13%" class="blue-normal"> 
            <input type="text" name="txtpage" class="blue-normal" value="<%=intCurPage%>" size="5">
          </td>
          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:viewpage(1);" onMouseOver="self.status='';return true"><font color="#990000">Go</font></a></td>
          <td align="right" valign="middle" width="15%" class="blue-normal"><%If CInt(intTotalPage) <> 0 Or intTotalPage <> "" Then%>Pages <%=intCurPage%>/<%=intTotalPage%><%End If%>&nbsp;&nbsp;</td>
          <td valign="middle" align="right" width="28%" class="blue-normal"><%If CInt(intCurPage) <> 1 Then%><a href="javascript:viewpage(2);" onMouseOver="self.status='';return true">Previous</a><%End If%><%If CInt(intCurPage) <> 1 And  CInt(intCurPage) <> CInt(intTotalPage) Then%>/<%End If%><%If CInt(intCurPage) <> CInt(intTotalPage) And (CInt(intTotalPage) <> 0 Or intTotalPage <> "") Then%><a href="javascript:viewpage(3);" onMouseOver="self.status='';return true"> Next</a><%End If%>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>
<p>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))
%>
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="F" value="<%=strFrom%>">
<input type="hidden" name="T" value="<%=strTo%>">
<input type="hidden" name="P" value="<%=intCurPage%>">
</form>
</body>
</html>
