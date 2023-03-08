<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	dim strProjectID,strSql,strStatus,strID
	dim rsCSODetails,rsCountry
	dim intDetailID,intNewMonthNo,strMonthName,dblManHours,dblPaymentSchedule,dblExchangeRate,strCurrencyCode,dblTotalHours,dblTotalPay
	dim dblThirdPartyEstValue,dblThirdPartyEstHours,strThirdPartyCurrency,dblThirdPartyExRate
	dim intMonthDetail,intYearDetail
	Dim arrlstFrom(2),arrCategories,intCategoryType
	
	arrCategories =Array("CSO Detail", "CSO by Level")	
'--------------------------------------------------
' Get Invoices
'--------------------------------------------------
function GetCSODetailList(rsCSODetail,IDDetail)
	dim strResult,strBkg,strDate
	dim idx,dblTotal,dblGrandTotal
	
	dblGrandTotal=0
    dblGrandTotalTP=0            
    dblTotalPay=0
    dblTotalPayTP=0            
	dblTotalHours=0
	dblTotalHoursTP=0
	
	idx=0

	intMonthDetail=month(date())
	intYearDetail=Year(date())

	if rsCSODetail.RecordCount>0 then
		strResult=""

		Do while not rsCSODetail.EOF
			idx=idx+1
			dblTotal=0
			if rsCSODetail("CSODetailID")= cint(IDDetail) then
				
				intDetailID=rsCSODetail("CSODetailID")
				dblManHours=rsCSODetail("ManHours")
				dblPaymentSchedule=rsCSODetail("PaymentSchedule")	
				dblExchangeRate=rsCSODetail("ExchangeRate")
				intMonthDetail=rsCSODetail("MonthDetail")
				intYearDetail=rsCSODetail("YearDetail")
				
				dblThirdPartyEstValue=rsCSODetail("ThirdPartyEstValue")
				dblThirdPartyEstHours=rsCSODetail("ThirdPartyEstHours")
				strThirdPartyCurrency=rsCSODetail("ThirdPartyCurrency")
				dblThirdPartyExRate=rsCSODetail("ThirdPartyExRate")
				
				
			end if
						
			strBkg="#E7EBF5"
			if (idx mod 2=1) then strBkg="#FFF2F2"
			strMonthNameDisplay=MonthName(rsCSODetail("MonthDetail"),2) & "-" & rsCSODetail("YearDetail")

			strResult=strResult & "<tr bgcolor='" & strBkg & "'> "
            strResult=strResult & "<td valign='top' class='blue'>" & idx & ".</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'><a href='javascript:showdata(" & rsCSODetail("CSODetailID") & ")' class='c'><b>" & strMonthNameDisplay & "</b></a></td>"
			strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsCSODetail("ManHours"),2)  & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsCSODetail("PaymentSchedule"),2) & "</td>"
            dblTotal=cdbl(rsCSODetail("PaymentSchedule")) * cdbl(rsCSODetail("ExchangeRate"))
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(dblTotal,2) & "</td>"

            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsCSODetail("ThirdPartyEstHours"),2)  & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsCSODetail("ThirdPartyEstValue"),2)  & "</td>"

            dblTotalTP=cdbl(rsCSODetail("ThirdPartyEstValue")) * cdbl(rsCSODetail("ThirdPartyExRate"))
	          
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(dblTotalTP,2)  & "</td>"
            
            strResult=strResult & "</tr>"
            dblGrandTotal=dblGrandTotal + dblTotal
            dblGrandTotalTP=dblGrandTotalTP + dblTotalTP
            
            dblTotalPay=dblTotalPay + cdbl(rsCSODetail("PaymentSchedule"))
            dblTotalPayTP=dblTotalPayTP + cdbl(rsCSODetail("ThirdPartyEstValue"))
            
			dblTotalHours=dblTotalHours + cdbl(rsCSODetail("ManHours"))
			dblTotalHoursTP=dblTotalHoursTP + cdbl(rsCSODetail("ThirdPartyEstHours"))
            
			rsCSODetail.MoveNext
		loop

'		if dblTotal<>0 then
			strResult=strResult & "<tr bgcolor='#FFFFFF'><td colspan='2' align='right' valign='top' class='blue'>Total</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblTotalHours,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblTotalPay,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblGrandTotal,2) & "</td>" & _
									
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblTotalHoursTP,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblTotalPayTP,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblGrandTotalTP,2) & "</td>" & _
									"</tr>"
'		end if
		
	end if

	GetCSODetailList=strResult
end function

'--------------------------------------------------
' 
'--------------------------------------------------
function GetAPKList(rsAPKSearch)
	
	dim strReturn
	strReturn=""
	if rsAPKSearch.recordCount>0 then	
		
		do while not rsAPKSearch.EOF
			
			strReturn=strReturn & "<option value='" & rsAPKSearch("ProjectID") & "'>" & rsAPKSearch("ProjectID") & "-" & rsAPKSearch("ProjectName") &  "</option>" & vbCrLf 			
			rsAPKSearch.MoveNext
		loop	
		
	end if
	
	GetAPKList=strReturn
end function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
'--------------------------------------------------
' User can update all project invoice
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgInvoice = False
	Else
		varGetRight = session("RightOn")
		fgInvoice = False
		For ii = 0 To Ubound(varGetRight, 2)
'Response.Write 	varGetRight(0, ii)		 & "<br>"
			If varGetRight(0, ii) = "Invoice" Then

				fgInvoice = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If		
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	strProjectID=Left(Request.Form("txthidden"),15)
	strStatus=Request.Form("fgstatus")
	if strStatus="" then strStatus="A"
	
	strAPKList=""
	
	selectRow=Request.QueryString("r")	
	if selectRow="" then selectRow=-1	
	
	strID=-1
'--------------------------------------------------
' Get currency
'--------------------------------------------------
	Call GetRecordset("SELECT CurrencyCode FROM ATC_Projects WHERE ProjectID='" & strProjectID  & "'",rsCurrency)	
	strCurrencyCode=rsCurrency("CurrencyCode")
	
	strSql="SELECT * FROM ATC_Currency WHERE fgActivate=1"
	Call GetRecordset(strSql,rsCurrencyExist)
	
'--------------------------------------------------
' 
'--------------------------------------------------	
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase

	If Request.QueryString("act") = "save" and Request.QueryString("choose_menu")="" Then
	
		If objDatabase.dbConnect(strConnect) Then		
			
			strSql=""
			intDetailID=Request.Form("txtID")
			
			intMonthNo=Request.Form("txtMonthNo")
			
			intMonthDetail=Request.Form("lstmonthF")
			intYearDetail=Request.Form("lstYearF")
			
			strMonthName=Request.Form("txtMonthName")
			
			strMonthName=intMonthDetail & "/" & intYearDetail
			dblManHours=Request.Form("txtManHours")
			dblPaymentSchedule=Request.Form("txtPaymentSchedule")
			dblExchangeRate=Request.Form("txExchangeRate")
			if dblExchangeRate="" then dblExchangeRate=1
			
			dblThirdPartyEstValue=Request.Form("txtThirdPartyEstValue")
			dblThirdPartyEstHours=Request.Form("txtThirdPartyEstHours")
			strThirdPartyCurrency=Request.Form("lbCurrencyTP")
			dblThirdPartyExRate=Request.Form("txtThirdPartyExRate")
			if dblThirdPartyExRate="" then dblThirdPartyExRate=1

			select case strStatus
			
				'For add new
				case "A"
					strSql="INSERT INTO ATC_ProjectCSODetails (ProjectID,MonthName,ManHours,PaymentSchedule,ExchangeRate,MonthDetail,YearDetail,ThirdPartyEstValue,ThirdPartyEstHours,ThirdPartyCurrency,ThirdPartyExRate) VALUES " &_
							"('" & strProjectID & "','" & strMonthName & "'," & dblManHours & "," & dblPaymentSchedule & "," & dblExchangeRate &_
							"," & cint(intMonthDetail) & "," & intYearDetail &_
							"," & dblThirdPartyEstValue & "," & dblThirdPartyEstHours & ",'" & strThirdPartyCurrency & "'," & dblThirdPartyExRate & ")"
						
				'For edit
				case "E"
					strSql="UPDATE ATC_ProjectCSODetails SET " &_
								"MonthName='" & strMonthName & "'," &_
								"ManHours=" & dblManHours & ","&_
								"PaymentSchedule=" & dblPaymentSchedule & "," &_
								"ExchangeRate=" & dblExchangeRate & "," &_
								"monthDetail=" & cint(intMonthDetail) & "," &_
								"YearDetail=" & intYearDetail & "," &_
								"ThirdPartyEstValue=" & dblThirdPartyEstValue & "," &_
								"ThirdPartyEstHours=" & dblThirdPartyEstHours & "," &_
								"ThirdPartyCurrency='" & strThirdPartyCurrency & "'," &_
								"ThirdPartyExRate=" & dblThirdPartyExRate &_
							" WHERE CSODetailID =" & intDetailID
				'For delete
				case "D"
					strSql="DELETE ATC_ProjectCSODetails WHERE CSODetailID =" & intDetailID
					intMonthNo=""
					strMonthName=""
					dblManHours=""
					dblPaymentSchedule=""
					dblExchangeRate=""	
			end select
			if strSql<>"" then

				If objDatabase.runActionQuery(strSQL) Then
					strError = "Update successful."
				Else
					strError = objDatabase.strMessage
				End If	
			
			end if
			
		end if

	end if
	
	strSearchAPK=Request.Form("txtSearch")
		
	if strSearchAPK<>"" then
				
		strSql="SELECT ProjectID,ProjectName FROM ATC_Projects WHERE " &_
			" ProjectID like '%" & strSearchAPK & "%' ORDER BY ProjectID"
		
		Call GetRecordset(strSql,rsAPKSearch)
		if gMessage="" then strAPKList=GetAPKList(rsAPKSearch)
			
	end if
		
	if Request.QueryString("act") = "g" then strProjectID=Request.Form("lstSearch")
	
	strSql="SELECT a.CSODetailID,a.ProjectID,a.MonthNo,a.[MonthName],a.ManHours,ISNULL(a.PaymentSchedule,0) as PaymentSchedule ,a.ExchangeRate,a.TakeANote, a.MonthDetail, a.YearDetail, " & _
	        "ISNULL(ThirdPartyEstValue,0) as ThirdPartyEstValue,isnull(ThirdPartyEstHours,0) as ThirdPartyEstHours,ThirdPartyCurrency,isnull(ThirdPartyExRate,1) as ThirdPartyExRate  FROM ATC_ProjectCSODetails a " &_
			"WHERE a.ProjectID='" & strProjectID & "' ORDER BY a.YearDetail,a.MonthDetail "

	Call GetRecordset(strSql,rsCSODetails)
	
	
	if gMessage="" then strLast=GetCSODetailList(rsCSODetails,cint(selectRow))

	
	If Request.QueryString("act") = "save"	then
			
		If objDatabase.dbConnect(strConnect) Then		
			strSql="UPDATE ATC_Projects SET CSOMainHours=" & dblTotalHours & _
					", Value=" & dblTotalPay & _
					" WHERE ProjectID='" & strProjectID & "'"
			objDatabase.runActionQuery(strSQL)	
		end if
	end if

	arrlstFrom(0) = selectmonth("lstmonthF",intMonthDetail , -1)
	arrlstFrom(1) = selectyear("lstYearF", intYearDetail, 2000, year(now()) + 2, 0)
'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)
	Set objEmployee = Nothing
	
'--------------------------------------------------
' Make list of menu
'--------------------------------------------------

	If isEmpty(session("Menu")) Then 
		getRes = getarrMenu(intUserID)
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End If	
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	
<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>
<script language="javascript">
<!--
	
function showdata(r)
{
	document.frmreport.fgstatus.value = "E";
	document.frmreport.action = "Pro_CSODetails.asp?r=" + r;
	document.frmreport.submit();
}
	
function adddata()
{
	window.document.frmreport.fgstatus.value = "A"
	//window.document.frmreport.txtMonthName.value = "";
	
	//selObj.selectedIndex = num;
	
	window.document.frmreport.lstmonthF.selectedIndex = <%=Month(Date())-1%>;
	window.document.frmreport.lstYearF.selectedIndex = <%=Year(Date())-2000%>;
	
	window.document.frmreport.txtManHours.value = "";
	window.document.frmreport.txtPaymentSchedule.value = "";
	window.document.frmreport.txExchangeRate.value = "";
	window.document.frmreport.txtMonthNo.value="<%=intNewMonthNo%>"
	window.document.frmreport.lstmonthF.focus();
}

function Category(type)
{
	//window.document.frmreport.txtCategoryType.value = type
	window.document.frmreport.action = "Pro_CSODetails.asp"			
	window.document.frmreport.submit();
}

	
function savedata()
{
	if (checkdata()==true)
	{
		window.document.frmreport.action = "Pro_CSODetails.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	var answer = confirm("Are you sure you want to remove current item?");
	if (answer){
		window.document.frmreport.fgstatus.value = "D";
		window.document.frmreport.action = "Pro_CSODetails.asp?act=save";
		window.document.frmreport.submit();
	}	
}

function back_menu()
{
	window.document.frmreport.action = "n_projectlist.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function sub_menu()
{
	window.document.frmreport.action = "Pro_CSOByLevel.asp";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

	
function checkdata()
{
	var dblManHours=document.frmreport.txtManHours.value
	var dblPaySche=document.frmreport.txtPaymentSchedule.value
	var dblExrate=document.frmreport.txExchangeRate.value
	
	if (dblManHours==""){
		alert("Man hours must be required.");
		document.frmreport.txtManHours.focus();
		return false;
	}
	
	if (isNaN(dblManHours) ==  true) 
	{
		alert("Man hours must be number.");
		document.frmreport.txtManHours.focus(); 
		return false;
	}
	
	if (dblPaySche==""){
		alert("Payment schedule must be required.");
		document.frmreport.txtPaymentSchedule.focus();
		return false;
	}
	
	if (isNaN(dblPaySche) ==  true){
		alert("Payment schedule must be number.");
		document.frmreport.txtPaymentSchedule.focus();
		return false;
	}
	
	if (dblExrate==""){
		alert("The Exchange Rate must be required.");
		document.frmreport.txExchangeRate.focus();
		return false;
	}
	
	if (isNaN(dblExrate) ==  true) 
	{
		alert("The Exchange Rate must be number.");
		document.frmreport.txExchangeRate.focus(); 
		return false;
	}
	return true;
}

function search()
{
	document.frmreport.action = "Pro_CSODetails.asp?act=f"
	document.frmreport.target = "_self";
	document.frmreport.submit();
}

function go()
{
	document.frmreport.action = "Pro_CSODetails.asp?act=g"
	document.frmreport.target = "_self";
	document.frmreport.submit();
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
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<%
	If strError1 = "" Then
%>        
          <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%		If strError <> "" Then%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red">&nbsp;<b><%=strError%></b></td>
			</tr>
<%		End If%>				
			
				<tr align="center"> 
					<td class="blue" height="30" align="left" width="23%"> &nbsp;&nbsp; &nbsp;&nbsp; 
						<A href="javascript:back_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">Project List</a> | 
						<A href="javascript:sub_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">CSO by Level</a></td>
			     </tr>
				<tr> 
					<td align="center" valign="middle">
						<table width="98%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="15%" class="blue-normal" valign="middle" align="right"> Search for APK &nbsp; </td>
								<td width="20%" ><input type="text" name="txtsearch" class="blue-normal" size="15" style="width:98%" value="<%=strSearchAPK%>"></td>
								<td width="45%">
									<select name="lstSearch" style="width:98%" class="blue-normal" onChange="javascript:go()">
										<option value="-1"></option>
										<%=strAPKList%>
									</select></td>
								<td width="20%">
									<table width="100%" border="0" cellspacing="3" cellpadding="0" height="20" align="left">
										<tr> 
											<td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
												<a href="javascript:search();" class="b" onMouseOver="self.status='Search'; return true;" onMouseOut="self.status=''">Search</a></td>
										</tr>
									</table>
								</td>
							</tr>
						</table> 
					</td>
				</tr>                
			    <tr align="center"> 
				    <td class="title" height="50" align="center" >CSO Details</td>
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
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%">APK </td>
                              <td valign="middle" width="30%" class="blue"><%=strProjectID%></td>
                              <td valign="top" width="25%" class="blue-normal" align="center">&nbsp;</td>                             

                            </tr>
                           
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top"  class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Month</td>
                              <td valign="middle" class="blue-normal"> 
                              <%
									Response.Write arrlstFrom(0)
									Response.Write arrlstFrom(1)%>	
                              </td>
                              <td valign="top" class="blue-normal" align="center"></td>

                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Man hours</td>
                              <td valign="middle" class="blue-normal"> 
                                <input type="text" name="txtManHours" class="blue-normal" size="20" style="width:60%" value="<%=dblManHours%>">&nbsp;
                              </td>
                              <td valign="top" class="blue-normal" align="left"></td>
                             
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Invoice Schedule </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtPaymentSchedule" class="blue-normal" size="20" style="width:80%" value="<%=dblPaymentSchedule%>"><%if strCurrencyCode<>"" then%> (<%=strCurrencyCode%>)<%end if%></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Exchange Rate </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txExchangeRate" class="blue-normal" size="20" style="width:80%" value="<%=dblExchangeRate%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                                                       
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >&nbsp;</td>
                              <td valign="middle" class="blue-normal">&nbsp;</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="top" class="blue"><u>Outsourcing </u> </td>
                              <td valign="middle" class="blue-normal"> </td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                             

                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">TP Man hours</td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtThirdPartyEstHours" class="blue-normal" size="20" style="width:70%" value="<%=dblThirdPartyEstHours%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>     

                          
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">TP Value </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtThirdPartyEstValue" class="blue-normal" style="width:70%" value="<%=dblThirdPartyEstValue%>">
                              <select class='blue-normal' name='lbCurrencyTP' style='WIDTH: 28%'>
								<option value="" <%if strCurrency="" then%>selected<%end if%>>&nbsp;</option>
                            <%do while not rsCurrencyExist.EOF%>							
								<option value="<%=rsCurrencyExist("CurrencyCode")%>" <%if strThirdPartyCurrency=rsCurrencyExist("CurrencyCode") then%>selected<%end if%>><%=rsCurrencyExist("CurrencyCode")%></option>
								
							<% rsCurrencyExist.Movenext
							loop%>
                              
							</select>
                              </td>
                              <td valign="top" class="blue-normal" align="left">                                                            
                              &nbsp;</td>
                            </tr>
                            
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Ex. Rate </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtThirdPartyExRate" class="blue-normal" size="20" style="width:70%" value="<%=dblThirdPartyExRate%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>               
                            
                          </table>
                          <input type="hidden" name="txtID" value="<%=intDetailID%>">
                           
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="180" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60"> 
                                      <div align="center" class="blue"><a href="javascript:adddata()" onMouseOver="self.status='Please click here to add new record';return true" onMouseOut="self.status='';return true" class="b">Add</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:savedata()" onMouseOver="self.status='Please click here to save changes';return true" onMouseOut="self.status='';return true" class="b">Save</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:deletedata()" onMouseOver="self.status='Please click here to delete this record';return true" onMouseOut="self.status='';return true" class="b">Delete</a></div>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>

                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                  
                            <tr bgcolor="#8CA0D1">
                              <td width="5%" rowspan="2" align="center" bgcolor="#8CA0D1" class="blue">No</td>
                              <td width="15%" rowspan="2" align="center" class="blue">Month</td>
                              <td colspan="3" align="center" class="blue">CSO Information </td>
                              <td colspan="3" align="center" class="blue">Outsourcing </td>
                            </tr>
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" align="center" width="12%">Man hours</td>
                              <td class="blue" align="center" width="14%">Payment Schedule<br><%if strCurrencyCode<>"" then%> (<%=strCurrencyCode%>)<%end if%></td>
                              <td class="blue" align="center" width="14%">Payment Schedule<br>(USD)</td>

                             <td class="blue" align="center" width="12%">TP Man hours</td>
                              <td class="blue" align="center" width="14%">TP Value</td>
                              <td class="blue" align="center" width="14%">TP Value<br>(USD)</td>

                            </tr>

<%=strLast%>
                          </table>
<%if strLast<>"" then%>                          
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue-normal">&nbsp;&nbsp;* Click on Month. to update</td>
                            </tr>
                          </table>
<%end if%>                          
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
<%	Else
		If strError <> "" Then
%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red">&nbsp;<%=strError%></td>
				</tr>
<%		End If%>				

		  <tr>
         	<td class="red" align="center" valign="middle"><b><%=strError1%></b></td>
		  </tr>	          
<%	End If%>		  
        </table>
<%
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
	Response.Write(arrTmp(1))%>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))%>
	
<input type="hidden" name="txthidden" value="<%=strProjectID%>">
<input type="hidden" name="fgstatus" value="<%=strStatus%>">

<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">

</form>

</body>
</html>