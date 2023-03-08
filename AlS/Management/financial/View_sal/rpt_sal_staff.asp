<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->
<!-- #include file = "../../../inc/library.asp"-->

<%
Sub GenerateReport()

	dim idxDirect,intNum,strSQL,arrStaff
	dim dblSalGrantTotal,dblUnpaidGrantTotal,dblOTGrantTotal,dblNetGrantTotal,dblOTPreGrantTotal
	dblSalGrantTotal=0
	dblUnpaidGrantTotal=0
	dblOTGrantTotal=0
	dblNetGrantTotal=0
	dblOTPreGrantTotal=0
	
	'Two part 0: Direct staff - 1: indirect staff
	For idxDirect=0 to 1
		
		'Get list of staff for indirect or direct
		strSQL = "exec GetStaffForSalarySheet '" & strFirstDay & "','" & strLastDay & "'," & intDepartment & ",'" & strName & "'," & intType & "," & idxDirect 

		arrStaff=GetStaffForSalarySheet(strSQL)
		
		if not IsEmpty(arrStaff) then
			
			call GetSalarySheet(arrStaff,strFirstDay,strLastDay,idxDirect,arrSalSheet)

			dblNetGrantTotal=dblNetGrantTotal +  cdbl(arrSalSheet(3,ubound(arrSalSheet,2)))
			dblSalGrantTotal=dblSalGrantTotal + cdbl(arrSalSheet(4,ubound(arrSalSheet,2)))
			dblUnpaidGrantTotal=dblUnpaidGrantTotal + cdbl(arrSalSheet(5,ubound(arrSalSheet,2)))
			dblOTGrantTotal=dblOTGrantTotal + cdbl(arrSalSheet(6,ubound(arrSalSheet,2)))
			dblOTPreGrantTotal=dblOTPreGrantTotal+ cdbl(arrSalSheet(10,ubound(arrSalSheet,2)))
		end if
		
	Next
		
	Redim Preserve arrSalSheet(13, ubound(arrSalSheet,2) + 1)
	arrSalSheet(0,ubound(arrSalSheet,2))=-2
	arrSalSheet(1,ubound(arrSalSheet,2))="Grant Total"
	arrSalSheet(3,ubound(arrSalSheet,2))=dblNetGrantTotal
	arrSalSheet(4,ubound(arrSalSheet,2))=dblSalGrantTotal
	arrSalSheet(5,ubound(arrSalSheet,2))=dblUnpaidGrantTotal
	arrSalSheet(6,ubound(arrSalSheet,2))=dblOTGrantTotal
	arrSalSheet(10,ubound(arrSalSheet,2))=dblOTPreGrantTotal
	Session("arrSalSheet")=arrSalSheet

End Sub

'-----------------------------------------------------------------
'Get Gross Up
'-----------------------------------------------------------------
Function GetGrossUp(byval dblTotal, byval intTaxType)
	dim dblGross

	if cint(intTaxType) = 2 then	
		'dblGross=8000000
		if dblTotal<=8000000 then 
			dblGross=dblTotal
		elseif dblTotal<=18800000 then
			dblGross=(dblTotal-800000)/0.9
		elseif dblTotal<=42800000 then
			dblGross=(dblTotal-2800000)/0.8
		elseif dblTotal<=63800000 then
			dblGross=(dblTotal-7800000)/0.7
		else
			dblGross=(dblTotal-15800000)/0.6
		end if		
	else
		'dblGross=5000000
		if dblTotal<=5000000 then 
			dblGross=dblTotal
		elseif dblTotal<=14000000 then
			dblGross=(dblTotal-500000)/0.9
		elseif dblTotal<=22000000 then
			dblGross=(dblTotal-2000000)/0.8
		elseif dblTotal<=32500000 then
			dblGross=(dblTotal-4500000)/0.7
		else
			dblGross=(dblTotal-8500000)/0.6
		end if		
	end if
	
	GetGrossUp=dblGross	
End Function

'-----------------------------------------------------------------
'Social Insurance
'-----------------------------------------------------------------
Function SocialInsurance(byval intYear,byval dblNET, byval blnProbation)
	
	dim dblLimit
	
	dblLimit=9000000
	if intYear>=2008 then dblLimit=10800000


	if 	dblNET >=dblLimit then dblNET=dblLimit

	SocialInsurance=dblNET * 0.2 * IIF(blnProbation,0,1)

End Function

'-----------------------------------------------------------------
'Health Insurance
'-----------------------------------------------------------------
Function HealthInsurance(byval dblNET, byval blnProbation)
	
	HealthInsurance=dblNET * 0.03 * IIF(blnProbation,0,1)
		
End Function

'-----------------------------------------------------------------
'Probation Value
'-----------------------------------------------------------------
Function Probation(byval dblNET, byval blnProbation)
	dim dblProbation
	dblProbation=0

	if blnProbation then dblProbation=dblNET
	
	Probation=dblProbation
End Function
'**********************************************************************************
	
	Dim strTitle, strFunction, strMenu
	Dim objEmployee, strError, intPageSize, fgRight 'view all or Not
	Dim strName,intDepartment,intType,intMonth,intYear,strTitle2,strFirstDay,strLastDay
	dim arrSalSheet
	
	strName = Request.Form("txtname")
	
	intDepartment = Request.Form("lbdepartment")
	if intDepartment="" then intDepartment=0
	intType = Request.Form("lbperson")
	if intType="" then intType=1
	
	intMonth = Request.Form("M")
	if intMonth="" then intMonth=Month(date())
	intYear	= Request.Form("Y")	
	if intYear="" then intYear=Year(date())
	strTitle2	= SayMonth(intMonth) & " - " & intYear
	if cint(intDepartment)>0 then strTitle2 = "Department: @@Depart <br>" & strTitle2
	
	strFirstDay = FirstOfMonth(intMonth,intYear)								' Get the first day in a month
	strLastDay = FirstOfMonth(intMonth,intYear) + (GetDay(intMonth,intYear) -1)	' Get the last day in a month
	
	if not IsEmpty(Session("arrSalSheet")) then Session("arrSalSheet")=empty
	'--------------------------------------------------
	'Initiate array of salary sheet.It has got 11 column
	'---------------------------------------------------
	Redim arrSalSheet(13,-1)
	
	call GenerateReport()
		
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../../message.htm")
	End If					

	intUserID = session("USERID")
	
'--------------------------------------------------
' Initialize department array
'--------------------------------------------------
	varDepartment = GetDepartment()
	if not IsEmpty(varDepartment) then
		For ii = 0 To ubound(varDepartment,2)
			if CInt(intDepartment)=CInt(varDepartment(0,ii)) then strTitle2 = Replace(strTitle2,"@@Depart", varDepartment(1,ii))
		next
	end if
	
'--------------------------------------------------
' End Of initializing department array
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))

	strFunction = "<a class='c' href='../../../welcome.asp?choose_menu=AA'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  strtmp1 & "&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage();'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' title='Log Out'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#'>Help</a>&nbsp;&nbsp;&nbsp;"
	Set objEmployee = Nothing
	
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

	Call ReadFromTemplateAll(arrPageTemplate, "../../../templates/template1/", "ats_report.htm")

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
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "<img src='../../../images/" & arrTmp(6, 0) & "' border='0'>" )
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
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../../timesheet.css">
<script language="javascript" src="../../../library/library.js"></script>

<script language="javascript">
<!--
ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	var url;
	url = "../../../logout.asp";
	if (ns)
		document.location = url;
	if (ie)
	{
		window.document.frmreport.action = url;
		window.document.frmreport.target = "_self";
		window.document.frmreport.submit();
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
			document.frmreport.action = "rpt_sal_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "rpt_sal_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function submitform()
{
	window.document.frmreport.M.value = window.document.frmreport.lbmonth.options[window.document.frmreport.lbmonth.selectedIndex].value;
	window.document.frmreport.Y.value = window.document.frmreport.lbyear.options[window.document.frmreport.lbyear.selectedIndex].value

	window.document.frmreport.action = "rpt_sal_staff.asp?act=vra";
	window.document.frmreport.submit();
}

function printpage() 
{ //v2.0
	if ("<%=strError%>" == "")
	{
		var objNewWindow;
		window.status = "";
	 
		strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=680,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=yes,scrollbars=yes";
	              
		if((objNewWindow) && (!objNewWindow.closed))
			objNewWindow.focus();	
		else 
		{
			objNewWindow = window.open('rpt_print_preview.asp?title=' + '<%=strTitle2%>', "MyNewWindow", strFeatures);
		}
		window.status = "Opened a new browser window.";  
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
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#FFFFFF" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="28%" valign="middle" height="28"> 
            <table border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td width="35%" class="blue-normal" valign="middle" height="28">Name&nbsp;&nbsp;</td>
                <td width="62%" valign="middle" height="28"> 
                  <input type="text" name="txtname" class="blue-normal" size="16" style="width:150" value="<%=showvalue(strName)%>">
                </td>
                <td width="3%" valign="middle" height="28">&nbsp; </td>
              </tr>
              <tr> 
                <td width="35%" class="blue-normal" valign="middle" height="28">Department&nbsp;</td>
                <td width="62%" valign="middle" height="28"> 
                  <select id="lbdepartment" size="1" name="lbdepartment" class="blue-normal">
                    <option value="0" selected>&nbsp;</option>
<%			For ii = 0 To ubound(varDepartment,2)%>                    
					<option <%If CInt(intDepartment)=CInt(varDepartment(0,ii)) Then%> selected <%End If%> value="<%=varDepartment(0,ii)%>"><%=showlabel(varDepartment(1,ii))%></option>
<%			Next%>					
                  </select>
                </td>
                <td width="3%" valign="middle">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td width="35%" valign="middle" align="center"> 
            <table width="90%" border="0" cellspacing="0" cellpadding="1">
              <tr> 
                <td class="blue-normal" width="28%" valign="middle" height="28">Month&nbsp;</td>
                <td class="blue-normal" width="32%" valign="middle" height="28"> 
				  <select name="lbmonth" size="1" class="blue-normal">
					<%For ii=1 To 12%>
					  <option <%If CInt(intMonth)=ii Then%>selected<%End If%> value="<%=ii%>"><%=SayMonth(ii)%></option>
					<%next%>
				  </select>
                </td>
                <td width="13%" class="blue-normal" valign="middle" height="28">Year </td>
                <td class="blue-normal" valign="middle" height="28" width="27%"> 
				  <select name="lbyear" size="1" class="blue-normal">
					<%For ii=2000 To Year(Date)%>
					    <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
					<%Next%>
				  </select>
                </td>
              </tr>
              <tr>
                <td class="blue-normal" width="28%">Staff Type&nbsp;</td>			              
                <td class="blue-normal" colspan="3" valign="middle" height="28">
 				  <select name="lbperson" size="1" class="blue-normal">
					<option <%If CInt(intType)=1 Or intType = "" Then%>selected<%End If%> value="1">Vietnamese Person</option>
					<option <%If CInt(intType)=2 Then%>selected<%End If%> value="2">Foreigner Person</option>
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
          <td bgcolor="#8CA0D1"><img src="../../../IMAGES/DOT-01.GIF" width="1" height="1"></td>
        </tr>
        <tr> 
          <td>&nbsp; </td>
        </tr>
      </table>
    </td>
  </tr>
</table>      

      <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
			<td>
<%
'--------------------------------------------------
' Write the title of report page
'--------------------------------------------------

	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary Salary of Employees")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strTitle2)
	
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", formatdatetime(date,vbLongDate))
	Response.Write(arrPageTemplate(1))
%>			
			</td>
		</tr>
        <tr> 
          <td bgcolor="#617DC0"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="3">
              <tr bgcolor="#E7EBF5"> 
                  <td class="blue" align="center" >No.</td>
                  <td class="blue" align="center" >StaffID</td>
                  <td class="blue" align="center" width="12%">Full Name</td>
                  <td class="blue" align="center">Bank Account</td>
                  <td class="blue" align="center">Bank Detail</td>                  
                  <td class="blue" align="center">Net Salary</td>
                  <td class="blue" align="center">Probation & other</td>                  
                  <td class="blue" align="center">Salary<br> (as ATS hours)</td>
                  <td class="blue" align="center">Unpaid (minus)</td>

                  <td class="blue" align="center">Salary in month</td>
                  <td class="blue" align="center">OT <br> this month</td>
                  <td class="blue" align="center">OT <br> last month</td>
                  <td class="blue" align="center">Total<br>(Inc. OT last month)</td>
                  <td class="blue" align="center">Gross up</td>
                  <td class="blue" align="center">PIT payable</td>
                  <td class="blue" align="center">SI</td>
                  <td class="blue" align="center">HI</td>
              </tr>
<%
	If ubound(arrSalSheet,2) >= 0 Then
		intCount = 0
		strDepartment=""
		dim subtotal(11)
		dim total(4)
		dim grantTotal(4)

		for jj=0 to ubound(subtotal)
			if jj<=4 then 
				grantTotal(jj)=0
				total(jj)=0
			end if
			subtotal(jj)=0
		next 
		
		For ii = intCurRow To ubound(arrSalSheet,2)	
				
			if strDepartment<>arrSalSheet(7,ii) then
				if strDepartment<>"" then%>
				<tr bgcolor="#FFFFFF"> 
					<td valign="top" class="blue" align="right" colspan="5"><%="Subtotal for " & Ucase(strDepartment)%></td>
					<td valign="top" class="blue" align="right"><%=formatnumber(subtotal(0),0)%></td>
					<td valign="top" class="blue" align="right"><%=formatnumber(subtotal(1),0)%></td>
					<td valign="top" class="blue" align="right"><%=formatnumber(subtotal(2),0)%></td>
					<td valign="top" class="blue" align="right"><%=formatnumber(cdbl(subtotal(3)) * (-1),0)%></td>		
					
					<%For kk=4 to Ubound(subtotal)%>
					<td valign="top" class="blue" align="right"><%=formatnumber(subtotal(kk),0)%></td>
					<%Next%>
				</tr>
				
<%					'4: probation
					grantTotal(4)=grantTotal(4) + subtotal(1)
					for jj=0 to ubound(subtotal)
						'GranTotal - 0:GrossUp - 1:PIT Payment - 2: SI - 3:HI
						if jj>=8 then grantTotal(jj-8)=grantTotal(jj-8) + subtotal(jj)
						subtotal(jj)=0
					next 
					
				end if
				strDepartment=arrSalSheet(7,ii)
				if arrSalSheet(7,ii)<>"" then%>								
				<tr bgcolor="#D2DAEC"> 
					<td valign="top" class="blue" colspan="17">&nbsp;&nbsp;<%=arrSalSheet(7,ii)%></td>
				</tr>
				
<%				end if
			end if
			
			If cint(arrSalSheet(0,ii))<0 Then
%>              <tr bgcolor="#C2CCE7"> 
					<td valign="top" class="blue" align="right" colspan="5"><%=arrSalSheet(1,ii)%></td>
					<%'NET Sal.%>
					<td valign="top" class="blue" align="right"><%=formatnumber(arrSalSheet(3,ii),0)%></td>
					<%'Probation%>
					<td valign="top" class="blue" align="right"><%=IIF(cint(arrSalSheet(0,ii))=-1 AND grantTotal(4)>0,formatnumber(cdbl(grantTotal(4)),0),"")%></td>					
					<%'Sal ATS%>
					<td valign="top" class="blue" align="right"><%=formatnumber(arrSalSheet(4,ii),0)%></td>
					<%'Unpaid%>
					<td valign="top" class="blue" align="right"><%=formatnumber(cdbl(arrSalSheet(5,ii)) * (-1),0)%></td>							
					<%'Salary in Month%>
					<td valign="top" class="blue" align="right"><%=formatnumber(cdbl(arrSalSheet(3,ii))- cdbl(arrSalSheet(5,ii)),0)%></td>
					<%'OT this month%>
					<td valign="top" class="blue" align="right"><%=formatnumber(arrSalSheet(6,ii),0)%></td>
					<%'OT last month%>
					<td valign="top" class="blue" align="right"><%=formatnumber(arrSalSheet(10,ii),0)%></td>
					<%'Total (Inc. OT) - Bank transfer: NET - Unpaid + OT Last month%>
					<td valign="top" class="blue" align="right"><%=formatnumber(cdbl(arrSalSheet(3,ii))- cdbl(arrSalSheet(5,ii)) + cdbl(arrSalSheet(10,ii)),0)%></td>
<%					For kk=0 to ubound(grantTotal) -1 
						if cint(arrSalSheet(0,ii))=-1 then%>
					<td valign="top" class="blue" align="right"><%=formatnumber(grantTotal(kk),0)%></td>
<%							total(kk)=total(kk)+ grantTotal(kk)
							grantTotal(kk)=0
						else%>
					<td valign="top" class="blue" align="right"><%=formatnumber(Total(kk),0)%></td>
<%						end if
					Next%>
				</tr>
<%			Else%>      
				 <tr bgcolor="#FFFFFF">
					<td class="blue" align="center"><%=arrSalSheet(0,ii)%></td>
					<td class="blue-normal"><%=arrSalSheet(13,ii)%></td>
					<td class="blue-normal"><%=arrSalSheet(1,ii)%></td>
					<td class="blue-normal"><%=arrSalSheet(2,ii)%></td>
					<td class="blue-normal">&nbsp;<%=arrSalSheet(8,ii)%></td>
					
					<td class="blue-normal" align="right"><%=formatnumber(arrSalSheet(3,ii),0)%></td>
<%					dblPro=arrSalSheet(9,ii)%>					
					<td valign="center" class="blue-normal" align="right"><%=IIF(dblPro=0,"",formatnumber(dblPro,0))%></td>
					<td class="blue-normal" align="right"><%=formatnumber(arrSalSheet(4,ii),0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(cdbl(arrSalSheet(5,ii)) * (-1),0)%></td>				
					
					<td class="dark-red" align="right" bgcolor="#FFF2F2"><%=formatnumber(cdbl(arrSalSheet(3,ii))- cdbl(arrSalSheet(5,ii)),0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(arrSalSheet(6,ii),0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(arrSalSheet(10,ii),0)%></td>
<%					
					totalIncOT=cdbl(arrSalSheet(3,ii)) + cdbl(arrSalSheet(10,ii)) - cdbl(arrSalSheet(5,ii))%>
					<td class="dark-red" align="right" bgcolor="#FFF2F2"><%=formatnumber(totalIncOT ,0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(GetGrossUp(totalIncOT,arrSalSheet(11,ii)),0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(GetGrossUp(totalIncOT,arrSalSheet(11,ii))-totalIncOT,0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(SocialInsurance(intYear,arrSalSheet(3,ii),arrSalSheet(9,ii)),0)%></td>
					<td class="blue-normal" align="right"><%=formatnumber(HealthInsurance(arrSalSheet(3,ii),arrSalSheet(9,ii)),0)%></td>
				</tr>
<%				'NET Salary
				subtotal(0)=subtotal(0) + arrSalSheet(3,ii)
				'Probation
				subtotal(1)=subtotal(1) + dblPro
				'Salary in ATS
				subtotal(2)=subtotal(2) + arrSalSheet(4,ii)
				'Unpaid
				subtotal(3)=subtotal(3) + arrSalSheet(5,ii)
				'Salary inmonth
				subtotal(4)=subtotal(4) + (cdbl(arrSalSheet(3,ii))- cdbl(arrSalSheet(5,ii)))
				'OT in month
				subtotal(5)=subtotal(5) +  arrSalSheet(6,ii)
				'OT last month
				subtotal(6)=subtotal(6) + arrSalSheet(10,ii)
				'Total (inc. OT last month)
				subtotal(7)=subtotal(7) + (cdbl(arrSalSheet(3,ii)) + cdbl(arrSalSheet(10,ii)) - cdbl(arrSalSheet(5,ii)))
				'Gross up
				subtotal(8)=subtotal(8)  + GetGrossUp(totalIncOT,arrSalSheet(11,ii))
				'PIT
				subtotal(9)=subtotal(9) + GetGrossUp(totalIncOT,arrSalSheet(11,ii))-totalIncOT
				'SI
				subtotal(10)=subtotal(10) + SocialInsurance(intYear,arrSalSheet(3,ii),arrSalSheet(9,ii))
				'HI
				subtotal(11)=subtotal(11) + HealthInsurance(arrSalSheet(3,ii),arrSalSheet(9,ii))
				
			End If
			'If intCount = intPagesize Then Exit For
		Next
	End If		
%>
            </table>
          </td>
        </tr>
        <tr>
			<td>&nbsp;&nbsp;</td>
		</tr>        
      </table>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))
%>
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="P" value="<%=intCurPage%>">
</form>
</body>
</html>
