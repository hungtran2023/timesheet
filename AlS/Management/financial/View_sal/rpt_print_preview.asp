<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->
<!-- #include file = "../../../inc/library.asp"-->

<%
	Dim ii, intRow
	Dim varUser
	Dim strTitle
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
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	strTitle = Request.QueryString("title")
	arrSalSheet=Session("arrSalSheet")
	If isArray(varStaff) Then
		intRows = Ubound(varStaff,2)
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
<title>Atlas Industries - Timesheet System</title>

<link rel="stylesheet" href="../../../timesheet.css">

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="445" style="height:&quot;76%&quot;" align="center">
  <tr> 
    <td bgcolor="#FFFFFF" valign="top"> 
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
			<td>
<%
'--------------------------------------------------
' Write the title of report page
'--------------------------------------------------

	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary Salary of Employees")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strTitle)
	
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
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>
