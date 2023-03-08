<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<%
	Dim ii, intRow
	Dim varUser
	Dim strTitle
	
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	strTitle = Request.QueryString("title")
	varStaff   = session("varStaff")
	If isArray(varStaff) Then
		intRows = Ubound(varStaff,2)
	End If	 		

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
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<table width="90%" border="0" cellspacing="0" cellpadding="0" height="445" style="height:&quot;76%&quot;" align="center">
  <tr> 
    <td bgcolor="#FFFFFF" valign="top"> 
<%
'--------------------------------------------------
' Write the title of report page
'--------------------------------------------------
arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Staff Hours")
arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strTitle)
arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", formatdatetime(date,vbLongDate))
Response.Write(arrPageTemplate(1))

strLast=Session("StrLast")
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#617DC0"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="3">
               <tr > 
                <td class="blue" align="center" width="3%" bgcolor="#E7EBF5" rowspan="2">No.</td>
                <td class="blue" align="center" width="5%" bgcolor="#E7EBF5" rowspan="2">StaffID</td>
                <td class="blue" align="center" width="15%" bgcolor="#E7EBF5" rowspan="2">Employee Name </td>
                <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">Jobtitle </td>
                <td class="blue" align="center" colspan="5" bgcolor="#E7EBF5">Worked Hours</td>
                <td class="blue" align="center" width="4%"valign="bottom" bgcolor="#E7EBF5" rowspan="2">Downtime<br>(4)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" rowspan="2">PD<br>(5)</td>
                <td class="blue" align="center" colspan="6" bgcolor="#E7EBF5">Off Hours </td>
                <td class="blue" align="center" width="7%" bgcolor="#E7EBF5" rowspan="2">Total<br>Available-hours<br>(1a)+(1b)+(2)+(3)+(4)</td>
                <td class="blue" align="center" width="7%" bgcolor="#E7EBF5" rowspan="2">Total hours<br>(1a)+(1b)+(1c)+(2)+(3)+(4)+(5)+(6)+(7)+(8)+(9)+(10)</td>
                <td class="blue" align="center" width="5%" bgcolor="#E7EBF5" rowspan="2">Client<br> Hours(%)</td>
              </tr>
              <tr> 
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >Client Billable Hrs<br>(1a)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >Client Non-Billable Hrs<br>(1b)</td>
                 <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >OT Hrs<br>(1c)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5">ATL<br>(2)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5">GA<br>(3)</td>
				
				<td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >PH<br>(6)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >AH<br>(7)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >SL<br>(8)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >OL<br>(9)</td>
				<td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >Time<br>InLieu(10)</td>
                <td class="blue" align="center" width="4%" valign="bottom" bgcolor="#E7EBF5" >UL<br>(11)</td>
				
              </tr>
<%Response.Write strLast%>            

		      
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
</html>