<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<%
'--------------------------------------------------
' Preparing data
'--------------------------------------------------
strfromto = Request.QueryString("fromto")
strprintdate = Request.QueryString("printdate")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_report.htm")

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
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%"  align="center" >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top"> 
    		<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Projects")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strfromto)
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
			%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                <tr> 
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5">Project 
                    ID </td>
                  <td class="blue" align="center" width="22%" bgcolor="#E7EBF5">Project 
                    Name </td>
                  <td class="blue" align="center" width="25%" bgcolor="#E7EBF5">Full 
                    Name </td>
                  <td class="blue" align="center" width="26%" bgcolor="#E7EBF5">Sub-task 
                    Description </td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5">Hours<br> 
                    Worked </td>
                </tr>
<%

	arrSrc = session("arrSumPro")

	lastU = Ubound(arrSrc, 2)
Response.Write lastU
	For i = 0 to 11594
		if cdbl(arrSrc(5, i)) = 0 then
%>
			<tr bgcolor="#FFFFFF">
				<td valign='top' width='15%' class='blue-normal'>&nbsp;<%=showlabel(arrSrc(0, i))%></td>
				<td valign='top' width='22%' class='blue-normal'>&nbsp;<%=showlabel(arrSrc(1, i))%></td>
				<td valign='top' width='25%' class='blue-normal'>&nbsp;<%=showlabel(arrSrc(2, i))%></td>
				<td valign='top' width='26%' class='blue-normal'>&nbsp;<%=showlabel(arrSrc(3, i))%></td>
				<td valign='top' width='12%' class='blue-normal' align='right'><%=FormatNumber(arrSrc(4, i), 2)%></td>
<%
		else
%>			<tr bgcolor="#FFF2F2">
				<td valign='top' colspan='4' class='blue' align='right'><%=arrSrc(3, i)%></td>
				<td valign='top' width='12%' class='blue' align='right'><%=FormatNumber(arrSrc(4, i), 2)%></td>
<%
		end if
%>			</tr>
<%
	Next
	set arrSrc = nothing
%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>