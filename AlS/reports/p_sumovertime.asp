<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<%

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
Function GenerateSumReportByStaff(rs)
	dim strReport,intNo,dblTotal,dblTotalAvailable
	dim dblSubtotal(5),dblGrandTotal(5)
	dim strName
	
	strReport=""

	call ResetArrayByInt(dblSubtotal)
	call ResetArrayByInt(dblGrandTotal)
	
		rs.MoveFirst
		
		intNo=1
		intI=0
				
		rs.Sort="Fullname"
  	    
  	    strName=""
		do while not rs.EOF 'and intI < rs.PageSize
 
		    'if strName<> rs("IDNumber") then
		        ' if strName<>"" then
		            ' strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		                    ' "<td valign='top' colspan='4' class='blue' align='right'>Total (" & strName & "): </td>"
		            ' for jj=0 to UBound(dblSubtotal)-1
		                    ' strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
		            ' next       
		            ' strReport= strReport & "</tr>"
		            ' call ResetArrayByInt(dblSubtotal)
		        ' end if
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		        strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("IDNumber") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Fullname") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Jobtitle") & "</td>" 
		        strName= rs("IDNumber")
		    'else
		        'strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		         'strReport= strReport & "<td valign='top' class='blue-normal' colspan='3'>--</td>" 
		    'end if
		    
		    dblTotal=0
    		
		    for j=0 to 3    		    
    		    dblTotal=dblTotal + CDbl(rs.Fields(6+j))  
    		    dblSubtotal(j)=dblSubtotal(j)+ CDbl(rs.Fields(6+j))
    		    dblGrandTotal(j)=dblGrandTotal(j)+ CDbl(rs.Fields(6+j))
		    next
            dblSubtotal(j)=dblSubtotal(j)+ dblTotal
    		dblGrandTotal(j)=dblGrandTotal(j)+ dblTotal
            
            strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("Project") & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue'align='right'>"& FormatNumber(dblTotal,1) & "</td>" & _
                    "</tr>" 
    			
			    intNo=intNo+1
			    rs.MoveNext
			    intI=intI+1
		loop

        ' strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		           ' "<td valign='top' colspan='4' class='blue' align='right'>Total: </td>"
        ' for jj=0 to UBound(dblSubtotal)-1
                ' strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
        ' next       
        ' strReport= strReport & "</tr>"
        
        strReport= strReport & "<tr bgcolor='#FFE1E1'>" &_
       "<td valign='top' colspan='4' class='blue' align='right'>Overall Total:</td>"
        for jj=0 to UBound(dblGrandTotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblGrandTotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
	GenerateSumReportByStaff=strReport
End Function

'--------------------------------------------------
'
'--------------------------------------------------
Function GenerateSumReportByProject(rs)
	dim strReport,intNo,dblTotal,dblTotalAvailable
	dim dblSubtotal(5),dblGrandTotal(5)
	dim strName
	
	strReport=""

	call ResetArrayByInt(dblSubtotal)
	call ResetArrayByInt(dblGrandTotal)
	
		rs.MoveFirst
		
		intNo=1
		intI=0
				
		rs.Sort="Project"
  	    
  	    strName=""
		do while not rs.EOF 'and intI < rs.PageSize
 
		    if strName<> rs("Project") then
		        if strName<>"" then
		            strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		                    "<td valign='top' colspan='4' class='blue' align='right'>Total(" & strName &") </td>"
		            for jj=0 to UBound(dblSubtotal)-1
		                    strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
		            next       
		            strReport= strReport & "</tr>"
		            call ResetArrayByInt(dblSubtotal)
		        end if
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		         strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("Project") & "</td>"
		        strName= rs("Project")
		    else
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		         strReport= strReport & "<td valign='top' class='blue-normal'>--</td>" 
		    end if
		    
		    dblTotal=0
    		
		    for j=0 to 3    		    
    		    dblTotal=dblTotal + CDbl(rs.Fields(6+j))  
    		    dblSubtotal(j)=dblSubtotal(j)+ CDbl(rs.Fields(6+j))
    		    dblGrandTotal(j)=dblGrandTotal(j)+ CDbl(rs.Fields(6+j))
		    next
            dblSubtotal(j)=dblSubtotal(j)+ dblTotal
    		dblGrandTotal(j)=dblGrandTotal(j)+ dblTotal
    		
             strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("IDNumber") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Fullname") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Jobtitle") & "</td>" & _
            
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue'align='right'>"& FormatNumber(dblTotal,1) & "</td>" & _
                    "</tr>" 
    			
			    intNo=intNo+1
			    rs.MoveNext
			    intI=intI+1
		loop

        strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		           "<td valign='top' colspan='4' class='blue' align='right'>Total: </td>"
        for jj=0 to UBound(dblSubtotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
        strReport= strReport & "<tr bgcolor='#FFE1E1'>" &_
       "<td valign='top' colspan='4' class='blue' align='right'>Overall Total:</td>"
        for jj=0 to UBound(dblGrandTotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblGrandTotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
	GenerateSumReportByProject=strReport
End Function


'--------------------------------------------------
' Preparing data
'--------------------------------------------------
strfromto = Request.QueryString("fromto")
strprintdate = Request.QueryString("printdate")

intReportType=cint(Request.QueryString("type"))

set rsSumHours=session("varSumHours")
strLast=""
if rsSumHours.recordcount>0 then 
    if intReportType=2 then
        strLast=GenerateSumReportByProject(rsSumHours)
    else
        strLast=GenerateSumReportByStaff(rsSumHours)	
    end if
end if

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
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Overtime")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strfromto)
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
			%>

        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                <tr> 
                  <%if intReportType=1 then%>                
                   <td class="blue" align="center" width="7%" bgcolor="#E7EBF5" rowspan="2">&nbsp;StaffID</td>
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Full Name</td>
                  <td class="blue" align="center" width="14%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Jobtitle</td>
                  <td class="blue" align="center" width="32%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Project Name</td>
                  
<%else %>                  
                  <td class="blue" align="center" width="32%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Project Name</td>
                  <td class="blue" align="center" width="7%" bgcolor="#E7EBF5" rowspan="2">&nbsp;StaffID</td>
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Full Name</td>
                  <td class="blue" align="center" width="14%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Jobtitle</td>
               
<%end if %>

                  
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5" colspan="2">Nomal day</td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5" colspan="2">Weekend</td>
                  <td class="blue" align="center" width="8%" bgcolor="#E7EBF5" rowspan="2">Total</td>                  
                </tr>
                <tr> 
                  <td class="blue" align="center" bgcolor="#E7EBF5">Normal</td>
                  <td class="blue" align="center" bgcolor="#E7EBF5">Night</td>
                  <td class="blue" align="center" bgcolor="#E7EBF5">Normal</td>
                  <td class="blue" align="center" bgcolor="#E7EBF5">Night</td>
                </tr>     
              
               
<%Response.Write strLast%>
                </tr>


              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>