<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<%
'--------------------------------------------------
' Preparing data
'--------------------------------------------------
strprintdate = Request.QueryString("printdate")
strFromTo=Request.QueryString("fromto")
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
<HTML>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css">
<script language="javascript" src="../library/library.js"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" height="445" style=height:"76%"  align="center" >
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
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "KPI Report")
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strFromTo)
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", strprintdate)
			Response.Write(arrPageTemplate(1))
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
		    
			    
<%=session("rpt_KPI")%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</BODY>
</HTML>
