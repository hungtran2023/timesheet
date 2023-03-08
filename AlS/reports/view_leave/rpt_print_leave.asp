<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
	
	Response.Buffer = True
	
	Dim intUserID, intMonth, intYear, intLeaveDueCur, intAppCur, intBalance
	Dim decBalancePast, dateExipred, intWorkHours
	Dim strConnect, objDatabase, strError

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	intMonth = Request.Form("M")
	intYear = Request.Form("Y")
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If	

	intUserID	= session("USERID")
	
	Set rsAnnual = session("rsAnnual")
	rsAnnual.Sort = "FullName ASC"
	rsAnnual.MoveFirst
	

%>	
<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmtms" method="post">

		<table style="TABLE-LAYOUT: fixed;" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
		  <tr> 
		    <td class="title" height="20" align="center">&nbsp;&nbsp;</td>
		  </tr>
		  <tr> 
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
			    <tr> 
				  <td class="title" height="50" align="center">View Annual Leave until &nbsp 1/<%=month(date)%>/<%=year(date)%></td>
			    </tr>
			    <tr> 
		          <td class="title" height="20" align="center">&nbsp;&nbsp;</td>
				</tr>
              </table>
		    </td>
		  </tr>
		  <tr> 
		    <td> 
			  <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365">
				<tr>
					<td align="right" class="blue-normal"><b><%=formatdatetime(date,vbLongDate)%>&nbsp;</b></td>
				</tr>
			    <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#003399">
                      <tr> 
                        <td bgcolor="#8FA4D3"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5" align="center">
                            <tr bgcolor="#617DC0" height="25"> 
                                <td class="white" align="center" width="15%"><b>StaffID </b></td>
                              <td class="white" align="center" width="35%"><b>FullName </b></td>
                              <td class="white" align="center" width="35%"><b>Job Title </b></td>
                              <td class="white" align="center" width="15%"><b>Balance (days) </b></td>
					        </tr>
<%
		Do While Not rsAnnual.EOF
		
			
			'if rsStaff("StaffID")=248 then Response.Write date & "-" & intBalance & "-" & myCmd("expiredDate")
%>		
					        <tr bgcolor="#E7EBF5" height="25"> 
					            <td valign="middle"  class="blue-normal"><%=rsAnnual("StaffID")%></td>
					          <td valign="middle" class="blue-normal"><%=rsAnnual("FullName")%></td>
	                          <td valign="middle" class="blue-normal"><%=rsAnnual("JobTitle")%></td>
		                      <td valign="middle" class="blue-normal" align="center"><%=formatnumber(rsAnnual("Balance"),2)%></td>
			                </tr>
<%

			rsAnnual.MoveNext
		Loop
							
%>			                
				          </table>
				        </td>
			          </tr>
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
