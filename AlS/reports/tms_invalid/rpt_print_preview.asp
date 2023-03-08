<%
	Dim ii, intRow
	Dim varUser
	Dim strTitle
	
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	strTitle = Request.QueryString("title")
	varUser   = session("varInvalidTMS")
	If isArray(varUser) Then
		intRow = Ubound(varUser,2)
	End If	 		
%>
<html>
<head>
<title>Atlas industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="1" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript">
<form name="frmreport" method="post">

<table width="610" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr align="center"> 
          <td class="title" height="50" align="center" colspan="2"> Missing Timesheet<br>
            <div class="blue-normal"><%=strTitle%></div>
          </td>
        </tr>
        <tr align="right"> 
          <td class="blue-normal" height="20" colspan="2">Printing Date: <%=formatdatetime(date,vbLongDate)%>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365">
        <tr> 
          <td bgcolor="#FFFFFF" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#617DC0"> 
                  <table width="100%" border="0" cellspacing="1" cellpadding="5">
                    <tr> 
                      <td class="blue" align="center" width="101" bgcolor="#E7EBF5">Date</td>
                      <td class="blue" align="center" width="229" bgcolor="#E7EBF5">Full Name </td>
                      <td class="blue" align="center" width="214" bgcolor="#E7EBF5">Report To </td>
                      <td class="blue" align="center" width="18" bgcolor="#E7EBF5">&nbsp;</td>
                    </tr>
<%                    
	If intRow >= 0 Then
		For ii = 0 To intRow
			If varUser(1,ii) = "Date" Then
%>					
                    <tr> 
                      <td valign="top" colspan="4" class="blue" bgcolor="#FFF2F2"><%=varUser(0,ii)%></td>
                    </tr>
<%
			Else
%>                    
                    <tr> 
                      <td valign="top" width="101" class="blue" align="right" bgcolor="#FFFFFF"><%=varUser(0,ii)%></td>
                      <td valign="top" width="229" class="blue-normal" bgcolor="#FFFFFF">&nbsp;<%=varUser(2,ii)%></td>
                      <td valign="top" width="214" class="blue-normal" align="center" bgcolor="#FFFFFF">&nbsp;<%=varUser(4,ii)%></td>
                      <td valign="top" width="18" class="blue-normal" align="center" bgcolor="#FFFFFF">&nbsp; 
<!--                        <input type="checkbox" name="chkremind" value="<%=varUser(0,ii)%>">-->
                      </td>
                    </tr>
<%
			End If
		Next
%>
                  </table>
<%
	End If
	
	Set varUser = Nothing
%>                  
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
