<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/library.asp"-->
<!-- #include file = "../inc/getmenu.asp"-->

<%
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim intUserID, intMonth, intYear, intWeekday, intDayNum, intDayCount, intRow, eRow, intTotalRow, ii, kk 
	Dim dblOffHour, dblOverHour, dblOffTotal, dblOverTotal
	Dim strFirstDay, varTimesheet, varEvent

'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	intMonth = Request.Form("lbmonth")
	intYear	 = Request.Form("lbyear")

	intRow		= -1
	eRow		= -1
	intDayNum	= GetDay(intMonth,intYear)				' Numbers of days in a month

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If					

	intUserID	= session("USERID")
	strFirstDay = FirstOfMonth(intMonth,intYear)		' Get the first day in a month				
	intDayCount	= curDayNum(intDayNum,strFirstDay)		' Numbers of days since the first day in month to now

	varTimesheet = session("varTimesheet")				' Array stores timesheet data
	varEvent	 = session("varEvent")					' Array stores event data
	
	If isarray(varTimesheet) Then
		intRow	= Ubound(varTimesheet,3)
	End If
	
	If isarray(varEvent) Then
		eRow	= Ubound(varEvent,3)
	End If

'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='../welcome.asp' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:goback()' onMouseOver='self.status=&quot;Back&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Back</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
  				  "<a class='c' href='javascript:viewleave()' onMouseOver='self.status=&quot;View annual leave of staff&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>View Leave</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage();' onMouseOver='self.status=&quot;Print timesheet detail page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "templates/template1/")
%>	

<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Atlas Industries - Timesheet Detail</title>

<link rel="stylesheet" href="../timesheet.css">

</head>

<script language="javascript" src="../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objNewWindow;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	var url;
	url = "../logout.asp";
	if (ns)
		document.location = url;
	else
	{
		window.document.frmtms.action = url;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function gopage()
{
	document.frmtms.action = "../tools/preferences.asp";
	document.frmtms.submit();
}

function printpage()
{
	window.status = "";
 
		strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=680,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('tms_detail_print.asp?m=' + '<%=intMonth%>' + '&y=' + '<%=intYear%>', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function goback()
{
	if (ns)
		document.location = "timesheet.asp?act=vpa";
	else
	{
		window.document.frmtms.action = "timesheet.asp?act=vpa";
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}


function viewleave()
{
	if (ns)
		document.location = "staff_view_leave.asp";
	else
	{
		window.document.frmtms.action = "staff_view_leave.asp";
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmtms" method="post">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
%>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="445" style="height:&quot;76%&quot;" align="center">
  <tr> 
    <td bgcolor="#FFFFFF" valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="blue" width="19%" valign="middle">&nbsp; </td>
          <td class="blue-normal" align="right" width="42%" valign="middle">&nbsp;</td>
          <td align="right" width="18%" valign="middle">&nbsp; </td>
          <td class="blue" align="right" width="21%" valign="middle">&nbsp; </td>
        </tr>
        <tr> 
          <td class="title" height="50" align="center" colspan="4">Timesheet Details<br><div class="blue-normal"><%=SayMonth(intMonth)%>, <%=intYear%></div></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#617DC0"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="5">
              <tr> 
                <td class="blue" align="center" width="50" bgcolor="#E7EBF5">Date</td>
                <td class="blue" align="center" width="305" bgcolor="#E7EBF5">Project - Event</td>
                <td class="blue" align="center" width="53" bgcolor="#E7EBF5">Normal Hours</td>
                <td class="blue" align="center" width="51" bgcolor="#E7EBF5">Overtime</td>
                <td class="blue" align="center" width="265" bgcolor="#E7EBF5">Comment</td>
              </tr>
<%
	dblOffTotal  = 0
	dblOverTotal = 0

	If intDayCount >= 1 Then
		For kk = 1 to intDayCount
			dblOffHour	= "&nbsp;"
			dblOverHour = "&nbsp;"
			strNotes	= "&nbsp;"
			fgShow = True

			strDateCheck = Cstr(intMonth) & "/" & Cstr(kk) & "/" & Cstr(intYear)
			strDateShow = Cstr(kk) & "/" & SayMonth(intMonth) & "/" & Cstr(intYear)
			
			intHoliday = isHoliday(kk)
			If intHoliday >= 0 Then
				strDateShow = strDateShow & " - Public Holiday"
			Else
				Select Case Weekday(CDate(strDateCheck))
					Case 1
						strDateShow = strDateShow & " - Sunday"
					Case 2
						strDateShow = strDateShow & " - Monday"
					Case 3
						strDateShow = strDateShow & " - Tuesday"
					Case 4
						strDateShow = strDateShow & " - Wednesday"
					Case 5
						strDateShow = strDateShow & " - Thursday"
					Case 6
						strDateShow = strDateShow & " - Friday"
					Case 7	
						strDateShow = strDateShow & " - Saturday"
				End Select		
			End If

			For ii = 0 To intRow
				dblOverHour = "&nbsp;"
				dblOffHour = "&nbsp;"
				If (CDbl(varTimesheet(kk, 0, ii)) + CDbl(varTimesheet(kk, 1, ii))) > 0 Then				
					
					dblOffTotal = dblOffTotal + CDbl(varTimesheet(kk, 0, ii))
					dblOverTotal = dblOverTotal + CDbl(varTimesheet(kk, 1, ii))
					
					If CDbl(varTimesheet(kk, 0, ii)) > 0 Then
						dblOffHour = formatnumber(varTimesheet(kk, 0, ii),1)					
					End If
					If CDbl(varTimesheet(kk, 1, ii)) > 0 Then
						dblOverHour = formatnumber(varTimesheet(kk, 1, ii),1)
					End If	
					strNotes = showlabel(trim(varTimesheet(kk, 4, ii)))

					strTitle = trim(varTimesheet(0, 0, ii)) & " (" & trim(varTimesheet(intDayNum+3,0,ii)) & ")"

					If fgShow Then
						fgShow = False
%>              
              <tr> 
                <td valign="top" colspan="5" class="blue" bgcolor="#FFF2F2"><%=strDateShow%></td>
              </tr>
<%
					End If
%>              
              <tr> 
                <td valign="top" colspan="2" class="blue-normal" bgcolor="#FFFFFF" align="left">&nbsp;<%=strTitle%></td>
                <td valign="top" width="53" class="blue-normal" bgcolor="#FFFFFF" align="right"><%=dblOffHour%>&nbsp;</td>
                <td valign="top" width="51" class="blue-normal" align="right" bgcolor="#FFFFFF"><%=dblOverHour%>&nbsp;</td>
                <td valign="top" width="265" class="blue-normal" bgcolor="#FFFFFF">&nbsp;<%=strNotes%></td>
              </tr>
<%
				End If
			Next
			
			NumRow = eRow-3
			For ii = 0 To NumRow
				dblOverHour = "&nbsp;"
				dblOffHour = "&nbsp;"
				If (CDbl(varEvent(kk, 0, ii)) + CDbl(varEvent(kk, 1, ii))) > 0 Then				
					
					dblOffTotal  = dblOffTotal + CDbl(varEvent(kk, 0, ii))
					dblOverTotal = dblOverTotal + CDbl(varEvent(kk, 1, ii))
					
					If CDbl(varEvent(kk, 0, ii)) > 0 Then
						dblOffHour = formatnumber(varEvent(kk, 0, ii),1)					
					End If
					If CDbl(varEvent(kk, 1, ii)) > 0 Then
						dblOverHour = formatnumber(varEvent(kk, 1, ii),1)
					End If	
					strNotes = showlabel(trim(varEvent(kk, 4, ii)))
					strTitle = trim(varEvent(0, 0, ii))
					
					If fgShow Then
						fgShow = False
%>           
              <tr> 
                <td valign="top" colspan="5" class="blue" bgcolor="#FFF2F2"><%=strDateShow%></td>
              </tr>
<%
					End If
%>
              <tr> 
                <td valign="top" colspan="2" class="blue-normal" bgcolor="#FFFFFF" align="left">&nbsp;<%=strTitle%></td>
                <td valign="top" width="53" class="blue-normal" bgcolor="#FFFFFF" align="right"><%=dblOffHour%>&nbsp;</td>
                <td valign="top" width="51" class="blue-normal" align="right" bgcolor="#FFFFFF"><%=dblOverHour%>&nbsp;</td>
                <td valign="top" width="265" class="blue-normal" bgcolor="#FFFFFF">&nbsp;<%=strNotes%></td>
              </tr>
<%
				End If
			Next
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
<!--<table width="780" border="0" cellspacing="0" cellpadding="0" height="20" align="center">
  <tr> 
    <td align="right" bgcolor="#E7EBF5"> 
      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
        <tr class="black-normal"> 
          <td align="right" valign="middle" width="37%" class="blue-normal">Page</td>
          <td align="center" valign="middle" width="13%" class="blue-normal"> 
            <input type="text" name="txtpage" class="blue-normal" value="1" size="2" style="width:50">
          </td>
          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="tms_viewdetails.asp#"><font color="#990000">Go</font></a></td>
          <td align="right" valign="middle" width="15%" class="blue-normal">Pages 1/10&nbsp;&nbsp;</td>
          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="tms_viewdetails.asp#">Previous Page</a> /<a href="tms_viewdetails.asp#"> Next</a>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>-->
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
</form>
</body>
</html>