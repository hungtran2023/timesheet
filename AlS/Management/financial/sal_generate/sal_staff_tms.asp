<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/library.asp"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->

<%
	Response.Buffer = True
	
	Dim intUserID, intMonth, intYear, intWeekday, intDayNum, intDayCol, intDayCount, intRow, eRow, intTotalRow, ii, kk, intCurMonth 
	Dim dblHour, dblTotal, strHour
	Dim strFirstDay, strParm, strURLSetHour, strColorOpt, strError, varTimesheet, varEvent
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	If Request.Form("M") = "" Then
		intMonth = Month(Date)
	Else
		intMonth = Request.Form("M")
	End If
	If Request.Form("Y") = "" Then
		intYear = Year(Date)
	Else
		intYear	= Request.Form("Y")
	End If		

	intCurMonth = Month(Date)
	strAction	= Request.QueryString("act")
	
	intRow		= -1
	eRow		= -1
	intDayNum	= GetDay(intMonth,intYear)				' Numbers of days in a month
	intDayCol	= intDayNum + 6

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../../message.htm")
	End If					

	intUserID	= session("USERID")
	intStaffID  = Request.Form("txthidden")
	
	strFirstDay = FirstOfMonth(intMonth,intYear)		' Get the first day in a month
	strLastDay = FirstOfMonth(intMonth,intYear) + (intDayNum -1)	' Get the last day in a month

'--------------------------------------------------
' The timesheet array initializing function is called 
' when session("varTimesheet")/session("varEvent") is not initialized
' or user changes month/year to view timesheet    
'--------------------------------------------------

	If Request.QueryString("act") = "" Then
		If Not IsEmpty(session("varTimesheet")) And Not IsEmpty(session("varEvent")) Then
			session("varTimesheet") = Empty
			session("varEvent") = Empty
		End If
	End If
			
	If (IsEmpty(session("varTimesheet")) And IsEmpty(session("varEvent"))) Or (Request.QueryString("act") = "vmya") Then
		
		strError	=  tmsInitial(intStaffID,intMonth,intYear)


		If strError = "" Then
			varTimesheet = session("varTimesheet")		' Array stores timesheet data
			varEvent	 = session("varEvent")			' Array stores event data
		Else
			varEvent	 = session("varEvent")	
		End If
	Else
		varTimesheet = session("varTimesheet")			' Array stores timesheet data
		varEvent	 = session("varEvent")				' Array stores event data
	End If

	call AnalyseOT(intMonth,intYear,varEvent)
	
	
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
	strFunction = "<a class='c' href='javascript:back_menu();' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view salary&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;&quot;;return true'>Preferences</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage()' onMouseOver='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"

	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Salary sheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../../../tms/templates/template1/")

%>	

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../../timesheet.css">

</head>

<script language="javascript" src="../../../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objNewWindow, objPrintWindow;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function viewtms()
{
	var URL;

	window.document.frmtms.M.value = window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value;
	window.document.frmtms.Y.value = window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value

	URL = "sal_staff_tms.asp?act=vmya";

	window.document.frmtms.action = URL;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function selstaff()
{
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=530,height=325,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('sal_select_staff.asp?view=t', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function logout()
{
	var url;
	url = "../../../logout.asp";
	if (ns)
		document.location = url;
	if (ie)
	{
		window.document.frmtms.action = url;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function back_menu()
{
	window.document.frmtms.action = "sal_list_staff.asp?b=1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function printpage()
{
	window.status = "";
	
	strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=450,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";

	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();	

	objPrintWindow = window.open('sal_print_preview.asp?m=' + '<%=intMonth%>' + '&y=' + '<%=intYear%>' + '&intStaffID=' + '<%=intStaffID%>', "MyNewWindow", strFeatures);
	objPrintWindow.focus();

	window.status = "Opened a new browser window.";  
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
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
<%
	If strError <> "" Then
%>
  <tr height="20">
    <td>
      <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td class="red" align="center">&nbsp;<b><%=strError%></b></td>
        </tr>
      </table>    
    </td>
  </tr>  
<%	End If%>      
  <tr> 
    <td valign="top">
      <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td bgcolor="#8FA4D3"> 
            <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
              <tr> 
                <td colspan="2" rowspan="2" class="white" bgcolor="#617DC0"> 
                  <div align="center"> <b>Project </b> </div>
                </td>
                <td colspan="<%=intDayNum%>" class="blue-normal" align="right" bgColor="#617DC0"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="blue-normal">
                    <tr> 
                      <td width="57%" class="white">&nbsp;&nbsp;<%=strTitle1%></td>
                      <td align="right" width="35%">
					    <select name="lbyear" size="1" class="blue-normal">
						<%For ii=2000 To Year(Date)%>
					      <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
						<%Next%>
						</select>
						<select name="lbmonth" size="1" class="blue-normal">
						<%For ii=1 To 12%>
					      <option <%If CInt(intMonth)=ii Then%>selected<%End If%> value="<%=ii%>"><%=SayMonth(ii)%></option>
						<%Next%>
						</select>
                      </td>
                      <td width="8%" align="right"> 
                        <table border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa" width="40">
                          <tr> 
                            <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
                              <div align="center" class="blue"><a href="javascript:viewtms();" class="b">Go</a></div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td rowspan="2" class="white" bgcolor="#617DC0"> 
                  <div align="center"><b>Total</b></div>
                </td>
              </tr>
              <tr bgcolor="#617DC0">
<%				For kk=1 To intDayNum
					intWeekDay = WeekDay(strFirstDay+(kk-1))%>
			        <td width="19"><div align="center" class=<%If(intWeekDay=1 or intWeekDay=7)then%>"holiday"<%else%>"white"<%end if%>><b><%=kk%></b></div></td>
<% 				Next%>  
			  </tr>
					  
<!--**************************** For Project And SubTask *********************************************-->
					  
<%
	intTotalRow = intRow
	If intTotalRow <= 5 Then intTotalRow = 5
	
	For ii = 0 To intTotalRow
		If ii <= intRow Then
			If varTimesheet(intDayCol-1,0,ii) = 0 Then
%>					  	
              <tr> 
                <td width="8"><img src="../../../images/cross.gif" width="8" height="14"></td>
<%
				' This row stores subtask name
				If trim(varTimesheet(intDayCol-2,0,ii)) = "S" Then
					strTitle=varTimesheet(intDayCol-3,0,ii)
					strLable=showlabel(varTimesheet(intDayCol-3,0,ii))
				else
					strTitle=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)
					strLable=showlabel(varTimesheet(0,0,ii))
				end if
%>                        
					<td width="118" class="blue" bgcolor="#FFF2F2">
						<a href="javascript:void(0);" title="<%=strTitle%>" onMouseOver="self.status=&quot;<%=strTitle%>&quot;;return true" onMouseOut="self.status='';return true" class="c">
								<b>&nbsp;<%=strLable%></b>
						</a>
					</td>
				
<%				For kk = 1 To intDayNum
					dblHour = varTimesheet(kk, 0, ii) + varTimesheet(kk, 1, ii)
					strHour = "&nbsp;"					
					If dblHour > 0 Then	strHour = formatnumber(dblHour,1)
					
					intWeekDay = WeekDay(strFirstDay+(kk-1))
					strColorOpt = "#FFFFFF"
					Select Case intWeekDay
						Case 1
							strColorOpt = SUNCOLOR
						Case 7
							strColorOpt = SATCOLOR
					End Select
					If isHoliday(kk) >= 0 Then
						strColorOpt = HOLIDAYCOLOR
					End If	
	%>                        
                <td bgcolor="<%=strColorOpt%>" align="center" class="blue-normal" width="19"><%=strHour%></td>
<%
				Next
				
				If CDbl(varTimesheet(intDayCol-5, 0, ii)) > 0 Then
					dblTotal = formatnumber(varTimesheet(intDayCol-5, 0, ii),1)
				Else
					dblTotal = "&nbsp;"
				End If		
%>  
                <td bgcolor="#FFF2F2" align="right" class="blue"><%=dblTotal%>&nbsp;</td>
              </tr>
<%
			End If
		Else
%>                      
              <tr> 
                <td width="8" bgcolor="#FFC6C6" class="white">&nbsp;</td>
                <td width="118" bgcolor="#FFF2F2" class="blue-normal">&nbsp;</td>
<%
			For kk = 1 To intDayNum
				intWeekDay = WeekDay(strFirstDay+(kk-1))
				strColorOpt = ""
				Select Case intWeekDay
					Case 1
						strColorOpt = SUNCOLOR
					Case 7
						strColorOpt = SATCOLOR
				End Select
				If isHoliday(kk) >= 0 Then
					strColorOpt = HOLIDAYCOLOR
				End If	
%>                        
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal" width="19">&nbsp;</td>
<%
			Next
%>			
                <td bgcolor="#FFF2F2" align="right" class="blue-normal">&nbsp;</td>
              </tr>
<%
		End If
	Next	
%>           

<!--**************************** End Of Project And SubTask *******************************-->
  
<!--**************************** For Events and Others *******************************************-->
              <tr>
<%
	For ii = 0 To eRow
		If varEvent(intDayNum+2,0,ii) = -1 Or varEvent(intDayNum+2,0,ii) = -2 Or varEvent(intDayNum+2,0,ii) = -3 Then
%>
                <td colspan="2" bgcolor="#FFE1E1" class="blue"><b>&nbsp;<%=varEvent(0,0,ii)%></b></td>
<%
		Else
%>          
                <td colspan="2" class="blue-normal" bgcolor="#FFE1E1">&nbsp;<%=varEvent(0,0,ii)%></td>
<%
		End If

		For kk =1 To intDayNum
			dblHour = varEvent(kk, 0, ii) + varEvent(kk, 1, ii)
			strHour = "&nbsp;"
				
			If dblHour > 0 Then
				strHour = formatnumber(dblHour,1)
			End If	

			intWeekDay = WeekDay(strFirstDay+(kk-1))
			strColorOpt = "#E7EBF5"
			Select Case intWeekDay
				Case 1
					strColorOpt = SUNCOLOR
				Case 7
'					strColorOpt = SATCOLOR
					strColorOpt = "#D2DAEC"
			End Select
			If isHoliday(kk) >= 0 Then
				strColorOpt = HOLIDAYCOLOR
			End If	
%>                        
                <td bgcolor="<%=strColorOpt%>" align="center" class="blue-normal" width="19"><%=strHour%></td>
<%
		Next
		If varEvent(intDayNum+1, 0, ii) > 0 Then
			dblTotal = formatnumber(varEvent(intDayNum+1, 0, ii),1)
		Else
			dblTotal = "&nbsp;"
		End If		
%>
				<td bgcolor="#FFE1E1" align="right" class="blue"><%=dblTotal%>&nbsp;</td>
              </tr> 
<%
	Next
%>                      
<!--**************************** End Of Events and Others *********************************************-->
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--**************************** Salary payment detail *********************************************-->
<%
	arrSalaryStatus =GetSalaryStatus(intStaffID,strFirstDay,strLastDay)
	
	if not IsEmpty(arrSalaryStatus) then
%>
		<table width="780" cellspacing="0" cellpadding="0" align="center">
		<tr>
		<td>
		<%
			strTemplate=GenerateSalary(intStaffID,strFirstDay,strLastDay,arrSalaryStatus,dblGrantBasic,dblGrantOT)
			Response.Write(strTemplate)
			
			call Write_summary("templates/template1/",dblGrantBasic,dblGrantOT)

		%>
		</td>
		</tr>
		</table>
	<%End if%>  
<!--**************************** Footer *********************************************-->
  
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="txthidden" value="<%=intStaffID%>">
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">


</form>
</body>
</html>
