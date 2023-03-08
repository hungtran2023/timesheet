<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	Response.Buffer = True
	
	Dim intUserID, intMonth, intYear, intWeekday, intDayNum, intDayCol, intDayCount, intRow, eRow, intTotalRow, ii, kk, intCurMonth 
	Dim dblHour, dblTotal, strHour
	Dim strFirstDay, strParm, strURLSetHour, strColorOpt, strError, varTimesheet, varEvent,arrFingerprint
'***************************************************************
'
'***************************************************************
sub GetFingerprintByStaff()
	
	dim ii, intColumns
	dim objConn
	dim rs

	strOut=""
	strConnect = Application("g_strConnect")
	set objConn= Server.CreateObject("ADODB.Connection")
	objConn.Open strConnect  
	
	If objConn.State=1 Then

		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objConn
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "FingerprintByStaff"
		Set myParam = myCmd.CreateParameter("month",adInteger,adParamInput)
		myCmd.Parameters.Append myParam		
		Set myParam = myCmd.CreateParameter("year",adInteger,adParamInput)
		myCmd.Parameters.Append myParam

		Set myParam = myCmd.CreateParameter("staffID",adNumeric,adParamInput)
		
		myParam.Precision = 18
		myParam.NumericScale = 0


		myCmd.Parameters.Append myParam
		myCmd("month") = intMonth
		myCmd("year") = intyear
		myCmd("staffID") = intStaffID
	  
		SET rs=myCmd.Execute
		arrFingerprint=rs.GetRows()
		

	end if
	
End sub
'--------------------------------------------------
'
'--------------------------------------------------	
Function Format_hhmm(dblMinuteIn)
	dim strOut
	dim dblHour, dblMinute
	strOut=""
	if dblMinuteIn<>0 then
		dblHour=cint(abs(dblMinuteIn)\60)
		dblMinute=abs(dblMinuteIn) mod 60
		strOut=dblHour
		if dblHour<10 then strOut="0" & strOut
		strOut=strOut &":" 
		if dblMinute<10 then
			strOut=strOut & "0" & dblMinute
		else
			strOut=strOut & dblMinute
		end if
		if dblMinuteIn<0 then strOut="-" & strOut
	end if
	
	Format_hhmm=strOut
end function	
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
		Response.Redirect("../message.htm")
	End If					

	intUserID	= session("USERID")
	intStaffID  = Request.Form("txthidden")
	strFirstDay = FirstOfMonth(intMonth,intYear)		' Get the first day in a month				
	intDayCount	= curDayNum(intDayNum,strFirstDay)		' Numbers of days since the first day in month to now

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
	strFunction = "<a class='c' href='javascript:back_menu();' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view timesheet&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:viewdetail()' onMouseOver='self.status=&quot;View timesheet detail&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>View Detail</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
  				  "<a class='c' href='javascript:viewleave()' onMouseOver='self.status=&quot;View annual leave of staff&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>View Leave</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage()' onMouseOver='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"

	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Timesheet of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "templates/template1/")

%>	

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<script language="javascript" src="../../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objStaffWindow, objPrintWindow;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function viewtms()
{
	var URL;

	window.document.frmtms.M.value = window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value;
	window.document.frmtms.Y.value = window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value

	URL = "rpt_staff_tms.asp?act=vmya";

	window.document.frmtms.action = URL;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function selstaff()
{
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=530,height=325,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objStaffWindow) && (!objStaffWindow.closed))
		objStaffWindow.focus();	
	else 
	{
		objStaffWindow = window.open('rpt_select_staff.asp?view=t', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function logout()
{
	var url;
	url = "../../logout.asp";
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
	document.frmtms.action = "../../tools/preferences.asp";
	document.frmtms.submit();
}

function back_menu()
{
	window.document.frmtms.action = "rpt_list_staff.asp?b=1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function viewdetail()
{
	window.document.frmtms.action = "rpt_tms_detail.asp";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function viewleave()
{
	//alert("This page is fixing error.\n\n Please view Annual leave at daily timehseet page or contact HR for helping.\n\n Thank you,\n Uyen Chi")
	if (ns)
		document.location = "staff_view_leave.asp";
	else
	{
		window.document.frmtms.action = "staff_view_leave.asp";
		//window.document.frmtms.action = "rpt_tms_leave.asp";
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function printpage()
{
	window.status = "";
	
	strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=450,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";

	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();	
	
	objPrintWindow = window.open('rpt_print_preview.asp?m=' + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + '&y=' + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value + '&s=' + '<%=intStaffID%>', "MyNewWindow", strFeatures);
	objPrintWindow.focus();

	window.status = "Opened a new browser window.";  
}

function window_onunload() 
{
	if((objStaffWindow) && (!objStaffWindow.closed))
		objStaffWindow.close();
		
	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" language="javascript" onUnload="return window_onunload();">
<form name="frmtms" method="post">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
%>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
	If strError <> "" Then
%>
  <tr height="20">
    <td>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td class="red" align="center">&nbsp;<b><%=strError%></b></td>
        </tr>
      </table>    
    </td>
  </tr>  
<%	End If%>      
  <tr> 
    <td valign="top">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td bgcolor="#8FA4D3"> 
            <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
              <tr> 
                <td colspan="2" rowspan="2" class="white" bgcolor="#617DC0"> 
                  <div align="center"> <b>Project </b> </div>
                </td>
                <td width="60%"colspan="<%=intDayNum%>" class="blue-normal" align="right" bgColor="#617DC0"> 
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
						  <option <%If CInt(intMonth)=1 Then%>selected<%End If%> value="1">January</option>
						  <option <%If CInt(intMonth)=2 Then%>selected<%End If%> value="2">February</option>
						  <option <%If CInt(intMonth)=3 Then%>selected<%End If%> value="3">March</option>
						  <option <%If CInt(intMonth)=4 Then%>selected<%End If%> value="4">April</option>
						  <option <%If CInt(intMonth)=5 Then%>selected<%End If%> value="5">May</option>
						  <option <%If CInt(intMonth)=6 Then%>selected<%End If%> value="6">June</option>
						  <option <%If CInt(intMonth)=7 Then%>selected<%End If%> value="7">July</option>
						  <option <%If CInt(intMonth)=8 Then%>selected<%End If%> value="8">August</option>
						  <option <%If CInt(intMonth)=9 Then%>selected<%End If%> value="9">September</option>
						  <option <%If CInt(intMonth)=10 Then%>selected<%End If%> value="10">October</option>
						  <option <%If CInt(intMonth)=11 Then%>selected<%End If%> value="11">November</option>
						  <option <%If CInt(intMonth)=12 Then%>selected<%End If%> value="12">December</option>
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
<%
				For kk=1 To intDayNum
					intWeekDay = WeekDay(strFirstDay+(kk-1))
					strColorOpt = ""
					Select Case intWeekDay
						Case 1
							strColorOpt = DAYCOLOR
						Case 7
							strColorOpt = DAYCOLOR
					End Select
					If Not strColorOpt="" Then%>
				        <td width="2%"><div align="center" class="white"><font color="<%=strColorOpt%>"><b><%=kk%></b></font></div></td>
<%					Else%>
   				        <td width="2%"><div align="center" class="white"><b><%=kk%></b></div></td>
<%					End If
   				Next%>  
			  </tr>
					  
<!--**************************** For Project And SubTask *********************************************-->
					  
<%
	intTotalRow = intRow
	If intTotalRow <= 9 Then
		intTotalRow = 9
	End If	
	For ii = 0 To intTotalRow
		If ii <= intRow Then
			If varTimesheet(intDayCol-1,0,ii) = 0 Then
%>					  	
              <tr> 
                <td height="20" width="8"><img src="../../images/cross.gif" width="8" height="14"></td>
<%
				If trim(varTimesheet(intDayCol-2,0,ii)) = "S" Then
%>                        
                <td class="blue" bgcolor="#FFF2F2"><a href="javascript:void(0);" title="<%=varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;-<%=showlabel(varTimesheet(intDayCol-3,0,ii))%></b></a></td>
<%
				ElseIf trim(varTimesheet(intDayCol-2,0,ii)) = "N" Then
%>                        
                <td class="blue" bgcolor="#FFF2F2"><a href="javascript:void(0)" title="<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=showlabel(varTimesheet(0,0,ii))%></b></a></td>
<%	
				ElseIf trim(varTimesheet(intDayCol-2,0,ii)) = "P" Then
%>                        
                <td  class="blue" bgcolor="#FFF2F2"><a href="javascript:void(0);" title="<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=showlabel(varTimesheet(0,0,ii))%></b></a></td>
<%
				End If
				For kk = 1 To intDayNum
					dblHour = varTimesheet(kk, 0, ii) + varTimesheet(kk, 1, ii)
					strHour = "&nbsp;"
					
					If dblHour > 0 Then
						strHour = formatnumber(dblHour,1)
					End If	
					
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
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal"><%=strHour%></td>
<%
				Next
				If CDbl(varTimesheet(intDayCol-5, 0, ii)) > 0 Then
					dblTotal = formatnumber(varTimesheet(intDayCol-5, 0, ii),1)
				Else
					dblTotal = "&nbsp;"
				End If		
%>  
                <td bgcolor="#FFF2F2" align="right" class="blue"><b><%=dblTotal%></b>&nbsp;</td>
              </tr>
<%
			End If
		Else
%>                      
              <tr> 
                <td height="20" width="8" bgcolor="#FFC6C6" class="white">&nbsp;</td>
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
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal">&nbsp;</td>
<%
			Next
%>			
                <td height="20" bgcolor="#FFF2F2" align="right" class="blue-normal">&nbsp;</td>
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
                <td height="20" colspan="2" bgcolor="#FFE1E1" class="blue"><b>&nbsp;<%=varEvent(0,0,ii)%></b></td>
<%
		Else
%>          
                <td height="20" colspan="2" class="blue-normal" bgcolor="#FFE1E1">&nbsp;<%=varEvent(0,0,ii)%></td>
<%
		End If

		For kk =1 To intDayNum
			dblHour = varEvent(kk, 0, ii) + varEvent(kk, 1, ii)
			strHour = "&nbsp;"
				
			If dblHour > 0 Then
				strHour = formatnumber(dblHour,1)
			End If	

			intWeekDay = WeekDay(strFirstDay+(kk-1))
			strColorOpt = ""
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
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#E7EBF5" <%End If%> align="center" class="blue-normal" ><%=strHour%></td>
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
<!--**************************** End Of Events and Others *********************************************-->

<!--**************************** For Timekeeping *******************************************-->			
<%	call GetFingerprintByStaff()
	arrTimekeeping=Array("Start Time (hh:mm)","Finish Time (hh:mm)", "Duration (hours)", "<b>Daily Balance</b> (hours)", "<b>Weekly Balance</b> (hours)") 
		intColspan=0
		dblWeeklyBalance=0
		For ii=0 to UBound (arrTimekeeping)%>
			<tr>
					<td colspan="2" bgcolor="#ffc6c6" class="blue" height="20">&nbsp;<%=arrTimekeeping(ii)%></td>
					
<%
			jj=0
			
			Do while (month(arrFingerprint(0,jj))<>intMonth AND day(arrFingerprint(0,jj))<>1)
				if ii>3 then
					dblDailyBalance =cdbl(arrFingerprint(3,jj))- cdbl(arrFingerprint(4,jj))
					dblWeeklyBalance=dblWeeklyBalance + dblDailyBalance
				end if
				
'response.write  dblWeeklyBalance & "<br>"
				jj=jj+1
				
			loop

			For kk =1 To intDayNum
				intWeekDay = WeekDay(strFirstDay+(kk-1))				
				strColorOpt = ""
				Select Case intWeekDay
					Case 1
						strColorOpt = SUNCOLOR
					Case 7
						strColorOpt = "#D2DAEC"
				End Select
				If isHoliday(kk) >= 0 Then
					strColorOpt = HOLIDAYCOLOR
				End If	
%>							
<% if ii<=3 then %>
					<td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#E7EBF5" <%End If%> align="center" class="blue-normal">
						<%if ii<2 then
								response.write arrFingerprint(ii+1,jj)
						else
								if ii=2 then
									response.write Format_hhmm(cdbl(arrFingerprint(ii+1,jj)))
								else
									dblDailyBalance =cdbl(arrFingerprint(ii,jj))- cdbl(arrFingerprint(ii+1,jj))
									response.write Format_hhmm(dblDailyBalance)
									
									
								end if
						end if
						%>
					</td>
<%else
		intColspan=intColspan+1		
		dblWeeklyBalance =dblWeeklyBalance +(cdbl(arrFingerprint(3,jj))- cdbl(arrFingerprint(4,jj)))
'response.write  dblWeeklyBalance & "<br>"		
		if intWeekDay=6 or kk= intDayNum then 
%>		
			<td bgcolor="#E7EBF5" align="right" class="blue" colspan="<%=intColspan%>"><%=Format_hhmm(dblWeeklyBalance)%>&nbsp;
			</td>
			
<%		intColspan=0
		dblWeeklyBalance=0
		end if
end if
			
				jj=jj+1
			Next	%>
				<td bgcolor="#FFE1E1" align="right" class="blue">&nbsp;</td>
			</tr>		
<%		next
%>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  
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
