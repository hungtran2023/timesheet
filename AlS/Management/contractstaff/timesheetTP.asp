<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/libraryForTPTimesheet.asp"-->


<%
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	Dim intUserID, intMonth, intYear, intWeekday, intDayNum, intDayCol, intDayCount, intRow, eRow, intTotalRow, ii, kk, intCurMonth 
	Dim dblHour, dblTotal, strHour
	Dim strFirstDay, strParm, strURLSetHour, strColorOpt, strError, strShow, varTimesheet, varEvent,dLeaveDate
	Dim strDateLock

	strDateLock=Cdate("31-Jan-2015")
	dateLimit=cint(150)
	
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
		Response.Redirect("../../message.htm")
	End If		
	
				
	'blnAlert=false
	intUserID	= session("USERID")
	strTPUserid=Request.Form("txtUserid")
	
	strSql="SELECT Fullname,Username FROM HR_TPStaff WHERE TPUserid=" & strTPUserid
	Call GetRecordset(strSql,rsTP)

	if not rsTP.EOF then
	    strTitle1	= "Timesheet of <b>" & rsTP("Fullname") & " (" & rsTP("Username") & ")"
	end if
	
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
		'blnAlert=true
	End If
	
	If Request.QueryString("act") = "vpae" Then
		strError = "You can't remove this task, because it has data."
		Response.Cookies("introw") = ""
		Response.Cookies("assignid") = ""
	ElseIf Request.QueryString("act") = "vpa" Then	
		Response.Cookies("introw") = ""
		Response.Cookies("assignid") = ""
	End if
			
	If (IsEmpty(session("varTimesheet")) And IsEmpty(session("varEvent"))) Or (Request.QueryString("act") = "vmya") Then

		strError	=  tmsInitialForTP(strTPUserid,intMonth,intYear)
	
		If strError = "" Then
			varTimesheet = session("varTimesheet")		' Array stores timesheet data
			varEvent	 = session("varEvent")			' Array stores event data
		Else
			varTimesheet = Empty	
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
	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b><br>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:viewdetail()' onMouseOver='self.status=&quot;View timesheet detail&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>View Detail</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage()' onMouseOver='self.status=&quot;Print timesheet page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='../../logout.asp' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
'"<a class='c' href='javascript:gopage();' onMouseOver='self.status=&quot;Preferences&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Preferences</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
    
Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "../tms/templates/template1/")

%>	

<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<script type="text/javascript" src="../../library/library.js"></script>
<script type="text/javascript" src="../../library/menu.js"></script>

<script type="text/javascript">
<!--
var ns, ie, assignid;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function onLoad()
{
        loadMenus();
}

function loadMenus() 
{
var url_1 = "tms_removetask.asp?m=" + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + "&y=" + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value;

    window.myMenu1 = new Menu();

    myMenu1.addMenuItem("Update","timesheettp.asp?act=U", "", "", "", "frmtms");
    if ("<%=intMonth%>" == "<%=intCurMonth%>")
	{
		myMenu1.addMenuItem("Remove", url_1, "", "", "", "frmtms");
	}	
    myMenu1.menuHiliteBgColor = "#617DC0";
	myMenu1.menuItemWidth = 100;
	myMenu1.menuItemHeight = 20;
	myMenu1.writeMenus();
}

function menufunctions(intAssign, intRow)
{
	window.document.frmtms.assign.value = intAssign;
	window.document.frmtms.row.value = intRow;
	window.showMenu(window.myMenu1);
}

function gopage()
{
	document.frmtms.action = "../tools/preferences.asp";
	document.frmtms.submit();
}

function viewtms()
{
	var URL;

	window.document.frmtms.M.value = window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value;
	window.document.frmtms.Y.value = window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value

	URL = "timesheetTP.asp?act=vmya";

	window.document.frmtms.action = URL;
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

var objAddSubWindow, objSetHourWindow, objPrintWindow


function addsub() { //v2.0
    window.status = "";

    strFeatures = "top=" + (screen.height / 2 - 225) + ",left=" + (screen.width / 2 - 230) + ",width=500,height=600,toolbar=no,"
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";

    if ((objAddSubWindow) && (!objAddSubWindow.closed))
        objAddSubWindow.focus();
    else {
        objAddSubWindow = window.open('tms_addsubtask.asp?m=' + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + '&y=' + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value + '&act=' + '<%=strAction%>' + '&s=' + '<%=strTPUserid%>', "MyNewWindow", strFeatures);
    }
    window.status = "Opened a new browser window.";
}

function sethour(row, col, kind) {
    window.status = "";

    strFeatures = "top=" + (screen.height / 2 - 82) + ",left=" + (screen.width / 2 - 126) + ",width=252,height=265,toolbar=no,"
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=no";

    if ((objSetHourWindow) && (!objSetHourWindow.closed))
        objSetHourWindow.close();

    objSetHourWindow = window.open('tms_writehourTP.asp?r=' + row + '&c=' + col + '&k=' + kind + '&m=' + window.document.frmtms.lbmonth.options[window.document.frmtms.lbmonth.selectedIndex].value + '&y=' + window.document.frmtms.lbyear.options[window.document.frmtms.lbyear.selectedIndex].value + '&s=' + '<%=strTPUserid%>', "MyNewWindow", strFeatures);
    objSetHourWindow.focus();

    window.status = "Opened a new browser window.";
}
function viewdetail()
{
	window.document.frmtms.action = "tms_viewdetails.asp";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
}

function window_onunload() 
{
	if((objAddSubWindow) && (!objAddSubWindow.closed))
		objAddSubWindow.close();
		
	if((objSetHourWindow) && (!objSetHourWindow.closed))
		objSetHourWindow.close();
		
	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();
}

function printpage()
{
	window.status = "";
	
	strFeatures = "top=1,left="+(screen.width/2-380)+",width=800,height=450,toolbar=no," 
	              + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";

	if((objPrintWindow) && (!objPrintWindow.closed))
		objPrintWindow.close();

objPrintWindow = window.open('tms_print_preview.asp?m=' + '<%=intMonth%>' + '&y=' + '<%=intYear%>' + '&s=' + '<%=strTPUserid %>', "MyNewWindow3", strFeatures);
	objPrintWindow.focus();
	
	window.status = "Opened a new browser window.";  
}


function back_menu() {
    window.document.frmtms.action = "listofcontractstaff.asp?b=1";
    window.document.frmtms.target = "_self";
    window.document.frmtms.submit();
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
<%	End if%>      
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
                      <td width="60%">
<!--<div class="white">&nbsp;* To remove or rename a project/sub-task, please click on that project/sub-task.</div> -->
						<div class="white">&nbsp;<%=strTitle1%></div>
					  </td> 
                      <td align="right" width="35%">
					    <select name="lbyear" size="1" class="blue-normal">
						<%For ii=2000 To Year(Date) +1%>
					      <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
						<%Next%>
						</select>
						<select name="lbmonth" size="1" class="blue-normal">
						<%For ii=1 To 12%>
					      <option <%If CInt(intMonth)=ii Then%>selected<%End If%> value="<%=ii%>"><%=SayMonth(ii)%></option>
						<%Next%>						  
						</select>
                      </td>
                      <td width="5%" align="right"> 
                        <table border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa" width="40">
                          <tr> 
                            <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
                              <div align="center" class="blue"><a href="javascript:viewtms();" class="b" onMouseOver="self.status='View timesheet';return true" onMouseOut="self.status='';return true">Go</a></div>
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
	dim rsAssignment
	intTotalRow = intRow
	If intTotalRow <= 15 Then
		intTotalRow = 15
	End If	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 
	strSql="SELECT c.AssignmentID,DateTransfer,c.StaffID,c.FgDelete, d.fgActivate " & _
			"FROM (SELECT ProjectID, MIN(DateTransfer) as DateTransfer " & _
					"FROM ATC_Projectstage GROUP BY ProjectID) a " & _
			"INNER JOIN ATC_Tasks b ON a.ProjectID=b.ProjectID " & _
			"INNER JOIN ATC_Assignments c ON b.SubTaskID=c.SubTaskID " & _
			"INNER JOIN ATC_Projects d ON a.ProjectID=d.ProjectID " & _
			"WHERE c.StaffID=" & strTPUserid & " ORDER BY c.AssignmentID"

	
	If objDatabase.dbConnect(strConnect) Then
		Call GetRecordset(strSql,rsAssignment)


		For ii = 0 To intTotalRow
			If ii <= intRow Then
				If varTimesheet(intDayCol-1,0,ii) = 0 Then
%>					  	
              <tr> 
                <td width="8"><img src="../images/cross.gif" width="8" height="14"></td>
<%
					If trim(varTimesheet(intDayCol-2,0,ii)) = "S" Then
%>                        
                <td width="118" class="blue" bgcolor="#FFF2F2"><a href="javascript:menufunctions('<%=varTimesheet(intDayCol-4,0,ii)%>','<%=ii%>');" title="<%=varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;&nbsp;&nbsp;- <%=showlabel(varTimesheet(intDayCol-3,0,ii))%></b></a></td>
<%
					ElseIf trim(varTimesheet(intDayCol-2,0,ii)) = "N" Then
%>                        
                <td width="118" class="blue" bgcolor="#FFF2F2"><a href="javascript:void(0)" title="<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=showlabel(varTimesheet(0,0,ii))%></b></a></td>
<%	
					ElseIf trim(varTimesheet(intDayCol-2,0,ii)) = "P" Then
%>
                <td width="118" class="blue" bgcolor="#FFF2F2"><a href="javascript:menufunctions('<%=varTimesheet(intDayCol-4,0,ii)%>','<%=ii%>');" title="<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>" onMouseOver="self.status=&quot;<%=varTimesheet(0,0,ii) & " _ " & varTimesheet(intDayCol-3,0,ii)%>&quot;;return true" onMouseOut="self.status='';return true" class="c"><b>&nbsp;<%=showlabel(varTimesheet(0,0,ii))%></b></a></td>
<%
					End If
					
					strShow = ""
				
					For kk = 1 To intDayNum
						dblHour = varTimesheet(kk, 0, ii) + varTimesheet(kk, 1, ii)

						strHour = "&nbsp;"
						strCurrentDate = CDate(intMonth & "/" & kk & "/" & intYear)

						strParm = CStr(ii) & "," & CStr(kk) & ",'P'" 
						strURLSetHour = "javascript:sethour("& strParm & ");"						
						intWeekDay = WeekDay(strFirstDay+(kk-1))

'if intUserID=888 then Response.Write kk & ":" & intDayCount & "-" & (kk < intDayCount)  & "<br>"

						If kk <= intDayCount Then
											
							If dblHour > 0 Then
								strHour = formatnumber(dblHour,1)
								
								if (strCurrentDate>=date()-dateLimit) AND strCurrentDate>strDateLock then
									rsAssignment.MoveFirst									 
									rsAssignment.Find "AssignmentID = " & varTimesheet(intDayCol-4,0,ii)
									if not rsAssignment.EOF then
										DateTransfer=cdate(rsAssignment("DateTransfer")) 
										blnLink= (not rsAssignment("fgDelete")) and (DateTransfer<=strCurrentDate) and (rsAssignment("fgActivate")) and (strCurrentDate>=date()-dateLimit) AND strCurrentDate>strDateLock

										if blnLink then
												strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this task&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & formatnumber(dblHour,1) & "</a>"
										end if
									end if
								end if
													
							else
							    'strHour = "&nbsp;"	
							    'strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this task&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>--</a>"

                                If varTimesheet(intDayCol-2, 0, ii) = "N" Then
	                                strHour = "&nbsp;"
                                Else
	                                strHour = "&nbsp;"
'Response.Write strCurrentDate & "--" & (date()-dateLimit) & "--" & strDateLock & "<br>"			                                
'Response.Write kk & "--" &  varTimesheet(intDayCol-2, 0, ii) & "--" & ii &  "<br>"
	                                
	                                if (strCurrentDate>=date()-dateLimit) AND strCurrentDate>strDateLock then
		                                rsAssignment.MoveFirst								 
		                                rsAssignment.Find "AssignmentID = " & varTimesheet(intDayCol-4,0,ii)
		                                if not rsAssignment.EOF then
			                                DateTransfer=cdate(rsAssignment("DateTransfer")) 
			                                blnLink= (not rsAssignment("fgDelete")) and (DateTransfer<=strCurrentDate) and (rsAssignment("fgActivate")) 
'Response.Write 	blnLink & "--" & DateTransfer & "--" & strCurrentDate & "<br>"			                                
			                                if blnLink then
				                              
					                                strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this task&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & "--" & "</a>"
				                              
			                                end if
		                                end if
	                                end if
                                			
                                End If	

							    
							End If	
							
						
						End If			
						
						'intWeekDay = WeekDay(strFirstDay+(kk-1))
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
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#FFFFFF" <%End If%> align="center" class="blue-normal" width="19"><%=strHour%></td>
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
<%				End If
			Else%>                      
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
	End If
	objDatabase.dbDisConnect()																' Disconnect to SQL database	
	Set objDatabase = Nothing	
%>           

<!--**************************** End Of Project And SubTask *******************************-->
  
<!--**************************** Add Sub-Task *********************************************-->

<%If strError = "" Then%>
			  <tr>

                <td width="8" bgcolor="#FFC6C6" class="white"><img src="../../images/cross.gif" width="8" height="14"></td>
                <td width="118" bgcolor="#FFF2F2" class="blue-normal"><a href="javascript:addsub();" >&nbsp;Add SubTask</a></td>


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

<!--**************************** End Of Add Sub-Task *********************************************-->

<%End if%>

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
				
			strParm = CStr(ii) & "," & CStr(kk) & ",'E'"
			
			strURLSetHour = "javascript:sethour("& strParm & ");"						
			intWeekDay = WeekDay(strFirstDay+(kk-1))
			
			strHour = IIf(dblHour > 0,formatnumber(dblHour,1),"&nbsp;")
			
			if strCurrentDate="" then strCurrentDate = CDate(intMonth & "/" & kk & "/" & intYear)
			
			
				If (trim(varEvent(0,0,ii)) = "General/Admin" OR  trim(varEvent(0,0,ii)) = "Personal Time" OR trim(varEvent(0,0,ii)) = "Other Leave") AND strCurrentDate>strDateLock then
					strHour = "<a class='c' href=" & strURLSetHour & " onMouseOver='self.status=&quot;Write hour for this event&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>" & IIf(dblHour > 0,formatnumber(dblHour,1),"--") & "</a>"
				end if

			If strError = "No data for your request." And trim(varEvent(0,0,ii)) <> "Annual Holiday" Then
				strHour = "&nbsp;"
			End If

			
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
                <td <%If strColorOpt <> "" Then%> bgcolor="<%=strColorOpt%>" <%Else%> bgcolor="#E7EBF5" <%End If%> align="center" class="blue-normal" width="19"><%=strHour%></td>
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
  
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(1))
%>
<input type="hidden" name="M" value="<%=intMonth%>">
<input type="hidden" name="Y" value="<%=intYear%>">
<input type="hidden" name="txtUserid" value="<%=strTPUserid%>">
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
<input type="hidden" name="assign" value="<%=Request.Form("assign")%>">
<input type="hidden" name="row" value="">
</form>
<SCRIPT LANGUAGE="JavaScript">
<!--
//For IE
if (document.all) 
{
    onLoad();
}

//-->
</SCRIPT>
<%If Request.QueryString("act") = "U" Then%>
<SCRIPT language="javascript">
	addsub();
</SCRIPT>
<%End If%>

</body>
</html>
