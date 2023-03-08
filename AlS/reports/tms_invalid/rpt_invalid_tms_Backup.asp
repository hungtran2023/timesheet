<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<%
	Dim intUserID, intMonth, intYear, intDayNum, intWeekday, intRow, intCount,strAct
	Dim varFullName, varFrom, varTo, varPre, getRes, varUser, varInvalidTMS,rsEmailCount
	Dim strUserName, strTitle, strFunction, strMenu, strURL, strType, strTitle2, strFrom, strTo, strFirstDay, strCurDate, strDateShow

function GetNumberOfEmailByStaff(byval staffID)
	dim intNum
	intNum=""
	
	if not rsEmailCount.EOF then
		rsEmailCount.Filter="RecipientID=" & staffID
		if not rsEmailCount.EOF then intNum=rsEmailCount("numOfReminder")
		rsEmailCount.Filter=""
	end if
	GetNumberOfEmailByStaff=intNum
	
End function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID	= session("USERID")
	
'--------------------------------------------------
' Initialize variables	
'--------------------------------------------------
	
	intCount = -1
	Redim varInvalidTMS(5,-1)
	
	strAct=Request.QueryString("act")
	strConnect = Application("g_strConnect")	' Connection string 
	
	Set objDatabase = New clsDatabase 
	
	if strAct="" Or strAct="f" then
	
		if strAct="" then
			strTo=date()-1
			if Weekday(cdate(strTo),1)=7 then strTo=date()-2
			if Weekday(cdate(strTo),1)=1 then strTo=date()-3
			
			strFrom=strTo
			intMonth = month(date())
			intYear	 = year(date())
			strType="D"
	
		'For filter
		ElseIf strAct = "f" Then
	
			strType = Request.Form("rdotype")
			If strType = "D" Then
				strFrom		= Request.Form("txtFrom")
				strTo		= Request.Form("txtTo")
				
				strFrom=cdate(ConvertTommddyyyy(strFrom))
				strTo	=cdate(ConvertTommddyyyy(strTo))
					
			Else
	
				intMonth = Request.Form("lbmonth")
				intYear	 = Request.Form("lbyear")
				strTitle2 = SayMonth(intMonth) & "/" & intYear
					
				strFrom=cdate(intMonth & "/01/" & intYear)
				strTo	=DateAdd("m",1,strFrom) -1	
			End if
		end if
		
		strTitle2	= "From " & ddmmyyyy(strFrom) & " To " & ddmmyyyy(strTo)
		intDayNum	= strTo - strFrom
		strFirstDay = strFrom
		
		'--------------------------------------------------
		' Initialize holiday array
		'--------------------------------------------------
		
		If isEmpty(session("varHoliday")) = False Then	session("varHoliday") = Empty
			
		If objDatabase.dbConnect(strConnect) Then
			strSQL = "exec GetListHolidays null, null, '" & strFrom & "', '" & strTo & "', 1"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					varHoliday = objDatabase.rsElement.GetRows
					session("varHoliday") = varHoliday
					objDatabase.closeRec
				End If
			Else
				strError = objDatabase.strMessage
			End If
		Else
			strError = objDatabase.strMessage		
		End If

		'--------------------------------------------------
		' End of initializing holiday array
		'--------------------------------------------------

		Set varHoliday = Nothing

	else
		strType = Request.Form("rdotype")
		strFrom		= Request.Form("txtFrom")
		strTo		= Request.Form("txtTo")
		intMonth = Request.Form("lbmonth")
		intYear	 = Request.Form("lbyear")
		
		varUser  = session("varInvalidTMS")
		
		If isArray(varUser) Then intRow = Ubound(varUser,2)
		
		If Request.QueryString("act") = "vrae" Then
			strError = "Please choose the checkbox before click Remind."
		Else
		    'For reminder
			intUser = Request.Form("chkremind").Count
			Redim varInvalidUser(1, intUser)
			If intUser > 0 Then
				For ii = 1 To intUser
					strTemp=split(Request.Form("chkremind")(ii),"#")
			
					varInvalidUser(0,ii) = strTemp(0)
					varInvalidUser(1,ii) = strTemp(1)
				Next
				session("varInvalidUser") = varInvalidUser			
				
			End If	
		End If	

	End If
	
	strSql="SELECT RecipientID, COUNT(RecipientID) as numOfReminder FROM ATC_StaffBeReminded WHERE EmailRemindID IN (SELECT EmailRemindID FROM ATC_EmailReminded WHERE YEAR(SentDate)=YEAR(Getdate())) GROUP BY RecipientID"
	call GetRecordset(strSQL,rsEmailCount)


'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strFunction = "<a class='c' href='javascript:gopage();' onMouseOver='self.status=&quot;Preferences&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Preferences</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:printpage()' onMouseOver='self.status=&quot;Print missing timesheet page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Print</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	Set objEmployee = Nothing
	
'--------------------------------------------------
' Make list of menu
'--------------------------------------------------

	If isEmpty(session("Menu")) Then 
		getRes = getarrMenu(intUserID)
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End If	

'--------------------------------------------------
' Get current URL
'--------------------------------------------------
	
	If Request.ServerVariables("QUERY_STRING") <> "" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
'--------------------------------------------------
' Get current menu that user is choosing
'--------------------------------------------------
	
	strChoseMenu = Request.QueryString("choose_menu")
	If strChoseMenu = "" Then strChoseMenu = "B"

	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))

	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1) <> "" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if

%>
<html>
<head>
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">

<script language="javascript" src="../../library/library.js"></script>
<script language="javascript">
<!--
var ns, ie, objNewWindow;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	var url;
	url = "../../logout.asp";
	if (ns)
		document.location = url;
	else
	{
		window.document.frmreport.action = url;
		window.document.frmreport.submit();
	}	
}

function setchecked(val) 
{
	with (document.frmreport) 
	{
		len = elements.length;
		for(var ii=0; ii<len; ii++) 
     		if (elements[ii].name == "chkremind") 
				elements[ii].checked = val
	}
}
function gopage()
{
	document.frmreport.action = "../../tools/preferences.asp";
	document.frmreport.submit();
}

function printpage() 
{ //v2.0
var row = "<%=intRow%>";
//	if ("<%=intRow%>" != "" && "<%=intRow%>" >= 0)
	//if (row != "" && row >= 0)
	//{
		window.status = "";
 
		strFeatures = "top=1,left="+(screen.width/2-350)+",width=630,height=680,toolbar=no," 
		          + "menubar=yes,location=no,directories=no,resizable=no,scrollbars=yes";
              
		if((objNewWindow) && (!objNewWindow.closed))
			objNewWindow.focus();	
		else 
		{
			objNewWindow = window.open('rpt_print_preview.asp?title=' + '<%=strtitle2%>', "MyNewWindow", strFeatures);
		}
		window.status = "Opened a new browser window.";  
	//}	
}

function checkdata()
{
	with (document.frmreport) 
	{
		len = elements.length;
		for(var ii=0; ii<len; ii++) 
		{
     		if (elements[ii].name == "chkremind") 
			{
				if (elements[ii].checked == true)
				{
					return true;
					break;
				}	
			}
		}		
	}
}

function sendmail()
{
	if (checkdata())
	{
		document.frmreport.action = "rpt_invalid_tms.asp?act=vra"
		document.frmreport.submit();
	}
	else
	{
		document.frmreport.action = "rpt_invalid_tms.asp?act=vrae"
		document.frmreport.submit();
	}	
}

function checkdataFilter()
{
	
	if (document.frmreport.rdotype[0].checked)
	{
		if (isnull(document.frmreport.txtFrom.value)==true)
		{
			alert("Please enter startdate before click here.")
			document.frmreport.txtFrom.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtFrom.value)==false)
			{			
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtFrom.focus();
				return false;
			}
		}
		
		if (isnull(document.frmreport.txtTo.value)==true)
		{
			alert("Please enter enddate before click here.")
			document.frmreport.txtTo.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtTo.value)==false)
			{
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtTo.focus();
				return false;
			}
		}
		
		if (comparedate(document.frmreport.txtFrom.value,document.frmreport.txtTo.value)==false)
		{
			alert("The startdate must be less than the finishdate.")
			document.frmreport.txtFrom.focus();
			return false;
		}
	}	
	return true;
}

function viewtms()
{
	if (checkdataFilter() == true)
	{
		document.frmreport.action = "rpt_invalid_tms.asp?act=f"
		document.frmreport.submit();
	}	
}

function document_onkeypress() 
{
var keycode = event.keyCode;
	if (keycode == 13) 
	{
		event.keyCode = 0;
		viewtms();
	}
}

//-->
</script>

<script LANGUAGE="javascript" FOR="document" EVENT="onkeypress">
<!--
 document_onkeypress()
//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(0))
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
	Response.Write(arrTmp(0))
	
%>		

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
		<tr bgcolor="#FFFFFF" >
  	      <td colspan="2" >
  	      
  	      <table width="80%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="32%">&nbsp;</td>
                <td width="8%">&nbsp;</td>
                <td width="20%">&nbsp;</td>
                <td width="8%">&nbsp;</td>
                <td width="15%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td valign="top" class="blue" align="right">
                        <input type="radio" name="rdotype" value="D" <%if strType="D" then%>checked<%End If%> language="javascript" onClick="document.frmreport.txtFrom.focus()"></td>
                <td class="blue">From</td>
                <td valign="top" class="blue-normal"><input type="text" name="txtFrom" id="txtFrom" size="10" class="blue-normal" language="javascript" onClick="document.frmreport.rdotype[0].checked=true" <%If strType="D" Then%>value="<%=ddmmyyyy(strFrom)%>"<%End If%>></td>
                <td class="blue">To</td>
                <td valign="top" class="blue-normal"><input type="text" name="txtTo" size="10" class="blue-normal" language="javascript" onClick="document.frmreport.rdotype[0].checked=true" <%If strType="D" Then%>value="<%=ddmmyyyy(strTo)%>"<%End If%>></td>
                <td >&nbsp;</td>
              </tr>
              <tr bgcolor="#FFFFFF">                
                <td align="right" class="blue"><input type="radio" name="rdotype" value="M" <%if strType="M" then%>checked<%End If%> language="javascript" onClick="document.frmreport.lbmonth.focus()"> </td>
                <td class="blue">Month</td>
                <td class="blue-normal"> 
					<select name="lbmonth" size="1" class="blue-normal" language="javascript" onFocus="document.frmreport.rdotype[1].checked=true">
						<%For i=1 to 12 %>
						<option  value="<%=i%>" <%If CInt(intMonth)=i Then%>Selected<%end if%>><%=SayMonth(i)%></option>
						<%Next%>
                  </select>
                </td>
                <td class="blue">Year</td>
                <td class="blue-normal">
                <select name="lbyear" size="1" class="blue-normal" language="javascript" onFocus="document.frmreport.rdotype[1].checked=true">
					<%For ii=Year(Date)-1 To Year(Date)%>
				      <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
					<%Next%>
				</select>
                </td>
                <td align="left"><table width="60" border="0" cellspacing="2" cellpadding="0" height="20" name="aa">
					<tr> 
					  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" width="59" height="20" > 
					    <div align="center" class="blue"> 
					      <a href="javascript:viewtms();" class="b" onMouseOver="self.status='Clich here to view missing timesheet';return true" onMouseOut="self.status='';return true">Submit</a> 
					    </div>
					  </td>
					</tr>
					</table></td>
              </tr>
            </table>  	      
  	      </td>
		</tr>
		<%If strError<>"" Then%> 
		<tr bgcolor="#FFFFFF" height="25">
  	      <td colspan="2" class="red">&nbsp;<b><%=strError%></b></td>
		</tr>
		<tr bgcolor="#617DC0" height="1">
  	      <td colspan="2"></td>
		</tr>
		<%End if%>

        <tr align="center"> 
          <td class="title" height="50" align="center" colspan="2"> Missing Timesheet<br>
            <div class="blue-normal"><%=strTitle2%></div>
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
                  <table width="100%" border="0" cellspacing="1" cellpadding="3">
                    <tr> 
                      <td class="blue" align="center" width="10%" bgcolor="#E7EBF5">Date</td>
                      <td class="blue" align="center" width="25%" bgcolor="#E7EBF5">Full Name </td>
                      
                      <td class="blue" align="center" width="30%" bgcolor="#E7EBF5">Report To </td>
                      <td class="blue" align="center" width="10%" bgcolor="#E7EBF5"> No. Of Reminder </td>
                      <td class="blue" align="center" width="10%" bgcolor="#E7EBF5">&nbsp;</td>
                    </tr>
<%                    
	If strAct <> "" and strAct<>"f" Then
		If intRow >= 0 Then
			For ii = 0 To intRow
				If varUser(1,ii) = "Date" Then
%>					
                    <tr> 
                      <td valign="top" colspan="6" class="blue" bgcolor="#FFF2F2"><%=varUser(0,ii)%></td>
                    </tr>
<%
				Else
%> 
                    <tr> 
                      <td valign="middle"  class="blue" align="right" bgcolor="#FFFFFF"><%=varUser(0,ii)%></td>
                      <td valign="middle" class="blue-normal" bgcolor="#FFFFFF">&nbsp;<%=showlabel(varUser(2,ii))%></td>
                      
                      <td valign="middle" class="blue-normal" bgcolor="#FFFFFF">&nbsp;<%=showlabel(varUser(4,ii))%></td>
                      <td valign="middle" class="blue-normal" bgcolor="#FFFFFF" align="right">&nbsp;<%=GetNumberOfEmailByStaff(varUser(5,ii))%></td>
                      <td valign="middle" class="blue-normal" align="center" bgcolor="#FFFFFF">&nbsp; 
                        <input type="checkbox" name="chkremind" value="<%=varUser(3,ii) & "#" & varUser(5,ii)%>">
                      </td>
                    </tr>
<%
				End If
			Next
		End If
		Set varUser = Nothing
	Else
		If intDayNum >= 0 Then
			intRow		= -1	
			For kk = 0 To intDayNum		
				strCurDate = strFirstDay + kk					
				If Day(strCurDate) < 10 Then
					strDateShow = "0" & Day(strCurDate) & "/" & SayMonth(Month(strCurDate)) & "/" & Year(strCurDate)
				Else	
					strDateShow = Day(strCurDate) & "/" & SayMonth(Month(strCurDate)) & "/" & Year(strCurDate)
				End If
						
'--------------------------------------------------
' Check this date if it is a working day or not
'--------------------------------------------------				
				intWeekday = WeekDay(strCurDate)
				If objDatabase.dbConnect(strConnect) Then
					strSQL = "SELECT fgDayOff FROM ATC_WeekDay WHERE WeekDayID=" & intWeekday
					If (objDatabase.runQuery(strSQL)) Then
						If Not objDatabase.noRecord Then
							If objDatabase.getColumn_by_name("fgDayOff") Then
								fgDayOff = 1
							Else
								fgDayOff = 0
							End If					
							objDatabase.closeRec
						End If
					Else
						strError = objDatabase.strMessage
					End If
				Else
					strError = objDatabase.strMessage
				End If		

				If isHoliday(Day(strCurDate)) = -1 Then
			
					If CInt(fgDayOff) = 0 Then
			
						If objDatabase.dbConnect(strConnect) Then			
							strSQL = "exec InvalidTMS '" & strCurDate & "'"
					
							If (objDatabase.runQuery(strSQL)) Then							
								If Not objDatabase.noRecord Then
									varUser = objDatabase.rsElement.GetRows
									intRow  = Ubound(varUser,2)
									objDatabase.closeRec
								else
									intRow=-1
								End If
							Else
								strError = objDatabase.strMessage 
								
							End If
						Else
'							Response.Write objDatabase.strMessage		
							strError = objDatabase.strMessage
						End If

						strDateShow= strDateShow & " - " & weekdayname(Weekday(strCurDate))
						If intRow <> "" And intRow >= 0 Then
							intCount = intCount + 1
							Redim Preserve varInvalidTMS(5,intCount)
							varInvalidTMS(0,intCount) = strDateShow
							varInvalidTMS(1,intCount) = "Date"
%>					
                    <tr> 
                      <td valign="middle" colspan="6" class="blue" bgcolor="#FFF2F2"><%=strDateShow%></td>
                    </tr>
<%
							For ii = 0 To intRow
								intCount = intCount + 1
								Redim Preserve varInvalidTMS(5,intCount)
								
								varInvalidTMS(0,intCount) = ii+1			'Number count
								varInvalidTMS(1,intCount) = varUser(1,ii)	'PC Name
								varInvalidTMS(2,intCount) = varUser(2,ii)	'Fullname
								varInvalidTMS(3,intCount) = varUser(0,ii)	'Email adress
								varInvalidTMS(4,intCount) = varUser(3,ii)	'Leader
								varInvalidTMS(5,intCount) = varUser(4,ii)	'LeaderID
%>
                    <tr> 
                      <td valign="middle" class="blue" align="right" bgcolor="#FFFFFF"><%=ii+1%></td>
                      <td valign="middle" class="blue-normal" bgcolor="#FFFFFF"><%=showlabel(varUser(2,ii))%></td>
                      
                      <td valign="middle" class="blue-normal" bgcolor="#FFFFFF"><%=showlabel(varUser(3,ii))%></td>
                      <td valign="middle" class="blue-normal" bgcolor="#FFFFFF" align="right"><%=GetNumberOfEmailByStaff(varUser(4,ii))%></td>
                      <td valign="middle" class="blue-normal" align="center" bgcolor="#FFFFFF"> 
                        <input type="checkbox" name="chkremind" value="<%=varUser(0,ii) & "#" & varUser(4,ii)%>">
                      </td>
                    </tr>
<%
							Next
						End If	
					End If
				End If
			Next
			If IsEmpty(Session("varInvalidTMS")) = False Then
				Session("varInvalidTMS") = Empty
			End If
				
			Session("varInvalidTMS") = varInvalidTMS
		End If
	End If	
%>                  
                  </table>
<%	If intRow <> "" And intRow >= 0 Then%>                  
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#FFFFFF"> 
                      <td width="69%" class="blue-normal" height="20">&nbsp;&nbsp;* Choose the checkbox, then click Remind to send a reminding email.</td>
                      <td width="31%" class="blue" align="right">
                        <a href="javascript:setchecked(1);" onMouseOver="self.status='';return true">Select All</a>&nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:setchecked(0);" onMouseOver="self.status='';return true">Clear All</a> &nbsp;&nbsp;&nbsp;<a href="javascript:sendmail();" onMouseOver="self.status='';return true">Remind</a>&nbsp;&nbsp;
                      </td>
                    </tr>
                  </table>
<%
	End If
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
<%
	Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))
%>
<input type="hidden" name="title" value="<%=strTitle2%>">
</form>

<%
	If Request.QueryString("act") = "vra" and Request.QueryString("choose_menu")="" Then
%>
<script language="javascript">
<!--
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-125)+",left="+(screen.width/2-230)+",width=800,height=500,toolbar=no," 
			      + "menubar=no,location=no,directories=no,resizable=yes,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('rpt_send_mail.asp', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  

//-->
</script>
<%	End If%>
</body>
</html>
