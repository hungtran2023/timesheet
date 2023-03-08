<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/libraryForTPTimesheet.asp"-->

<%
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1	
	
	Dim intCol, intRow, intUserID, intDayNum, intMonth, intYear, Col, Row, intAssignmentID, ii
	Dim dblNormalHour, dblOverHour,dblOldValue
	Dim strType, strProject, strDate, strNote, strTableTMS, strSQL, objDatabase
	
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	
	Col			= Request.QueryString("c")
	Row			= Request.QueryString("r")	
	strType		= Request.QueryString("k")
	intMonth	= Request.Querystring("m")
	intYear		= Request.Querystring("y")
	intDayNum	= GetDay(intMonth,intYear)				' Numbers of days in a month
	intDayCol	= intDayNum + 6

	strDay		= CDate(intMonth & "/" & Col & "/" & intYear)
	intWeekday	= Weekday(strDay)
	intRow		= -1
	eRow		= -1

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
%>	
<script type="text/javascript">
<!--
	opener.document.location = "../message.htm";
	window.close();
//-->
</script>
<%
	End If					

Function CheckStaffDevelop(col)
	dim blnAlready

	blnAlready= (CDbl(varEvent(Col, 0, 0))>0)

	CheckStaffDevelop=blnAlready	
end function

'--------------------------------------------------
' Check this date if it was holiday or not
'--------------------------------------------------
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	
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
	
	If isHoliday(col) >= 0 Then
		fgDayOff = 1
	End If
		
	Set objDatabase = Nothing

	intUserID	= session("USERID")
	intStaffID = Request.QueryString("s")

	
	intDayNum	= GetDay(intMonth,intYear)							' Numbers of days in a month
	strDate		= Col & "/" & SayMonth(intMonth) & "/" & intYear	 

	varTimesheet= session("varTimesheet")							' Array stores timesheet data
	varEvent	= session("varEvent")								' Array stores event data
	
	If isarray(varTimesheet) Then
		intRow	= Ubound(varTimesheet,3)
	End If
	
	If isarray(varEvent) Then
		eRow	= Ubound(varEvent,3)
	End If

	If strType  = "P" Then
	
		If CDbl(varTimesheet(Col, 0, Row)) = 0 Then
			dblNormalHour = ""
		Else
			dblNormalHour = FormatNumber(varTimesheet(Col, 0, Row),1)
		End If

		If CDbl(varTimesheet(Col, 1, Row)) = 0 Then
			dblOverHour = ""
		Else
			dblOverHour = FormatNumber(varTimesheet(Col, 1, Row),1)	
		End If
		
		If CDbl(varTimesheet(Col, 2, Row)) = 0 Then
			dblOTNormal = ""
		Else
			dblOTNormal = formatnumber(varTimesheet(Col, 2, Row),1)	
		End If
		
		If CDbl(varTimesheet(Col, 3, Row)) = 0 Then
			dblOTNight = ""
		Else
			dblOTNight = formatnumber(varTimesheet(Col, 3, Row),1)	
		End If
		strNote = showvalue(trim(varTimesheet(Col, 4, Row)))
		
		If trim(varTimesheet(intDayCol-2,0,Row)) = "S" Then
			strProject = trim(varTimesheet(intDayCol-3, 0, Row))
		Else
			strProject = trim(varTimesheet(0, 0, Row))
		End If	
		
		fgStaffDevelop=CheckStaffDevelop(col)

	ElseIf strType = "E" Then
	
		If CDbl(varEvent(Col, 0, Row)) = 0 Then
			dblNormalHour = ""
			dblOldValue=0
		Else
			dblNormalHour = formatnumber(varEvent(Col, 0, Row),1)
			dblOldValue=dblNormalHour
		End If

		If CDbl(varEvent(Col, 1, Row)) = 0 Then
			dblOverHour = ""
		Else
			dblOverHour = formatnumber(varEvent(Col, 1, Row),1)	
		End If
		
		If CDbl(varEvent(Col, 2, Row)) = 0 Then
			dblOTNormal = ""
		Else
			dblOTNormal = formatnumber(varEvent(Col, 2, Row),1)	
		End If
		
		If CDbl(varEvent(Col, 3, Row)) = 0 Then
			dblOTNight = ""
		Else
			dblOTNight = formatnumber(varEvent(Col, 3, Row),1)	
		End If
		
		strNote = showvalue(trim(varEvent(Col, 4, Row)))
		strProject = trim(varEvent(0, 0, Row))

	End If	
	
	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	intDepartID = varFullName(2)
	Set objEmployee = Nothing

'--------------------------------------------------
' Check data in the form for inserting/updating/deleting of timesheet table
'--------------------------------------------------
		
		
	If Request.QueryString("act") = "U" Then
		Dim strError
		
		strDay = intMonth & "/" & Col & "/" & intYear


		If strError1 = "" Then
			
				 strError = tmsWriteHourforTP(intStaffID, Col, Row, strType, intMonth, intYear, "txtoffhour","txtOTNormal","txtOTNight", "txtnote")
			 
				 if strError="" then
					
						if Request.Form("chkPersonalDev")<>"" then					
							strError = tmsWriteHourForStaffDevelopment(intStaffID, Col, 0, "E", intMonth, intYear, 0.5)
							
							if strError<>"" then							
%>

<script type="text/javascript">
<!--
	alert("<%=strError%>" + "\n" + "<%=strError1%>");
-->
</script>

<%							
							
							
								strError=""
							end if					
						end if
				end if
				
			
		
			If strError = "" Then
				varCookies = "timesheet_" & intStaffID
				ndate = DateAdd("d",1,now)
				Response.Cookies(varCookies) = date
				Response.Cookies(varCookies).expires = ndate
%>

<script type="text/javascript">
<!--
    //alert("done");
    opener.document.frmtms.txtUserid.value = "<%=intStaffID%>";
	opener.document.frmtms.action = "timesheetTP.asp?act=vpa"
	opener.document.frmtms.submit()
	window.close();
//-->
</script>

<%		
			Else
				If InStr(1,strError,"@@") <> 0 Then
					varError = split(strError,"@@")
					strError1 = varError(0)
					strError2 = varError(1)
				ElseIf InStr(1,strError,"hours per day") Then
					strError1 = strError
				Else
					Response.Write strError
				End If
			End If						
		End If		
	End If	
%>
<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Timesheet System - Write timesheet</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">

<script type="text/javascript" src="../../library/library.js"></script>
<script type="text/javascript">
<!--

    ns = (document.layers) ? true : false;
    ie = (document.all) ? true : false;

var row = "<%=Row%>";
var kind = "<%=strType%>";
var col = "<%=Col%>";

function checkdata()
{
	var intWd=<%=intWeekDay%>;
	if ("<%=fgDayOff%>" == 0)
	{
		if (isnull(document.frmtms.txtoffhour.value)==false)
		{
			if (isNaN(document.frmtms.txtoffhour.value) ==  true) 
			{
				alert("Invalid value office hour field!");
				document.frmtms.txtoffhour.focus(); 
				return false;
			}	
			
			var offhour = document.frmtms.txtoffhour.value
			var vartemp	= offhour.split(".");
			if (vartemp.length > 1)
			{
				if (parseInt(vartemp[1].length) > 1)
				{
					alert("Invalid value office hour field. Please use this format: x.x");
					document.frmtms.txtoffhour.focus(); 
					return false;
				}
			}	
		}
	}	

	if ((kind == "P") || ((kind == "E") && (row == 0 || row == 1))) 
	{
		if (isnull(document.frmtms.txtOTNight.value)==false)				
		{
			if (isNaN(document.frmtms.txtOTNight.value) ==  true) 
			{
				alert("Invalid value overtime hour for normal rate."); 
				document.frmtms.txtOTNight.focus();
				return false;
			}			
		}
		
		if (isnull(document.frmtms.txtOTNormal.value)==false)				
		{
			if (isNaN(document.frmtms.txtOTNormal.value) ==  true) 
			{
				alert("Invalid value overtime hour for night rate."); 
				document.frmtms.txtOTNormal.focus();
				return false;
			}			
		}
		var OTNormalhour = document.frmtms.txtOTNormal.value	
		if (OTNormalhour>0)
		{
			if ((intWd != 1)&&(intWd != 7))
			{
				if (OTNormalhour>8.5)
				{
					alert("The OT for normal rate in working day must be less than 3.5."); 
					document.frmtms.txtOTNormal.focus();
					return false;
				}
			}
			else
			{
				if (OTNormalhour>13.5)
				{
					alert("The OT for normal rate in day off must be less than 12.5."); 
					document.frmtms.txtOTNormal.focus();
					return false;
				}
			}
		}
	}
	return true;
}
function writehour()
{
	if (checkdata() == true)
	{
		if (ns)
			document.location = "tms_writehourTP.asp?act=U&r=" + row + "&c=" + col + "&k=" + kind + "&m=" + "<%=intMonth%>" + "&y=" + "<%=intYear%>" + "&s=" + "<%=intStaffID%>";
		else
		{
		    window.document.frmtms.action = "tms_writehourTP.asp?act=U&r=" + row + "&c=" + col + "&k=" + kind + "&m=" + "<%=intMonth%>" + "&y=" + "<%=intYear%>" + "&s=" + "<%=intStaffID%>"; 
			window.document.frmtms.submit();
		}
	}
}


//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="frmtms" method="post">
<table width="252" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#C0CAE6" height="265">
  <tr> 
    <td valign="top"> 
      <table width="250" border="0" cellspacing="0" cellpadding="0" align="center">
<%
		If strError1 <> "" Or strError <> "" Then		
%>      
          <tr bgcolor="#C0CAE6" align="center" valign="middle"> 
            <td colspan="2" height="30" class="blue"><%=strError1%></td>
          </tr>
<%		End If%>

<%
		If strError2 <> "" Then		
%>      
          <tr bgcolor="#C0CAE6" align="center" valign="middle"> 
            <td colspan="2" height="30" class="blue"><%=strError2%></td>
          </tr>
<%		End If%>
          <tr bgcolor="#C0CAE6" align="center"> 
            <td colspan="2" height="40" class="title">Time Write</td>
          </tr>
          <tr bgcolor="#C0CAE6"> 
            <td width="99" class="blue-normal" height="26"> 
              <div align="right" class="c"> Project&nbsp; </div>
            </td>
            <td width="151" class="blue">&nbsp;<b><%=showlabel(strProject)%></b></td>
          </tr>
          <tr bgcolor="#C0CAE6"> 
            <td width="99" class="blue-normal" height="26"> 
              <div align="right"> Date&nbsp; </div>
            </td>
            <td width="151" class="blue">&nbsp;<b><%=strDate%></b></td>
          </tr>
<%		If CInt(fgDayOff) = 0 Then%> 
         
          <tr bgcolor="#C0CAE6"> 
            <td width="99" class="blue-normal" align="right" height="26">Normal Hours&nbsp; </td>
            <td width="151" class="text-blue01"> 
              <input type="text" name="txtoffhour" class="blue-normal" size="5" style="width:50" tabindex="1" value="<%=dblNormalHour%>">
            </td>
          </tr>

<%		End If%>          
<%		If (strType = "P") Or (strType = "E" And (Row = 0 Or Row = 1)) Then%>                
		  <tr bgcolor="#C0CAE6"> 
             <td width="99" height="26" align="right" valign="bottom" class="blue-normal">Overtime&nbsp;</td>
            <td width="151" class="text-blue01">
              <table width="90%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td class="blue-normal">Normal Rate</td>
                  <td class="blue-normal">Night Rate</span></td>
                </tr>
                <tr>
                  <td><input type="text" name="txtOTNormal" class="blue-normal" size="5" style="width:60" tabindex="2" value="<%=dblOTNormal%>"></td>
                  <td><input type="text" name="txtOTNight" class="blue-normal" size="5" style="width:60" tabindex="2" value="<%=dblOTNight%>"></td>
                </tr>
              </table></td></tr>
<%		End If%>          
          <tr bgcolor="#C0CAE6"> 
            <td width="99" class="blue-normal" align="right" height="26">Note&nbsp; </td>
            <td width="151" class="text-blue01"> 
              <input type="text" name="txtnote" class="blue-normal" size="10" style="width:130" tabindex="3" value="<%=strNote%>">
            </td>
          </tr>
<%		If CInt(fgDayOff) = 0 and not fgStaffDevelop and strType = "P" Then%>           
          <tr bgcolor="#C0CAE6"> 
            <td class="blue-normal" align="right" height="26"> </td>
            <td class="blue-normal"><input type="checkbox" name="chkPersonalDev" class="blue-normal" value="1">Personal Time</td>
          </tr>
<%		End If%>           
          <tr bgcolor="#C0CAE6">
            <td height="40" colspan="2"> 
              <table width="120" border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa">
                <tr> 
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" class="blue" height="20" align="center"> 
                    <a href="javascript:writehour();" class="b">Submit</a>
                  </td>
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" class="blue" height="20" align="center">
                    <a href="#" class="b" onClick="window.close()">Close</a>
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
