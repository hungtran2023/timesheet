<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
	
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim intUserID, intMonth, intYear, intLeaveDue, intCurLeave, intBalance
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
%>	
<script LANGUAGE="javascript">
<!--
	opener.document.location = "../../message.htm";
	window.close();
//-->
</script>
<%
	End If	

	intUserID	= session("USERID")
	intStaffID  = Request.Form("txthidden")

	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	intDepartID = varFullName(2)

	Set objEmployee = Nothing

	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then
	
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "sp_StaffLeavedueforthePast"
		Set myParam = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
		myCmd.Parameters.Append myParam	
		Set myParam = myCmd.CreateParameter("balance",adVarChar,adParamOutput,10)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("expiredDate",adVarChar,adParamOutput,11)
		myCmd.Parameters.Append myParam		
		myCmd("StaffID") = intStaffID
		myCmd.Execute
					
		decBalancePast=cdbl(myCmd("balance"))
		if not isnull(myCmd("expiredDate")) then dateExipred=cdate(myCmd("expiredDate"))
					
		Set myCmd_ = Server.CreateObject("ADODB.Command")
		Set myCmd_.ActiveConnection = objDatabase.cnDatabase
		myCmd_.CommandType = adCmdStoredProc
		myCmd_.CommandText = "sp_StaffLeavedueforthisyear"
		Set myParam = myCmd_.CreateParameter("StaffID",adInteger,adParamInput)
		myCmd_.Parameters.Append myParam
		Set myParam = myCmd_.CreateParameter("expiredDate",adDate,adParamInput)
		myCmd_.Parameters.Append myParam
		Set myParam = myCmd_.CreateParameter("leaveduethisYear",adVarChar,adParamOutput,10)
		myCmd_.Parameters.Append myParam
		Set myParam = myCmd_.CreateParameter("appBeforeExpire",adVarChar,adParamOutput,10)
		myCmd_.Parameters.Append myParam
		Set myParam = myCmd_.CreateParameter("appAfterExpire",adVarChar,adParamOutput,10)
		myCmd_.Parameters.Append myParam
		Set myParam = myCmd_.CreateParameter("WorkingHours",adVarChar,adParamOutput,10)
		myCmd_.Parameters.Append myParam
							
		myCmd_("expiredDate") = myCmd("expiredDate")
		myCmd_("StaffID") = intStaffID
		myCmd_.Execute
				
		intAppCur=cdbl(myCmd_("appBeforeExpire"))
		intLeaveDueCur=cdbl(myCmd_("leaveduethisYear"))
		intWorkHours=myCmd_("WorkingHours")
					
		intBalance=cdbl(intLeaveDueCur) + cdbl(decBalancePast) - cdbl(intAppCur)
		'Marcel has expiredDate be null value
					
		if not isnull(myCmd("expiredDate")) then
			if date>=dateExipred then
				if decBalancePast-intAppCur>0 then
					intAppCur=cdbl(myCmd_("appAfterExpire"))
				else
					intAppCur=cdbl(myCmd_("appAfterExpire")) + (intAppCur - decBalancePast)
				end if
				
				intBalance=intLeaveDueCur - intAppCur
			end if
		end if	
	
	Else
		strError = objDatabase.strMessage
	End If		
	
'--------------------------------------------------
' Get user's fullname and jobtitle
'--------------------------------------------------
	Set objEmployee = New clsEmployee
	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	strFunction = "<a class='c' href='javascript:back_menu()' onMouseOver='self.status=&quot;Return to main menu page&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:window.history.back();' onMouseOver='self.status=&quot;Back&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Back</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:selstaff();' onMouseOver='self.status=&quot;Select employee to view annual leave&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Select Employee</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='javascript:logout()' onMouseOver='self.status=&quot;Log out timesheet system&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Log Out</a>&nbsp;&nbsp;&nbsp;<img height='5' src='../../images/dot.gif' width='5'>&nbsp;&nbsp;&nbsp;" & _
				  "<a class='c' href='#' onMouseOver='self.status=&quot;Help&quot;;return true' onMouseOut='self.status=&quot;&quot;;return true'>Help</a>&nbsp;&nbsp;&nbsp;"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle1	= "Annual Leave of <b>" & varFullName(0) & " - " & varFullName(1) & "</b>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate(strTitle, strFunction, arrPageTemplate, "templates/template1/")

%>	
<html>
<head>
<meta HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css">

</head>

<script language="javascript" src="../../library/library.js"></script>

<script LANGUAGE="JavaScript">
<!--
var ns, ie, objNewWindow;
ns = (document.layers)? true:false
ie = (document.all)? true:false

function logout()
{
	URL = "../../logout.asp";
	if (ns)
		document.location = URL;
	else
	{
		window.document.frmtms.action = URL;
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function gopage()
{
	document.frmtms.action = "../../tools/preferences.asp";
	document.frmtms.submit();
}

function goback()
{
	if (ns)
		document.location = "timesheet.asp";
	else
	{
		window.document.frmtms.action = "timesheet.asp";
		window.document.frmtms.target = "_self";
		window.document.frmtms.submit();
	}	
}

function selstaff()
{
	window.status = "";
 
	strFeatures = "top="+(screen.height/2-225)+",left="+(screen.width/2-230)+",width=490,height=325,toolbar=no," 
              + "menubar=no,location=no,directories=no,resizable=no,scrollbars=yes";
              
	if((objNewWindow) && (!objNewWindow.closed))
		objNewWindow.focus();	
	else 
	{
		objNewWindow = window.open('tms_select_staff.asp?view=l', "MyNewWindow", strFeatures);
	}
	window.status = "Opened a new browser window.";  
}

function back_menu()
{
	window.document.frmtms.action = "tms_list_staff.asp?b=1";
	window.document.frmtms.target = "_self";
	window.document.frmtms.submit();
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
<table width="780" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6" background="../../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
    <td valign="top" height="100%" width="772">
      <table width="100%" border="0" cellspacing="1" cellpadding="0" align="center" style="height:79%" height="365">
<%If strError <> "" Then%>
		<tr>
          <td height="80"> 
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr bgcolor="#E7EBF5"> 
				<td class="blue" align="center"><%=strError%></td>
			  </tr>
			</table>
		  </td>	
		</tr>
<%End If%>			    
	    <tr> 
          <td height="80"> 
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr> 
				<td class="title" align="center">View Annual Leave until &nbsp 1/<%=month(date)%>/<%=year(date)%></td>
			  </tr>
  			  <tr> 
				<td class="blue-normal" align="center" bgcolor="#FFFFFF" height="20"><%=strTitle1%></td>
			  </tr>			  
			</table>
		  </td>
		</tr>
		<tr> 
		  <td valign="top"> 
			<table width="80%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" >
  			  
			 
			  <tr> 
           		<td bgcolor="#8FA4D3"> 
			      <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
                    <tr bgcolor="#617DC0" height="25"> 
                      <td class="white" width="25%">&nbsp;</td> 
                      <td class="white" align="center" width="25%"><b>Leave Due (hours) </b></td>
                      <td class="white" align="center" width="25%"><b>Application (hours) </b></td>
                      <td class="white" align="center" width="25%"><b>Balance (hours) </b></td>
					</tr>
					<tr bgcolor="#E7EBF5" height="25" > 
					  <td valign="middle" width="25%" class="blue-normal">&nbsp;&nbsp;<%=year(date)%></td>
					  <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(intLeaveDueCur,2)%></td>
	                  <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(intAppCur,2)%></td>
		              <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(cdbl(intLeaveDueCur)-cdbl(intAppCur))%></td>
			        </tr>
			        <%if not isnull(myCmd("expiredDate")) then%>
						    <%if date<dateExipred then%>
			        <tr bgcolor="#E7EBF5" height="25"> 
                      <td valign="middle" width="25%" class="blue-normal">&nbsp;&nbsp;The end of <%=year(date)-1%></td>                    					
					  <td valign="middle" align="center" width="25%" class="blue-normal">&nbsp;&nbsp;</td>
	                  <td valign="middle" align="center" width="25%" class="blue-normal">&nbsp;&nbsp;</td>
		              <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(cdbl(decBalancePast),2)%></td>
			        </tr>
							<%end if
					else%>
					<tr bgcolor="#E7EBF5" height="25"> 
					  <td valign="middle" width="25%" class="blue-normal">&nbsp;&nbsp;The end of <%=year(date)-1%></td>                    					
					  <td valign="middle" align="center" width="25%" class="blue-normal">&nbsp;&nbsp;</td>
					  <td valign="middle" align="center" width="25%" class="blue-normal">&nbsp;&nbsp;</td>
					  <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(cdbl(decBalancePast),2)%></td>
					</tr>
					<%end if%>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td >
			  	  <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
				  	 <tr height="25"> 
                      <td valign="middle" width="25%" class="blue">&nbsp;</td>                    					
					  <td valign="middle" width="25%" class="blue">&nbsp;</td>
	                  <td valign="middle" align="right" width="25%" class="blue">Real balance: </td>
		              <td valign="middle" align="center" width="25%" class="blue"><%=formatnumber(intBalance,2)%>&nbsp;(hours)</td>
			        </tr>
			        <tr> 
                      <td valign="middle" width="25%" class="blue">&nbsp;</td>                    					
					  <td valign="middle" width="25%" class="blue">&nbsp;</td>
	                  <td valign="middle" align="right" width="25%" class="blue">&nbsp;</td>
		              <td valign="middle" align="center" width="25%" class="red"><b><%=formatnumber(intBalance/cdbl(intWorkHours),2)%>&nbsp(days)</b></td>
			        </tr>
				  </table>
			  	</td>
			  </tr>
			  			  <%if not isnull(myCmd("expiredDate")) then                      
					 if date<dateExipred then%>
			   <tr> 
           			<td bgcolor="#FFFFFF" class="blue"> 
           				&nbsp;* Annual leave balance of <%Response.Write(year(Date)-1)%> will be expired on <%=myCmd("expiredDate")%>
					</td>
			  </tr>	
			  <tr> 
           			<td bgcolor="#FFFFFF" class="blue">&nbsp;* Balance in <%=year(date)-1%> until now:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           				<span class="red"><%if cdbl(decBalancePast)>cdbl(intAppCur) then%> 
           					<%=formatnumber(cdbl(decBalancePast-intAppCur)/cdbl(intWorkHours),2)%> <%else%>0.00<%end if%>&nbsp; (days)</span>
           			</td>
			  </tr>	
					<%end if
			  end if%>
			  
			  
			</table>
		  </td>
		</tr>
		
<%
	Dim dblSick,dblSickWithCer
	dblSick=0
	dblSickWithCer=0
	strQuery="SELECT SUM(Hours) as SUMHours, eventID FROM ATC_Timesheet WHERE StaffID=" & intStaffID & " and (EventID=6 OR EventID=9) GROUP BY eventID "
	If objDatabase.runQuery(strQuery) Then
		If not objDatabase.noRecord then
			Do until objDatabase.rsElement.EOF
				if cint(objDatabase.rsElement("eventID"))=6 then
					dblSick=cdbl(objDatabase.rsElement("SUMHours"))
				elseif cint(objDatabase.rsElement("eventID"))=9 then
					dblSickWithCer=cdbl(objDatabase.rsElement("SUMHours"))
				end if
				objDatabase.rsElement.MoveNext
			Loop
		end if
	end if
%>		
		<tr>
            <td height="50" align="center" valign="middle" class="title">View Sick Leave in <%=year(date)%> </td>
        </tr>
		<tr> 
		  <td valign="top"> 
			<table width="80%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399">
				
			  <tr> 
           		<td bgcolor="#8FA4D3"> 
			      <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
                    <tr bgcolor="#617DC0" height="25"> 
                      <td class="white" width="25%">&nbsp;</td> 
                      <td class="white" align="center" width="25%"><b>without certificate<sup>(1)</sup>(hours)</b></td>
                      <td class="white" align="center" width="25%"><b>with certificate<sup>(2)</sup><br>(hours)</b></td>
                      <td class="white" align="center" width="25%"><b>Total (hours) </b></td>
					</tr>
					<tr bgcolor="#E7EBF5" height="25" > 
					  <td valign="middle" width="25%" class="blue-normal">&nbsp;&nbsp;</td>
					  <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(dblSick,2)%></td>
	                  <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(dblSickWithCer,2)%></td>
		              <td valign="middle" align="center" width="25%" class="blue-normal"><%=formatnumber(dblSick + dblSickWithCer,2)%></td>
			        </tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td >
			  	  <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
			        <tr> 
                      <td valign="middle" width="25%" class="blue">&nbsp;</td>                    					
					  <td valign="middle" width="25%" class="blue">&nbsp;</td>
	                  <td valign="middle" align="right" width="25%" class="blue">&nbsp;</td>
		              <td valign="middle" align="center" width="25%" class="red"><b><%=formatnumber((dblSick + dblSickWithCer)/cdbl(intWorkHours),2)%>&nbsp(days)</b></td>
			        </tr>
					<tr> 
                        <td valign="middle" class="blue">&nbsp;</td>
                        <td valign="middle" class="blue">&nbsp;</td>
                        <td valign="middle" align="right" class="blue">&nbsp;</td>
                        <td valign="middle" align="center" class="red">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="4" valign="middle" class="blue-normal"><sup>(1)</sup> Maximum 2 days per year with full pay</td>
                      </tr>
                      <tr>
                        <td colspan="4" valign="middle" class="blue-normal"><sup>(2)</sup> 
							&#8226; <u>To local staff</u> : if taking sick leave of more than 2 days/year, staff must submit sick leave certificate from designated hospital / medical center which has been registered in Health Insurance Card.
                        </td>
                      </tr>
                      <tr>
                        <td colspan="4" valign="middle" class="blue-normal"> 
							&nbsp;&nbsp;&nbsp;&nbsp;&#8226; <u>To foreign staff</u> : the expatriate must submit doctor certificate either from SOS, Columbia Asia, FV Hospital, or from private doctor. </td>
                      </tr>
				  </table>
				
			  	</td>
			  </tr>
			  
			</table>
		  </td>
		</tr>
		
      </table>
    </td>
    <td width="2" background="../../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
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
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">
<input type="hidden" name="txthidden" value="<%=intStaffID%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">

</form>
</body>
</html>
