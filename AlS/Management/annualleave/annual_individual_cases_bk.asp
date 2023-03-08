<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError,rsBanks,intBankID,strAcc,strNote,blnProbation
	Dim arrlstFrom(2),arrlongmon

function OutBody(rsSrc)
	dim strOut
	dim i
	
	strOut=""
	i=0

	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
			
			intYearApply=cint(rsSrc("ApplyYear"))
			strExpired=""

			if cint(rsSrc("ExpiredDay"))>0 then
				strExpired="Expired Date: <b>" & rsSrc("ExpiredDay") & "-" & MonthName(rsSrc("ExpiredMonth"),true) & "</b><br>Keep Pass year:<b>" & rsSrc("KeepPassYear") & " days</b>" 
			end if
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:UpdateInformation(" & rsSrc("IndividualRuleID") & ");' " &_
						"class='c' OnMouseOver = 'self.status=&quot;Update Annual Leave Information &quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & intYearApply & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & FormatNumber(rsSrc("RatePerYear"),2) & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & strExpired & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("MoreHours") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("RuleNote") & "</td>"		
			strOut=strOut & "</tr>"
			i=i+1	
			rsSrc.MoveNext
		loop
		
	end if
	
	OutBody=strOut
End Function

'***************************************************************
'
'***************************************************************

function ExecuteSQL(strSql)

	dim strConnect,ret,strMessage
	dim objDb	

	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
		
	If objDb.dbConnect(strConnect) then
			
		ret = objDb.runActionQuery(strQuery)
				
		if ret=false then				
			strMessage = objDb.strMessage
		else
			strMessage="Update successfully."
		end if
			  
	else
		strMessage=objDb.strMessage
	end if
	
	ExecuteSQL=strMessage
end function
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	intStaffID = Request.Form("txthidden")

	if Request.QueryString("act") = "save" then
		intID=Request.Form("txtID")
		
		intYearApply=Request.Form("lstyearF")
		dblMore=Request.Form("txtmore")
		if dblMore="" then dblMore=0
		
		dblRate=Request.Form("txtRatePerYear")
		intMonth=0
		intDay=0
		dblKeepPassYear=0
		intExpired=Request.Form("radExpired")
	
		if Cint(intExpired)=1 then
			intMonth=cint(Request.Form("lstMonthF"))
			intDay=cint(Request.Form("lstdayF"))
			
			dblKeepPassYear=Request.Form("txtKeepPassYear")
		end if
	
		strNote=Request.Form("txtNote")
						
		fgDel=Request.Form("fgstatus")
		
		if fgDel<>"D" then
			if Cint(intID)=-1 then
				'Add new
				strQuery = "INSERT INTO ATC_AnnualLeaveIndividualRule(StaffID,RatePerYear,ApplyYear,ExpiredDay,ExpiredMonth,KeepPassYear,MoreHours,RuleNote) VALUES (" & _
								intStaffID & "," & dblRate & "," & intYearApply  & "," & intDay & "," & intMonth  & "," & dblKeepPassYear & "," & dblMore & ",'" & replace(strNote,"'","''") & "')"	

			else
				'Update
				strQuery = "UPDATE ATC_AnnualLeaveIndividualRule " & _
								"SET RatePerYear = " & dblRate & _
								   ",ApplyYear =  " & intYearApply & _
								   ",ExpiredDay = " & intDay & _
								   ",ExpiredMonth = " & intMonth & _
								   ",KeepPassYear = " & dblKeepPassYear & _
								   ",MoreHours = " & dblMore & _
								   ",RuleNote = '" & replace(strNote,"'","''") & " '" & _
							"WHERE IndividualRuleID=" & intID
				
				
			end if
		else
			strQuery = "DELETE FROM ATC_AnnualLeaveIndividualRule WHERE IndividualRuleID=" & intID
		end if
		
		strError=ExecuteSQL(strQuery)

	End If
'--------------------------------------------------
' Get History of Annual Leave
'--------------------------------------------------

	
	strSql="SELECT IndividualRuleID,StaffID,RatePerYear,ApplyYear,ExpiredDay,ExpiredMonth,KeepPassYear,MoreHours, RuleNote " &_
				"FROM ATC_AnnualLeaveIndividualRule " & _
				"WHERE StaffID=" & intStaffID & " ORDER BY ApplyYear DESC"
	

	Call GetRecordset(strSql,rsInviCases)
		
	strLast=OutBody(rsInviCases)
	
	If Request.QueryString("act") = "show"  Then			
		if rsInviCases.RecordCount>0 then		
			intID=Request.Form("txtID")		
			rsInviCases.MoveFirst
			rsInviCases.Filter="IndividualRuleID=" & intID			
			if rsInviCases.RecordCount>0 then
		
				intYearApply= rsInviCases("ApplyYear")
				dblMore=rsInviCases("MoreHours")
				if cdbl(dblMore)=0 then dblMore=""
		
				dblRate=rsInviCases("RatePerYear")
				intMonth=rsInviCases("ExpiredMonth")				
				intDay=rsInviCases("ExpiredDay")
				dblKeepPassYear=rsInviCases("KeepPassYear")				
				if cint(intDay)=0 then
					intMonth=month(Date())
					intDay=Day(Date())
					dblKeepPassYear=""
				end if				
				
				'intExpired=Request.Form("radExpired")				
				strNote=rsInviCases("RuleNote")
				
			end if
		end if			
	else		
		intID=-1
		intYearApply= year(Date())
		dblMore=""					
		dblRate=""
		intMonth=month(Date())
		intDay=Day(Date())
		dblKeepPassYear=""							
		strNote=""
	end if

'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)					
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
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	If strChoseMenu = "" Then strChoseMenu = "AE"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")


	arrlstFrom(0) = selectmonth("lstmonthF",intMonth , -1)
	arrlstFrom(1) = selectday("lstdayF", intDay, -1)
	arrlstFrom(2) = selectyear("lstyearF", intYearApply, 1999, year(date())+2, 0)

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--
	
function UpdateInformation(r)
{
	document.frmreport.txtID.value = r;
	document.frmreport.action = "annual_individual_cases.asp?act=show";
	document.frmreport.submit();
}
	
function addrule()
{		
	
	window.document.frmreport.txtID.value = -1;	
	window.document.frmreport.txtmore.value="";
	window.document.frmreport.txtNote.value = "";
	window.document.frmreport.txtRatePerYear.value = "";	
	window.document.frmreport.txtKeepPassYear.value = "";
	document.frmreport.radExpired[1].checked = true;
	
	showhide('expireddate',0)
	
	var i=0;	
	var intCount = window.document.frmreport.lstdayF.options.length;	
	for (i = 0; i < intCount; i++) {

		if (i==<%=day(Date())%>)
		{
			
			window.document.frmreport.lstdayF.options[i-1].selected = true;
		}
	}
	
	intCount = window.document.frmreport.lstmonthF.options.length;	
	for (i = 0; i < intCount; i++) {

		if (i==<%=month(Date())%>)
		{
			
			window.document.frmreport.lstmonthF.options[i-1].selected = true;
		}
	}
	
	intCount = window.document.frmreport.lstyearF.options.length;	
	
	for (i = 0; i < intCount; i++) {
	
		if (window.document.frmreport.lstyearF.options[i].value==<%=year(Date())%>)
		{			
			window.document.frmreport.lstyearF.options[i].selected = true;
		}
	}	
	window.document.frmreport.lstyearF.focus();
	
	
}
	
function savedata()
{
	if (checkdata())
	{		
		window.document.frmreport.action = "annual_individual_cases.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "annual_individual_cases.asp?act=save"			
	window.document.frmreport.submit();
}

function back_menu()
{
	window.document.frmreport.action = "annual_list_staff.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function specialcases()
{
	window.document.frmreport.action = "annual_detail_staff.asp";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function checkdata()
{
	if (document.frmreport.txtmore.value != "") {
		if (isNaN(document.frmreport.txtmore.value)==true) {
			alert("Please enter a number for More hours data.");
			document.frmreport.txtmore.focus();
			return false;
		}
	}
	
	if (document.frmreport.txtRatePerYear.value == "") {
		alert("Please enter Rate Per Year.");
		document.frmreport.txtRatePerYear.focus();
		return false;
		}
	else
		if (isNaN(document.frmreport.txtRatePerYear.value)==true) {
			alert("Please enter a number.");
			document.frmreport.txtRatePerYear.focus();
			return false;
		}
		else if (document.frmreport.txtRatePerYear.value<=0) {
			alert("The Rate value must be greater than 0.");
			document.frmreport.txtRatePerYear.focus();
			return false;			
		}
	
	if (document.frmreport.radExpired[0].checked)
	{
		var dateFrom=document.frmreport.lstdayF.value + "/" + document.frmreport.lstmonthF.value + "/" + document.frmreport.lstyearF.value
		if (isdate(dateFrom)==false){
			alert("The date (" + dateFrom + ") is invalid.");
			document.frmreport.lstdayF.focus();
			return false;
		}
		
		
		if (document.frmreport.txtKeepPassYear.value == "") {
			alert("Please enter number of days.");
			document.frmreport.txtKeepPassYear.focus();
		return false;
		}
		else
		if (isNaN(document.frmreport.txtKeepPassYear.value)==true) {
			alert("Please enter a number.");
			document.frmreport.txtKeepPassYear.focus();
			return false;
		}
		else if (document.frmreport.txtKeepPassYear.value<0) {
			alert("The number of days must be greater than 0.");
			document.frmreport.txtKeepPassYear.focus();
			return false;			
		}
	}
	
	return true;
}

function showhide(layer_ref,val) { 
var state = 'none'; 

	if (val == 0) { 
		state = 'none'; 
	} 
	else { 
		state = 'block'; 
	} 
	if (document.all) { //IS IE 4 or 5 (or 6 beta) 
		eval( "document.all." + layer_ref + ".style.display = state"); 
	} 
	if (document.layers) { //IS NETSCAPE 4 or below 
		document.layers[layer_ref].display = state; 
	} 
	if (document.getElementbyId &&!document.all) { 
		hza = document.getElementbyId(layer_ref); 
		hza.style.display = state; 
	} 
} 	
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
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
	      <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%		If strError <> "" Then%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red" colspan="2">&nbsp;<b><%=strError%></b></td>
				</tr>
<%		End If%>				
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="javascript:back_menu()" onMouseOver="self.status='';return true">Employee List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;</td>
                  <td class="blue" height="30" align="right" width="77%">
					            <table width="150" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                                      <div align="center" class="blue"><a href="javascript:specialcases()" onMouseOver="self.status='Please click here to view details .';return true" onMouseOut="self.status='';return true" class="b">Staff Annual Leave Detail</a></div>
                                    </td>
                                  </tr>
                                </table>
                  </td>
                </tr>                
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Individual Cases for Annual Leave</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
                  <td bgcolor="#FFFFFF" valign="top">
					<table width="55%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#003399">
                      <tr> 
                        <td > <table width="100%" border="0" align="center" cellpadding="10" cellspacing="0" >
                            <tr> 
                              <td bgcolor="#C0CAE6" >
                              
								<table width="100%" border="0" cellspacing="5" cellpadding="0">
                                  <tr> 
                                    <td valign="middle" class="blue-normal" width="30%">&nbsp;&nbsp;Full Name </td>
                                    <td valign="middle" width="70%" class="blue"> <%=strFullName%> </td>
                                  </tr>
                                  <tr> 
									<td valign="middle" class="blue-normal">&nbsp;&nbsp;Apply for</td>
									<td valign="middle" class="blue"><% Response.Write replace(arrlstFrom(2)," onClick='CheckMode(this)'","")%></td>
                                  </tr>
                                  
								<tr> 
									<td valign="middle" class="blue-normal">&nbsp;&nbsp;More hours</td>
									<td valign="middle" class="blue"><input type="text" name="txtmore" maxlength="30" class="blue-normal" size="20" style='width:50%' value="<%=dblMore%>"></td>
                                  </tr>                                
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;&nbsp;Rate per year *</td>
                                    <td valign="middle" class="blue-normal"> 
                                      <input type="text" name="txtRatePerYear" maxlength="30" class="blue-normal" size="20" style='width:50%' value="<%=dblRate%>">
                                    </td>
                                  </tr>                                                                    
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;&nbsp;Expired *</td>
                                    <td valign="middle" class="blue-normal"> 
										<table width="100%">
											<tr>
												<td valign="middle" class="blue-normal">
													<input type="radio" name="radExpired" value="1" onclick="showhide('expireddate',1);" <%if (intDay>0) and (intID<>-1) then %>checked<%End If%>> &nbsp;Yes
												</td>
												<td valign="middle" class="blue-normal">
													<input type="radio" name="radExpired" value="0"  onclick="showhide('expireddate',0);" <%if intDay=0 or intID=-1 then %>checked<%End If%>> &nbsp;No
												</td>
											</tr>
										</table>
                                    </td>                                    
                                  </tr>
                                  <tr>
									<td valign="middle" class="blue-normal">&nbsp;&nbsp;</td>
									<td valign="middle" class="blue-normal"> 
									<div name="expireddate" id="expireddate" style="display: <%if intDay>0 and intID<>-1 then %>block<%else%>none<%End If%>;">
										<table width="100%">
											<tr>
												<td valign="middle" class="blue-normal">
													Expired Date *
												</td>
												<td valign="middle" align="left">
												<%
														Response.Write arrlstFrom(1)
														Response.Write arrlstFrom(0)%>
												</td>
											</tr>
											<tr>
												<td valign="middle" class="blue-normal">
													Keep Pass year *&nbsp;
												</td>
												<td valign="middle" align="left">
												<input type="text" name="txtKeepPassYear" maxlength="100" class="blue-normal" size="20" style='width:95%' value="<%=dblKeepPassYear%>">							
												</td>
											</tr>
										</table>
										</div> 
                                    </td>                                    
                                  </tr>                                
								<tr> 
									<td valign="middle" class="blue-normal">&nbsp;&nbsp;Note</td>
									<td valign="middle" class="blue"><input type="text" name="txtNote" maxlength="100" size="20" style='width:95%'class="blue-normal" value="<%=strNote%>"></td>
                                  </tr>                                    
                                  <tr> 
                                    <td valign="middle" class="blue-normal">&nbsp;</td>
                                    <td valign="middle" class="blue-normal">
										       <table width="180" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60"> 
                                      <div align="center" class="blue"><a href="javascript:addrule()" onMouseOver="self.status='Please click here to add new record';return true" onMouseOut="self.status='';return true" class="b">Add</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:savedata()" onMouseOver="self.status='Please click here to save changes';return true" onMouseOut="self.status='';return true" class="b">Save</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:deletedata()" onMouseOver="self.status='Please click here to delete this record';return true" onMouseOut="self.status='';return true" class="b">Delete</a></div>
                                    </td>
                                  </tr>
                                </table></td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table> </td>
                </tr>
          <tr> 
            <td height="100%"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">

                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    				<tr>
					<td>&nbsp;</td>
				</tr>
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="10%">Apply for</td>
                              <td class="blue" align="center" width="15%">Rate per year</td>
                              <td class="blue" align="center" width="25%">Expired Information</td>
                              <td class="blue" align="center" width="15%">More Hours</td>
                              <td class="blue" align="center" width="35%">Note</td>
                            </tr>
<%Response.Write strLast%>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue-normal">&nbsp;&nbsp;* Click on Apply date to update</td>
                            </tr>
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
<%
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
	Response.Write(arrTmp(1))
%>		

<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(2))    
%>
<input type="hidden" name="txthidden" value="<%=intStaffID%>">
<input type="hidden" name="fgstatus" value="">
<input type="hidden" name="txtID" value="<%=intID%>">
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">

</form>

</body>
</html>
