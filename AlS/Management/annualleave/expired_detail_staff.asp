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
			
			strApplyDate=rsSrc("ApplyYear")
			strExpired="None Expired"

			if Cint(rsSrc("ExpiredDay")) <>0 then	strExpired= "Expired Date: <b>" & rsSrc("ExpiredDay") & "-" & MonthName(rsSrc("ExpiredMonth"),true) & "</b>"
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:UpdateInformation(" & rsSrc("StaffExpiredRuleID") & ");' " &_
						"class='c' OnMouseOver = 'self.status=&quot;Update Annual Leave Information &quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & strApplyDate & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & strExpired & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("KeepPassYear") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("NoteForExpire") & "</td>"		
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
function GetExRuleListBox(rsSrc,intExRuleID)
	dim strOut
	
'Response.Write 	intRuleID
	strOut=""
	
	if (rsSrc.RecordCount>0) OR not rsSrc.EOF then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
			
			strExpire="None Expired"
'Response.Write 	(cint(rsSrc("ExpiredDay"))=0) & "-" & (cint(rsSrc("KeepPassYear"))=0) & "-" & (Cint(rsSrc("ExpiredDay")) <>0 AND Cint(rsSrc("KeepPassYear"))<>0) & "<br>"
			If Cint(rsSrc("ExpiredDay")) <>0 OR Cint(rsSrc("KeepPassYear"))<>0 then
				if Cint(rsSrc("ExpiredDay")) <>0 then	strExpire= rsSrc("ExpiredDay") & "-" & MonthName(rsSrc("ExpiredMonth"),true) 
				strExpire=strExpire & " - Keep:" & rsSrc("KeepPassYear") & " days"		
			end if
			
			
			strSelect=""
			if cint(rsSrc("RuleYearlyID")) =cint(intExRuleID) then strSelect="selected"
			
			strOut=strOut & "<option value='" & rsSrc("RuleYearlyID") & "' " & strselect & " >" & strExpire  & "</option>"
			rsSrc.MoveNext
		loop
		
	end if

	GetExRuleListBox=strOut
end function

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
	intBankID=0
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	intStaffID = Request.Form("txthidden")
				

	if Request.QueryString("act") = "save" then
		intID=Request.Form("txtID")
		
		intRuleID=Request.Form("lbRuleEx")	
		intApplyYear=Request.Form("lstyearExF")
		strNote=Request.Form("txtNote")
		
		fgDel=Request.Form("fgstatus")

		if fgDel<>"D" then
			if Cint(intID)=-1 then
				'Add new
				strQuery = "INSERT INTO ATC_EmployeeExpiredRule (StaffID,ApplyYear,RuleYearlyID,NoteForExpire) VALUES (" & _
									 intStaffID & "," & intApplyYear & ",'" & intRuleID & "','" & replace(strNote,"'","''") & "')"	
			else
				'Update
				strQuery = "UPDATE ATC_EmployeeExpiredRule " & _
								"SET StaffID = " & intStaffID & _
									",RuleYearlyID = " & intRuleID & _
									",ApplyYear = '" & intApplyYear & "'" & _
									",NoteForExpire = '" & replace(strNote,"'","''") & "' " & _
								"WHERE StaffExpiredRuleID=" & intID
			end if
		else
			strQuery = "DELETE FROM ATC_EmployeeExpiredRule WHERE StaffExpiredRuleID=" & intID
		end if

'Response.Write strQuery		
		strError=ExecuteSQL(strQuery)
	else
		intID=-1
		intRuleID=0
		dateApplyDate=Date()
		strNote=""	
	End If
'--------------------------------------------------
' Get History of Annual Leave
'--------------------------------------------------

	
	strSql="SELECT a.*,ExpiredDay,ExpiredMonth,KeepPassYear   FROM ATC_EmployeeExpiredRule a " & _
				"INNER JOIN ATC_AnnualLeaveYearlyRule b ON a.RuleYearlyID=b.RuleYearlyID " & _
				"WHERE StaffID=" & intStaffID & " ORDER BY ApplyYear DESC"

	Call GetRecordset(strSql,rsStaffAnnualLeave)
		
	strLast=OutBody(rsStaffAnnualLeave)
	
	If Request.QueryString("act") = "show"  Then			
		if rsStaffAnnualLeave.RecordCount>0 then		
			intID=Request.Form("txtID")		
			rsStaffAnnualLeave.MoveFirst
			rsStaffAnnualLeave.Filter="StaffExpiredRuleID=" & intID			
			if rsStaffAnnualLeave.RecordCount>0 then
				intRuleID=rsStaffAnnualLeave("RuleYearlyID")
				dateApplyDate=cdate(rsStaffAnnualLeave("ApplyYear"))
				strNote=rsStaffAnnualLeave("NoteForExpire")
			end if
		end if			
	end if

'--------------------------------------------------
' Get Expired Rule
'--------------------------------------------------	
strSql="SELECT * FROM ATC_AnnualLeaveYearlyRule ORDER BY RuleYearlyID"
Call GetRecordset(strSql,rsExRule)

strRuleExListBox=GetExRuleListBox(rsExRule,intRuleID)
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

	arrlstFrom(2) = selectyear("lstyearF", dateApplyDate, 1999, year(date())+2, 0)

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
	document.frmreport.action = "expired_detail_staff.asp?act=show";
	document.frmreport.submit();
}
	
function addrule()
{
		
	document.frmreport.txtID.value = -1;	
	window.document.frmreport.lbRule.value=0
	window.document.frmreport.txtNote.value = "";
	
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
	window.document.frmreport.lbRule.focus();
}
	
function savedata()
{
	
	if (checkdata())
	{
	    //alert (window.document.frmreport.lbRuleEx.value)
		window.document.frmreport.action = "expired_detail_staff.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "expired_detail_staff.asp?act=save"			
	window.document.frmreport.submit();
}

function ViewAnnualLeave()
{
	window.document.frmreport.action = "staff_view_leave.asp";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function IndividualCases()
{
	window.document.frmreport.action = "annual_individual_cases.asp";
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
	
	if (window.document.frmreport.lbRuleEx.value==0)
	{
		alert("Please select a Annual Leave rule from list.");
		document.frmreport.lbRule.focus();
		return false	
		
	}	
	
	return true	
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
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="annual_list_staff.asp" onMouseOver="self.status='';return true">Employee List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;</td>
                  <td class="blue" height="30" align="right" width="77%">
					<table width="360" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                      <tr> 
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" >
                          <div align="center" class="blue"><a href="javascript:ViewAnnualLeave()" onMouseOver="self.status='Please click here to view staff Annual Leave.';return true" onMouseOut="self.status='';return true" class="b">View Annual Leave</a></div>
                        </td>
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue"><a href="javascript:specialcases()" onMouseOver="self.status='Please click here to view details .';return true" onMouseOut="self.status='';return true" class="b">Staff Annual Leave Detail</a></div>
                        </td>
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" >
                          <div align="center" class="blue"><a href="javascript:IndividualCases()" onMouseOver="self.status='Please click here to view some exception.';return true" onMouseOut="self.status='';return true" class="b">Individual cases</a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>                
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Staff Expired rule Detail</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="25%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="20%">Full Name</td>
                              <td valign="middle" width="35%" class="blue"><%=strFullName%></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
     
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Select an Expired rule</td>
                              <td valign="middle" class="blue"><select id='lbRuleEx' size='1' name='lbRuleEx' class='blue-normal'><option value='0'>&nbsp;</option>
								<%=strRuleExListBox%></select></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Apply From</td>
                              <td valign="middle" class="blue">
            <%											Response.Write Replace(arrlstFrom(2),"lstyearF","lstyearExF")%></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                              
                            </tr>  
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" class="blue">
								<input type="text" name="txtNote" maxlength="200" class="blue-normal" size="20" style='width:95%' value="<%=strNote%>">
            
							</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr>                                   
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
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
                                </table>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="15%">Apply Year</td>
                              <td class="blue" align="center" width="30%">Expired date</td>
                              <td class="blue" align="center" width="20%">Keep pass year</td>
                              <td class="blue" align="center" width="35%">Note</td>                                                            
                            </tr>
<%Response.Write strLast%>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue-normal">&nbsp;&nbsp;* Click on Apply year to update</td>
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
