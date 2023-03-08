<!-- #include file = "../../../class/CEmployee.asp"-->
<!-- #include file = "../../../inc/createtemplate.inc"-->
<!-- #include file = "../../../inc/getmenu.asp"-->
<!-- #include file = "../../../inc/constants.inc"-->
<!-- #include file = "../../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError,rsBanks,intBankID,strAcc,strNote,blnProbation

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../../message.htm")
	End If					

	intUserID = session("USERID")
	intBankID=0
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
	intStaffID = Request.Form("txthidden")
				
'--------------------------------------------------
' Check ACCESS right
'--------------------------------------------------

	If Request.QueryString("act") = "" Or Request.QueryString("act") = "save" Then
		If objDatabase.dbConnect(strConnect) Then
			If Request.QueryString("act") = "save" And (Request.Form("fgstatus") <> "" Or Request.Form("fgstatus") = "") Then
'				Call checkData()
				intWID = Request.Form("lbWHour")
				varDate = split(Request.Form("txtdate"),"/")
				If Not IsEmpty(varDate) Then
					strDate = CDate(varDate(1) & "/" & varDate(0) & "/" & varDate(2))
				End If	
				
				strSalary = encode(Request.Form("txtsalary"),128)
				If Request.Form("txtsalarytax") <> "" Then
					strSalaryTax = "'" & encode(Request.Form("txtsalarytax"),128) & "'"
				Else
					strSalaryTax = "Null"
				End If
				strCurrency = Request.Form("lbcurrency")
				strCurrencyTax = Request.Form("lbcurrencytax")
				
				blnOverTime = Request.Form("chkovertime")
				blnOverTime= IIF(blnOverTime<>"",1,0)				
		
				blnProbation= Request.Form("chkProbation")
				blnProbation = IIF(blnProbation<>"",1,0)
								
				strAcc=Request.Form("txtBankAcc")
				if strAcc<>"" then
					intBankID=Request.Form("lstBank")
					strNote=Request.Form("txtBankNote")
				else
					strAcc=""
					intBankID=""
					strNote=""
				end if
						
				If Request.Form("fgstatus") = "A" Then
					strSQL = "INSERT INTO ATC_SalaryStatus(StaffID, SalaryDate, WorkingHourID, Salary, Currency, SalaryTax, CurrencyTax, fgOverTimePay,BankID,AccountNo,BankDetail,fgProbation) VALUES" & _
								"(" & intStaffID & ", '" & strDate & "', " & intWID & ", '" & strSalary & "', '" & strCurrency & "', " & strSalaryTax & _
								", '" & strCurrencyTax & "', " & blnOverTime & "," & IIF(intBankID<>"", intBankID ,"null") & "," & IIF(strAcc<>"","'" & strAcc & "'","null") & "," & IIF(strNote<>"","'" & strNote & "'","null") & "," & blnProbation & ")"
				ElseIf Request.Form("fgstatus") = "E" Or Request.Form("fgstatus") = "" Then
					strSQL = "UPDATE ATC_SalaryStatus SET WorkingHourID=" & intWID & ", Salary='" & strSalary & "', SalaryTax=" & strSalaryTax & _
								", Currency='" & strCurrency & "', CurrencyTax='" & strCurrencyTax & "', fgOverTimePay=" & blnOverTime & _
								",BankID=" & IIF(intBankID<>"", intBankID ,"null") & ",AccountNo=" & IIF(strAcc<>"","'" & strAcc & "'","null") & ",BankDetail=" & IIF(strNote<>"","'" & strNote & "'","null") & _
								",fgProbation=" & blnProbation & _
								" WHERE StaffID=" & intStaffID & " AND SalaryDate='" & strDate & "'" 
				ElseIf Request.Form("fgstatus") = "D" Then								
					strSQL = "DELETE FROM ATC_SalaryStatus WHERE StaffID="& intStaffID & " AND SalaryDate='" & strDate & "'" 
				End if

				If objDatabase.runActionQuery(strSQL) Then
					strError = "Update successful."
				Else
					strError = objDatabase.strMessage
					if InStr(1,strError,"PK_ATC_SalaryStatus") then strError="The salary date is already in system. Please try with another."
				End If	
			End If

			strSQL = "SELECT UserID FROM zright WHERE UserID=" & intUserID
			If (objDatabase.runQuery(strSQL)) Then
				If Not objDatabase.noRecord Then				 
					strSQL = "SELECT * FROM ATC_SalaryStatus WHERE StaffID=" & intStaffID & " ORDER BY SalaryDate DESC"
					If (objDatabase.runQuery(strSQL)) Then
						If Not objDatabase.noRecord Then				 
							varSalary		= objDatabase.rsElement.GetRows
							session("varSalary") = varSalary

							If Day(varSalary(1,0)) < 10 Then
								sday		= "0" & Day(varSalary(1,0))
							Else
								sday		= Day(varSalary(1,0))	
							End If
							If Month(varSalary(1,0)) < 10 Then
								smonth		= "0" & Month(varSalary(1,0))
							Else
								smonth		= Month(varSalary(1,0))	
							End If	 
							strDate			= sday & "/" & smonth & "/" & Year(varSalary(1,0))

							strSalary		= decode(varSalary(3,0),128)
							strCurrency		= varSalary(5,0)
							strSalaryTax	= decode(varSalary(4,0),128)
							strCurrencyTax	= varSalary(6,0)
							blnOvertime =IIF(varSalary(7,0),1,0)
							
							intwID			= varSalary(2,0)
							
							intBankID=0
							if not isnull(varSalary(8,0)) then intBankID=cint(varSalary(8,0))
							
							strAcc=varSalary(9,0)
							strNote=varSalary(10,0)
							blnProbation =IIF(varSalary(11,0),1,0)
								
						End If	
					Else
						strError = objDatabase.strMessage
					End If		
				Else																' No record
					strError1 = "Sorry! You don't have right on this page"
				End If
			Else
				strError = objDatabase.strMessage
			End If			
		Else
			strError = objDatabase.strMessage
		End If	
		
	Else

		varSalary = session("varSalary")
		intRow = Request.QueryString("r")
			
		If Day(varSalary(1,intRow)) < 10 Then
			sday		= "0" & Day(varSalary(1,intRow))
		else
			sday		= Day(varSalary(1,intRow))
		End If
		If Month(varSalary(1,intRow)) < 10 Then
			smonth		= "0" & Month(varSalary(1,intRow))
		else
			smonth		= Month(varSalary(1,intRow))
		End If	 
		strDate			= sday & "/" & smonth & "/" & Year(varSalary(1,intRow))
		strSalary		= decode(varSalary(3,intRow),128)
		strCurrency		= Trim(varSalary(5,intRow))
		strSalaryTax	= decode(varSalary(4,intRow),128)
		strCurrencyTax	= Trim(varSalary(6,intRow))
		blnOvertime =IIF(varSalary(7,intRow),1,0)
		
		intwID			= varSalary(2,intRow)
		
		intBankID=0
		if not isnull(varSalary(8,intRow)) then intBankID=cint(varSalary(8,intRow))
							
		strAcc=varSalary(9,intRow)
		strNote=varSalary(10,intRow)
		blnProbation =IIF(varSalary(11,intRow),1,0)
	End If
	
'--------------------------------------------------
' Initialize workinghour array
'--------------------------------------------------
	
	If isEmpty(session("varWHour")) = False Then
		varWHour = session("varWHour")
		intNum = Ubound(varWHour,2)
	Else
		If objDatabase.dbConnect(strConnect) Then			
			strSQL = "SELECT * FROM ATC_WorkingHours ORDER BY Description"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					varWHour = objDatabase.rsElement.GetRows
					intNum = Ubound(varWHour,2)					
					session("varWHour") = varWHour
					objDatabase.closeRec
				End If
			Else
				Response.Write objDatabase.strMessage
			End If
		Else
			Response.Write objDatabase.strMessage		
		End If
	End If	

	Set objDatabase = Nothing
	
'--------------------------------------------------
' End Of initializing workinghour array
'--------------------------------------------------
strSql="SELECT * FROM ATC_Banks WHERE fgActivate=1 ORDER BY BankName"
Call GetRecordset(strSql,rsBanks)
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
				"<img src='../../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu = "" then strChoseMenu = "AD"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../../templates/template1/", "ats_menu.htm")


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

<link rel="stylesheet" href="../../../timesheet.css" type="text/css">
<script language="javascript" src="../../../library/library.js"></script>

<script language="javascript">
<!--
	
function showdata(r)
{
	document.frmreport.fgstatus.value = "E";
	document.frmreport.action = "sal_detail_staff.asp?act=show&r=" + r;
	document.frmreport.submit();
}
	
function addsal()
{
	window.document.frmreport.fgstatus.value = "A"
	window.document.frmreport.txtdate.value = "";
	window.document.frmreport.txtsalary.value = "";
	window.document.frmreport.txtsalarytax.value = "";
	window.document.frmreport.chkovertime.checked = false;
	window.document.frmreport.txtBankAcc.value = "";
	window.document.frmreport.txtBankNote.value = "";
	window.document.frmreport.lstBank.value=0
	window.document.frmreport.txtdate.focus();
}
	
function savedata()
{
	if (checkdata())
	{
		window.document.frmreport.action = "sal_detail_staff.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletesal()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "sal_detail_staff.asp?act=save"			
	window.document.frmreport.submit();
}

function back_menu()
{
	window.document.frmreport.action = "sal_list_staff.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function checkdata()
{
	if (document.frmreport.txtdate.value == "")
	{
		alert("Please enter the salary date.");
		document.frmreport.txtdate.focus();
		return false;
	}
	
	if (document.frmreport.txtsalary.value =="")
	{
		alert("Please enter the salary value.");
		document.frmreport.txtsalary.focus();
		return false;
	}
	
	if (isdate(document.frmreport.txtdate.value)==false)
	{
		alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
		document.frmreport.txtdate.focus();
		return false;
	}

	if (isNaN(document.frmreport.txtsalary.value) ==  true) 
	{
		alert("The salary value must be a number");
		document.frmreport.txtsalary.focus(); 
		return false;
	}
	
	if (document.frmreport.txtsalarytax.value !="")
	{
		if (isNaN(document.frmreport.txtsalarytax.value) ==  true) 
		{
			alert("The salary tax must be a number");
			document.frmreport.txtsalarytax.focus(); 
			return false;
		}
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
<%
	If strError1 = "" Then
%>        
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
                  <td class="blue" height="30" align="right" width="77%">&nbsp;</td>
                </tr>
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Salary Detail</td>
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
                              <td valign="top" width="13%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="14%">Full Name</td>
                              <td valign="middle" width="21%" class="blue"><%=strFullName%></td>
                              <td valign="top" width="4%" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left" width="14%">&nbsp;</td>
                              <td valign="middle" class="blue" width="34%" align="left">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Working Hours</td>
                              <td valign="middle" class="blue">
                              <select id="lbwhour" size="1" name="lbwhour" class="blue-normal">
<%		If intNum >= 0 Then
			For ii = 0 To intNum%>                    
								<option <%If CInt(intwID)=CInt(varWHour(0,ii)) Then%> selected <%End If%> value="<%=varWHour(0,ii)%>"><%=varWHour(2,ii)%></option>
<%			Next
		End If	%>
		 				     </select></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left">Salary</td>
                              <td valign="middle" class="blue" align="left"> 
                                <input type="text" name="txtsalary" class="blue-normal" size="20" style="width:80" value="<%=strSalary%>">
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="13%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="14%">Salary Date </td>
                              <td valign="middle" width="21%" class="blue-normal"> 
                                <input type="text" name="txtdate" class="blue-normal" size="20" style="width:80" value="<%=strDate%>">
                              </td>
                              <td valign="top" width="4%" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left" width="14%">Currency</td>
                              <td valign="middle" class="blue" width="34%" align="left"> 
                                <select id="lbcurrency" size="1" name="lbcurrency" class="blue-normal">
                                  <option value="USD" <%If strCurrency = "USD" Then%> selected <%End if%>>USD</option>
                                  <option value="VND" <%If strCurrency = "VND" Then%> selected <%End if%>>VND</option>
                                </select>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="13%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="14%">Probation</td>
                              <td valign="middle" width="21%" class="blue-normal"> 
                                <input type="checkbox" name="chkProbation" value="1" <%If CInt(blnProbation) = 1 Then%> checked <%End If%>>
                              </td>
                              <td valign="top" width="4%" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left" width="14%">Salary Tax </td>
                              <td valign="middle" class="blue" width="34%" align="left"> 
                                <input type="text" name="txtsalarytax" class="blue-normal" size="20" style="width:80" value="<%=strSalarytax%>">
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="13%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="14%">Overtime Pay </td>
                              <td valign="middle" width="21%" class="blue-normal"> 
                                <input type="checkbox" name="chkovertime" value="1" <%If CInt(blnOvertime) = 1 Then%> checked <%End If%>>
                              </td>
                              <td valign="top" width="4%" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left" width="14%">Currency Tax </td>
                              <td valign="middle" class="blue" width="34%" align="left"> 
                                <select id="lbcurrencytax" size="1" name="lbcurrencytax" class="blue-normal">
                                  <option value="USD" <%If strCurrencyTax = "USD" Then%> selected <%End if%>>USD</option>
                                  <option value="VND" <%If strCurrencyTax = "VND" Then%> selected <%End if%>>VND</option>
                                </select>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" colspan=6" class="blue">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
								<td valign="top" class="blue">&nbsp;</td>
                              <td valign="top" colspan="5" class="blue">Bank Information</td>
                            </tr>
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" colspan=6" class="blue">&nbsp;</td>
                            </tr>                            
                             <tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Bank Account #: </td>
                              <td valign="middle" class="blue-normal"><span class="blue">
                                <input name="txtBankAcc" type="text" class="blue-normal" id="txtBankAcc" style="width:100%" value="<%=strAcc%>" size="20">
                              </span></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left">Bank Name </td>
                              <td valign="middle" class="blue" align="left"><span class="blue-normal">
                                <select id="lstBank" size="1" name="lstBank" class="blue-normal" style="width:70%">
<%								rsBanks.MoveFirst
								Do while not rsBanks.EOF%>                                
                                  <option value="<%=rsBanks("BankID")%>" <%if cint(rsBanks("BankID"))=cint(intBankID) then%>selected<%end if%>><%=rsBanks(1)%></option>
<%									rsBanks.MoveNext
								loop%>                                  
								 </select>
                              </span></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note:</td>
                              <td colspan="3" valign="middle" class="blue-normal"><span class="blue">
                                <input name="txtBankNote" type="text" class="blue-normal" id="txtBankNote" style="width:100%" value="<%=strNote%>" size="20">
                              </span></td>
                              <td valign="middle" class="blue" align="left">&nbsp;</td>
                            </tr>
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">&nbsp;</td>
                              <td valign="middle" class="blue-normal">&nbsp;</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                              <td valign="middle" class="blue-normal" align="left">&nbsp;</td>
                              <td valign="middle" class="blue" align="left">&nbsp;</td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="180" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60"> 
                                      <div align="center" class="blue"><a href="javascript:addsal()" onMouseOver="self.status='Please click here to add new record';return true" onMouseOut="self.status='';return true" class="b">Add</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:savedata()" onMouseOver="self.status='Please click here to save changes';return true" onMouseOut="self.status='';return true" class="b">Save</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:deletesal()" onMouseOver="self.status='Please click here to delete this record';return true" onMouseOut="self.status='';return true" class="b">Delete</a></div>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="15%">Date</td>
                              <td class="blue" align="center" width="15%">Salary</td>
                              <td class="blue" align="center" width="5%">Currency</td>
                              <td class="blue" align="center" width="25%">Bank Account</td>
                              <td class="blue" align="center" width="40%">Bank detail</td>
                            </tr>
<%
		intRows = -1
		If IsArray(varSalary) Then
			intRows = Ubound(varSalary,2)
		End If	
		If intRows >= 0 Then
			For ii = 0 To intRows
				If Day(varSalary(1,ii)) < 10 Then
					sday	= "0" & Day(varSalary(1,ii))
				else
					sday	= Day(varSalary(1,ii))
				End If
				If Month(varSalary(1,ii)) < 10 Then
					smonth	= "0" & Month(varSalary(1,ii))
				else
					smonth	= Month(varSalary(1,ii))
				End If	 
				strDates	= sday & "/" & smonth & "/" & Year(varSalary(1,ii))
%>                            
                            <tr <%If ii Mod 2 = 0 Then%> bgcolor="#FFF2F2" <%Else%> bgcolor="#E7EBF5" <%End If%>> 
                              <td valign="top" class="blue">&nbsp;<a href="javascript:showdata('<%=ii%>');" onMouseOver="self.status='';return true" class="c"><%=strDates%></a></td>
                              <td valign="top" class="blue-normal" align="right"><%=decode(varSalary(3,ii),128)%>&nbsp;</td>
                              <td valign="top" class="blue-normal" align="center"><%=varSalary(5,ii)%></td>
                              <td valign="top" class="blue-normal" align="right"><%=varSalary(9,ii)%>&nbsp;</td>
                              <td valign="top" class="blue-normal" align="center"><%=varSalary(10,ii)%></td>
                            </tr>
<%
			Next
		End If	
%>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue-normal">&nbsp;&nbsp;* Click on salary date to update</td>
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
<%	Else
		If strError <> "" Then
%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red">&nbsp;<%=strError%></td>
				</tr>
<%		End If%>				

		  <tr>
         	<td class="red" align="center" valign="middle"><b><%=strError1%></b></td>
		  </tr>	          
<%	End If%>		  
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
<input type="hidden" name="txtstatus" value="<%=Request.Form("txtstatus")%>">
<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">

</form>

</body>
</html>
