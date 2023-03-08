<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	dim strProjectID,strSql,strStatus,strID
	dim rsInvoices,rsCountry
	dim strInvoiceNumber,strInvoiceDate,strValue,strCurrency,strExRate, dblBDM, dblProManagerID
	dim dblTPValue,strTPCurrency,dblTPValueExRate

'--------------------------------------------------
' Get Invoices
'--------------------------------------------------
function GetInvoiceList(rsInvoice,selectIdx)
	dim strResult,strBkg,strDate
	dim idx,dblTotal,dblGrandTotal, dblGrandTotalTP,dblTotalTP,dblOrgGrandTotal,dblOrgGrandTotalTP
	dblGrandTotal=0
	dblGrandTotalTP=0
	dblOrgGrandTotal=0
	dblOrgGrandTotalTP=0
	
	idx=0
	if rsInvoice.RecordCount>0 then
		strResult=""
		Do while not rsInvoice.EOF
			idx=idx+1
			dblTotal=0
			if selectIdx=idx then
				strInvoiceDate=rsInvoice("InvoiceDate")
				strInvoiceNumber=rsInvoice("InvoiceNumber")
				strValue=rsInvoice("InvoiceValue")
				strID=rsInvoice("InvoiceID")	
				strExRate=rsInvoice("ExchangeRate")
				dblBDM=rsInvoice("BDMID")
				dblProManagerID=rsInvoice("ProjectManagerID")
				
				dblTPValue=rsInvoice("ThirdPartyValue")
				strTPCurrency=rsInvoice("ThirdPartyCurrency")
				dblTPValueExRate=rsInvoice("ThirdPartyExRate")
			end if
			strBkg="#E7EBF5"
			if (idx mod 2=1) then strBkg="#FFF2F2"
			strDate=day(rsInvoice("InvoiceDate")) & "/" & month(rsInvoice("InvoiceDate")) & "/" & year(rsInvoice("InvoiceDate"))
			strResult=strResult & "<tr bgcolor='" & strBkg & "'> "
            strResult=strResult & "<td valign='top' class='blue'>" & idx & ".</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'><a href='javascript:showdata(" & idx & ")' class='c'><b>" & rsInvoice("InvoiceNumber") & "</b></a></td>"
			strResult=strResult & "<td valign='top' class='blue-normal' align='center'>" & strDate & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsInvoice("InvoiceValue"),2) & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsInvoice("ExchangeRate"),2) & "</td>"
            dblTotal=cdbl(rsInvoice("InvoiceValue")) * cdbl(rsInvoice("ExchangeRate"))
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(dblTotal,2) & "</td>"
            
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsInvoice("ThirdPartyValue"),2) & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsInvoice("ThirdPartyExRate"),2) & "</td>"
            dblTotalTP=cdbl(rsInvoice("ThirdPartyValue")) * cdbl(rsInvoice("ThirdPartyExRate"))
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(dblTotalTP,2) & "</td>"
            
            
            
            strResult=strResult & "</tr>"
            dblGrandTotal=dblGrandTotal + dblTotal
            dblGrandTotalTP=dblGrandTotalTP + dblTotalTP
            dblOrgGrandTotal=dblOrgGrandTotal + cdbl(rsInvoice("InvoiceValue"))
            dblOrgGrandTotalTP=dblOrgGrandTotalTP+ cdbl(rsInvoice("ThirdPartyValue"))
			rsInvoice.MoveNext
		loop
		
		if dblGrandTotal<>0 OR dblGrandTotalTP<>0 then
			strResult=strResult & "<tr bgcolor='#FFFFFF'>" & _
								"<td colspan='3' align='right' valign='top' class='blue'>Total</td>" & _								
								"<td valign='top' class='blue' align='right'>" & formatnumber(dblOrgGrandTotal,2) & "</td>" & _
								"<td colspan='2' valign='top' class='blue' align='right'>" & formatnumber(dblGrandTotal,2) & "</td>" & _
								"<td valign='top' class='blue' align='right'>" & formatnumber(dblGrandTotalTP,2) & "</td>" & _
								"<td colspan='2' valign='top' class='blue' align='right'>" & formatnumber(dblOrgGrandTotalTP,2) & "</td></tr>"
		end if
		
	end if
	GetInvoiceList=strResult
end function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
'--------------------------------------------------
' User can update all project invoice
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgInvoice = False
	Else
		varGetRight = session("RightOn")
		fgInvoice = False
		For ii = 0 To Ubound(varGetRight, 2)
'Response.Write 	varGetRight(0, ii)		 & "<br>"
			If varGetRight(0, ii) = "Invoice" Then

				fgInvoice = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If		
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	strProjectID=Request.Form("txthidden")
	strStatus=Request.Form("fgstatus")
	selectRow=Request.QueryString("r")
	if selectRow="" then selectRow=-1	
	
	strID=-1
'--------------------------------------------------
' Get currency
'--------------------------------------------------
	Call GetRecordset("SELECT CurrencyCode FROM ATC_Projects WHERE ProjectID='" & strProjectID & "'",rsCurrency)	

	strSql="SELECT * FROM ATC_Currency WHERE fgActivate=1"
	Call GetRecordset(strSql,rsCurrencyExist)
	
'--------------------------------------------------
' Initialize BDM recordset
'--------------------------------------------------	
		strSql="SELECT * FROM HR_BDM ORDER BY Firstname"	
		Call GetRecordset(strSql,rsBDM)

'--------------------------------------------------
' Initialize Project Manager recordset
'--------------------------------------------------	
	strSql = "SELECT DISTINCT a.UserID, e.Firstname + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') as Fullname " &_
			"FROM ATC_UserGroup a LEFT JOIN ATC_Group b ON a.GroupID = b.GroupID " &_
			"LEFT JOIN ATC_Permissions c ON b.GroupID = c.GroupID " &_
			"LEFT JOIN ATC_Functions d ON c.FunctionID = d.FunctionID " &_
			"LEFT JOIN ATC_PersonalInfo e ON a.UserID = e.PersonID " &_
			"WHERE d.Description = 'Manager' AND e.FirstName <> 'Managers' AND e.fgDelete = 0 ORDER BY Fullname"	
	Call GetRecordset(strSql,rsManagers)
		
		
'--------------------------------------------------
' 
'--------------------------------------------------	
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase

	If Request.QueryString("act") = "save" and Request.QueryString("choose_menu")="" Then
		If objDatabase.dbConnect(strConnect) Then		
			if strStatus="" then strStatus="A"
			strInvoiceNumber=Request.Form("txtInvoiceNo")
				
			dblBDM=Request.Form("lbBDM")
			dblProManagerID=Request.Form("lbProManager")
			
			
			if Request.Form("txtdate")<>"" then
				varDate = split(Request.Form("txtdate"),"/")
				If Not IsEmpty(varDate) Then strInvoiceDate = CDate(varDate(1) & "/" & varDate(0) & "/" & varDate(2))
			End If
			strSql=""
			strValue=Request.Form("txtValue")
			if strValue="" then strValue=0
			strExRate=Request.Form("txtexRate")
			if strExRate="" then strExRate=0
			
			dblTPValue=Request.Form("txtThirdPartyValue")
			if dblTPValue="" then dblTPValue=0
			strTPCurrency=Request.Form("lbCurrency")
			dblTPValueExRate=Request.Form("txtThirdPartyExRate")
			if dblTPValueExRate="" then dblTPValueExRate=1
			
			if strExRate="" then strExRate=1
			select case strStatus
				'For add new
				case "A"
					strSql="INSERT INTO ATC_ProjectInvoices (InvoiceNumber,ProjectID,BDMID,ProjectManagerID,InvoiceDate,InvoiceValue,ExchangeRate,OwnerID,ThirdPartyValue,ThirdPartyCurrency,ThirdPartyExRate) " &_
							"VALUES('" & strInvoiceNumber & "','" & strProjectID & "'," & IIF(dblBDM="", NULL, dblBDM ) & "," & IIF(dblProManagerID="", NULL, dblProManagerID ) & ",'" & strInvoiceDate & "'," & strValue & "," & strExRate & "," & intUserID &_
							"," & dblTPValue & ",'" & IIF(strTPCurrency="",NULL,strTPCurrency) & "'," & dblTPValueExRate & ")"
				'For edit
				case "E"
					strID=Request.Form("txtCurrentID")
					strSql="UPDATE ATC_ProjectInvoices SET InvoiceNumber='" & strInvoiceNumber & "'," & _
					                  " BDMID='" &  IIF(dblBDM="", NULL, dblBDM ) & "'," & _
									  " ProjectManagerID='" &  IIF(dblProManagerID="", NULL, dblProManagerID ) & "'," & _
									"InvoiceDate='" & strInvoiceDate & "',InvoiceValue=" & strValue & ",ExchangeRate=" & strExRate &_
									",OwnerID=" & intUserID & " ,ThirdPartyValue=" & dblTPValue & ",ThirdPartyCurrency='" & strTPCurrency & "',ThirdPartyExRate=" & dblTPValueExRate & _
									" WHERE InvoiceID=" & strID
				'For delete
				case "D"
					strID=Request.Form("txtCurrentID")
					strSql="DELETE ATC_ProjectInvoices WHERE InvoiceID=" & strID
					selectRow=-1
					strInvoiceNumber=""
					strInvoiceDate=""
					strValue=""
					strExRate=""
					 
					dblBDM=-1
					dblTPValue=""
					strTPCurrency=""
					dblTPValueExRate=""
					
					strID=-1
					strStatus=""
			end select
			if strSql<>"" then

				If objDatabase.runActionQuery(strSQL) Then
					strError = "Update successful."
				Else
					strError = objDatabase.strMessage
				End If	
			end if
			
		end if
	end if
	
	strSql="SELECT  InvoiceID, InvoiceNumber, ProjectID, InvoiceDate, InvoiceValue, ExchangeRate,OwnerID, DateCreated, BDMID, ProjectManagerID," & _
			"ISNULL(ThirdPartyValue,0) as ThirdPartyValue,ThirdPartyCurrency,ISNULL(ThirdPartyExRate,0) as ThirdPartyExRate " & _
				"FROM ATC_ProjectInvoices WHERE ProjectID='" & strProjectID & "' ORDER BY InvoiceDate"
	
	Call GetRecordset(strSql,rsInvoices)
	if gMessage="" then strLast=GetInvoiceList(rsInvoices,cint(selectRow))

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
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

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
	
function showdata(r)
{
	document.frmreport.fgstatus.value = "E";
	document.frmreport.action = "Pro_Invoice.asp?r=" + r;
	document.frmreport.submit();
}
	
function adddata()
{
	window.document.frmreport.fgstatus.value = "A"
	window.document.frmreport.txtdate.value = "";
	window.document.frmreport.txtInvoiceNo.value = "";
	window.document.frmreport.txtValue.value = "";
	window.document.frmreport.txtExRate.value = "";
		
	window.document.frmreport.txtThirdPartyValue.value = "";
	window.document.frmreport.txtThirdPartyExRate.value = "";
	document.getElementById("lbBDM").value = "";
	document.getElementById("lbProManager").value = "";
	
	window.document.frmreport.txtInvoiceNo.focus();
}
	
function savedata()
{
	if (checkdata()==true)
	{
		window.document.frmreport.action = "Pro_Invoice.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "Pro_Invoice.asp?act=save"			
	window.document.frmreport.submit();
}

function back_menu()
{
	window.document.frmreport.action = "n_projectlist.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}
	
function checkdata()
{
	var dateInvoice=document.frmreport.txtdate.value
	var strInvNu=document.frmreport.txtInvoiceNo.value
	var strValue=document.frmreport.txtValue.value
	
	if (strInvNu==""){
		alert("The Invoice Number must be required.");
		document.frmreport.txtInvoiceNo.focus();
		return false;
	}
	if (dateInvoice==""){
		alert("The Invoice date must be required.");
		document.frmreport.txtdate.focus();
		return false;
	}
	if(isdate(dateInvoice)==false) {
		alert("The invoice date (" + dateInvoice + ") is invalid.");
		document.frmreport.txtdate.focus();
		return false;
	}
	
	if (strValue==""){
		alert("The Invoice date must be required.");
		document.frmreport.txtValue.focus();
		return false;
	}
	if (isNaN(strValue) ==  true) 
	{
		alert("Invalid invoice value.");
		document.frmreport.txtValue.focus(); 
		return false;
	}
	return true;
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
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;
					<A href="javascript:back_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">Project List</a>
                  </td>
                  <td class="blue" height="30" align="right" width="77%">&nbsp;</td>
                </tr>
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Project Invoices</td>
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
                              <td valign="middle" class="blue-normal" width="20%">ProjectID</td>
                              <td valign="middle" width="30%" class="blue"><%=strProjectID%></td>
                              <td valign="top" width="25%" class="blue-normal" align="center">&nbsp;</td>
                              

                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top"  class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Invoice No.</td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtInvoiceNo" class="blue-normal" size="20" style="width:100%" value="<%=strInvoiceNumber%>">
                              </td>
                              <td valign="top" class="blue-normal" align="center"></td>

                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >BDM</td>
                              <td valign="middle" class="blue-normal"> 
							  <select class='blue-normal' name='lbBDM' id='lbBDM' style='WIDTH: 228px; HEIGHT: 20px'>
									<option value="" <%if dblBDM="" then%>selected<%end if%>>&nbsp;</option>
									<%do while not rsBDM.EOF%>							
										<option value="<%=rsBDM("BDMID")%>" <%if dblBDM=rsBDM("BDMID") then%>selected<%end if%>><%=rsBDM("Fullname")%></option>
									<% rsBDM.Movenext
									loop%>
							</select>
                                
                              </td>
                              <td valign="top" class="blue-normal" align="left"></td>
                             
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Project Manager</td>
                              <td valign="middle" class="blue-normal"> 
							  <select class='blue-normal' name='lbProManager' id='lbProManager' style='WIDTH: 228px; HEIGHT: 20px'>
									<option value="" <%if dblProManagerID="" then%>selected<%end if%>>&nbsp;</option>
									<%do while not rsManagers.EOF%>							
										<option value="<%=rsManagers("UserID")%>" <%if dblProManagerID=rsManagers("UserID") then%>selected<%end if%>><%=rsManagers("Fullname")%></option>
									<% rsManagers.Movenext
									loop%>
							</select>
                                
                              </td>
                              <td valign="top" class="blue-normal" align="left"></td>
                             
                            </tr>							
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Date</td>
                              <td valign="middle" class="blue-normal"> 
                                <input type="text" name="txtdate" class="blue-normal" size="20" style="width:60%" value="<%IF (strInvoiceDate<>"") then Response.Write day(strInvoiceDate) & "/" & month(strInvoiceDate) & "/" & year(strInvoiceDate) end if%>">&nbsp;(dd/mm/yyyy)
                              </td>
                              <td valign="top" class="blue-normal" align="left"></td>
                             
                            </tr>
                            
                                                        
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Original Value </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtValue" class="blue-normal" size="20" style="width:80%" value="<%=strValue%>"> (<%=rsCurrency("CurrencyCode")%>)</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                                                  
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Ex. Rate </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtExRate" class="blue-normal" size="20" style="width:80%" value="<%=strExRate%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                            
                                                        
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="top" class="blue"><u>Outsourcing cost</u> </td>
                              <td valign="middle" class="blue-normal"> </td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                             
                            
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">TP/JV Value </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtThirdPartyValue" class="blue-normal" style="width:70%" value="<%=dblTPValue%>">
                              <select class='blue-normal' name='lbCurrency' style='WIDTH: 28%'>
                              

	                              
								<option value="" <%if strCurrency="" then%>selected<%end if%>>&nbsp;</option>
							<%do while not rsCurrencyExist.EOF%>
							
								<option value="<%=rsCurrencyExist("CurrencyCode")%>" <%if strTPCurrency=rsCurrencyExist("CurrencyCode") then%>selected<%end if%>><%=rsCurrencyExist("CurrencyCode")%></option>
								
							<% rsCurrencyExist.Movenext
							loop%>
                              
							</select>
                              </td>
                              <td valign="top" class="blue-normal" align="left">
                           
                                                            
                              &nbsp;</td>
                            </tr>
                            
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Ex. Rate </td>
                              <td valign="middle" class="blue-normal"> 
                              <input type="text" name="txtThirdPartyExRate" class="blue-normal" size="20" style="width:70%" value="<%=dblTPValueExRate%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>                             
                          </table>
                          

                          <input type="hidden" name="txtCurrentID" value="<%=strID%>">
<%'For Tu Tram or Uyen Chi
if cdbl(intUserID)=251 OR cdbl(intUserID)=252 or cdbl(intUserID)=527 OR fgInvoice then%>                            
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="180" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60"> 
                                      <div align="center" class="blue"><a href="javascript:adddata()" onMouseOver="self.status='Please click here to add new record';return true" onMouseOut="self.status='';return true" class="b">Add</a></div>
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
<%end if%>                          
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                           <tr bgcolor="#8CA0D1">
                              <td width="5%" rowspan="2" align="center" bgcolor="#8CA0D1" class="blue">No</td>
                              <td width="9%" rowspan="2" align="center" class="blue">Invoice Number</td>
                              <td width="10%" rowspan="2" align="center" class="blue">Date</td>
                              
                              <td colspan="3" align="center" class="blue">Invoice Information </td>
                              <td colspan="3" align="center" class="blue">Outsourcing cost</td>
                            </tr>
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" align="center" width="13%">Value</td>
                              <td class="blue" align="center" width="10%">Ex. Rate</td>
                              <td class="blue" align="center" width="15%">Total<br>(USD)</td>

                             <td class="blue" align="center" width="13">Value</td>
                              <td class="blue" align="center" width="10%">Ex. Rate</td>
                              <td class="blue" align="center" width="15%">Total<br>(USD)</td>

                            </tr>

<%=strLast%>
                          </table>
<%if strLast<>"" then%>                          
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue-normal">&nbsp;&nbsp;* Click on Invoce No. to update</td>
                            </tr>
                          </table>
<%end if%>                          
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
<input type="hidden" name="txthidden" value="<%=Request.Form("txthidden")%>">
<input type="hidden" name="fgstatus" value="<%=strStatus%>">

<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">

</form>

</body>
</html>