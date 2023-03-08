<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	dim strProjectID,strSql,strStatus,strID

'--------------------------------------------------
' Get Invoices
'--------------------------------------------------
function GetDetailList(rsDetail)
	dim strResult,strBkg,strDate
	dim idx,dblTotal,dblCSOTotalHours, dblCSOTotalPayment
	dblCSOTotalHours=0
	dblCSOTotalPayment=0
	dblTotalPay=0
	dblTotalHours=0
	idx=0


	if rsDetail.RecordCount>0 then
		strResult=""

		Do while not rsDetail.EOF
			idx=idx+1
			dblTotal=0

			strBkg="#E7EBF5"
			if (idx mod 2=1) then strBkg="#FFF2F2"
			
	        strMonthNameDisplay=monthname(rsDetail("periodMonth"),2) & "-" & rsDetail("periodYear")
	        
			strResult=strResult & "<tr bgcolor='" & strBkg & "'> "
            strResult=strResult & "<td valign='top' class='blue'>" & idx & ".</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & strMonthNameDisplay & "</td>"
			strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsDetail("CSOHours"),2)  & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsDetail("Hours"),2) & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsDetail("CSOPayment"),2) & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsDetail("InvoiceValue"),2) & "</td>"
           
            strResult=strResult & "</tr>"
            
            dblTotalHours=dblTotalHours + cdbl(rsDetail("Hours"))
            dblTotalPay=dblTotalPay+ cdbl(rsDetail("InvoiceValue"))
            dblCSOTotalHours=dblCSOTotalHours + cdbl(rsDetail("CSOHours"))
            dblCSOTotalPayment=dblCSOTotalPayment + cdbl(rsDetail("CSOPayment"))
            
			rsDetail.MoveNext
		loop
	    if (dblTotalHours+dblCSOTotalHours+dblTotalPay+dblCSOTotalPayment)<>0 then
			strResult=strResult & "<tr bgcolor='#FFFFFF'><td colspan='2' align='right' valign='top' class='blue'>Total</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblCSOTotalHours,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblTotalHours,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblCSOTotalPayment,2) & "</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblTotalPay,2) & "</td>" & _
									"</tr>"
        end if
	end if

	GetDetailList=strResult
end function
'--------------------------------------------------
' 
'--------------------------------------------------
function GetAPKList(rsAPKSearch)
	
	dim strReturn
	strReturn=""
	if rsAPKSearch.recordCount>0 then	
		
		do while not rsAPKSearch.EOF
			
			strReturn=strReturn & "<option value='" & rsAPKSearch("ProjectID") & "'>" & rsAPKSearch("ProjectID") & "-" & rsAPKSearch("ProjectName") &  "</option>" & vbCrLf 			
			rsAPKSearch.MoveNext
		loop	
		
	end if
	
	GetAPKList=strReturn
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
	strProjectID=Left(Request.Form("txthidden"),15)

	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase
		
	strSearchAPK=Request.Form("txtSearch")
		
	if strSearchAPK<>"" then
				
		strSql="SELECT ProjectID,ProjectName FROM ATC_Projects WHERE " &_
			" ProjectID like '%" & strSearchAPK & "%' ORDER BY ProjectID"
		
		Call GetRecordset(strSql,rsAPKSearch)
		if gMessage="" then strAPKList=GetAPKList(rsAPKSearch)
			
	end if
	
	if Request.QueryString("act") = "g" then strProjectID=Request.Form("lstSearch")
		
	strSql=" SELECT Period,periodMonth, periodYear, (Hours+OTHours) as Hours, InvoiceValue, CSOHours, CSOPayment FROM rp_ProjectPerformanceByPeriod WHERE  ProjectID='" & strProjectID & "' ORDER BY Period "
			
		
	Call GetRecordset(strSql,rsDetail)
	if gMessage="" then strLast=GetDetailList(rsDetail)
	
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
	
	


function back_menu()
{
	window.document.frmreport.action = "listofprojectperformance.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

	

function search()
{
	document.frmreport.action = "Pro_PerformmanceDetails.asp?act=f"
	document.frmreport.target = "_self";
	document.frmreport.submit();
}

function go()
{
	document.frmreport.action = "Pro_PerformmanceDetails.asp?act=g"
	document.frmreport.target = "_self";
	document.frmreport.submit();
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
				  <td class="red">&nbsp;<b><%=strError%></b></td>
			</tr>
<%		End If%>				
			
				<tr align="center"> 
					<td class="blue" height="30" align="left" width="23%"> &nbsp;&nbsp; &nbsp;&nbsp; 
						<A href="javascript:back_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">Project Performance</a>
			     </tr>
				<tr> 
					<td align="center" valign="middle">
						<table width="98%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="15%" class="blue-normal" valign="middle" align="right"> Search for APK &nbsp; </td>
								<td width="20%" ><input type="text" name="txtsearch" class="blue-normal" size="15" style="width:98%" value="<%=strSearchAPK%>"></td>
								<td width="45%">
									<select name="lstSearch" style="width:98%" class="blue-normal" onChange="javascript:go()">
										<option value="-1"></option>
										<%=strAPKList%>
									</select></td>
								<td width="20%">
									<table width="100%" border="0" cellspacing="3" cellpadding="0" height="20" align="left">
										<tr> 
											<td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
												<a href="javascript:search();" class="b" onMouseOver="self.status='Search'; return true;" onMouseOut="self.status=''">Search</a></td>
										</tr>
									</table>
								</td>
							</tr>
						</table> 
					</td>
				</tr>                
			    <tr align="center"> 
				    <td class="title" align="center" ><p>Project Performance Details <p><span class="blue" style="font-size: 10pt">
				         <%=strProjectID%></span><p> </td>
			    </tr>
			</table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td bgcolor="#617DC0"> 
            
                            <table width="100%" border="0" cellspacing="1" cellpadding="5">
                                <tr bgcolor="#8CA0D1"> 
                                  <td class="blue" align="center" width="5%">No.</td>
                                  <td class="blue" align="center" width="15%">Month</td>
                                  <td class="blue" align="center" width="20%">CSO Hours</td>
                                  <td class="blue" align="center" width="20%">Actual Hours</td>
                                  <td class="blue" align="center" width="20%">Payment Schedule</td>
                                  <td class="blue" align="center" width="20%">Invoices</td>
                                </tr>
                                    <%=strLast%>
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
	Response.Write(arrTmp(1))%>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))%>
	
<input type="hidden" name="txthidden" value="<%=strProjectID%>">

</form>

</body>
</html>