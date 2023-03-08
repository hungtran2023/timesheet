<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	dim strProjectID,strSql,strStatus,strID
	dim rsCSODetails,rsCountry
	dim intCSOLevelID,intStaffLevel,dblContingency,dblHours
	Dim arrlstFrom(2),arrCategories,intCategoryType
	
	arrCategories =Array("0.Group Manager","1.Expat Manager","2.VN Manager","3.Senior Architect & Eng","4.Senior Technician","5.Architect & Engineer","6.Technician")	
'--------------------------------------------------
' Get Invoices
'--------------------------------------------------
function GetCSODetailList(rsCSODetail,IDDetail)
	dim strResult,strBkg,strDate
	dim idx,dblTotal,dblGrandTotal
	
	dblGrandTotal=0
	dblTotalHours=0
	idx=0

	if rsCSODetail.RecordCount>0 then
		strResult=""
		
'CSOLevelID,StaffLevel,Contingency,Hours

		Do while not rsCSODetail.EOF
			idx=idx+1
			dblTotal=0
			if cint(rsCSODetail("CSOLevelID"))= cint(IDDetail) then
				
				intCSOLevelID=rsCSODetail("CSOLevelID")
				intStaffLevel=rsCSODetail("StaffLevel")	
				dblContingency=rsCSODetail("Contingency")
				dblHours=rsCSODetail("Hours")
				
			end if
			strBkg="#E7EBF5"
			if (idx mod 2=1) then strBkg="#FFF2F2"
					
			strResult=strResult & "<tr bgcolor='" & strBkg & "'> "
            strResult=strResult & "<td valign='top' class='blue'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & idx & ".</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'><a href='javascript:showdata(" & rsCSODetail("CSOLevelID") & ")' class='c'><b>" & arrCategories(rsCSODetail("StaffLevel")) & "</b></a></td>"
			strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsCSODetail("Contingency"),2)  & "</td>"
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(rsCSODetail("Hours"),2) & "</td>"
            
            dblTotal=cdbl(rsCSODetail("Hours")) * cdbl(rsCSODetail("Contingency"))
            
            strResult=strResult & "<td valign='top' class='blue-normal' align='right'>" & formatnumber(dblTotal,2) & "</td>"
            strResult=strResult & "</tr>"
            dblGrandTotal=dblGrandTotal + dblTotal
            
			rsCSODetail.MoveNext
		loop

'		if dblTotal<>0 then
			strResult=strResult & "<tr bgcolor='#FFFFFF'><td colspan='4' align='right' valign='top' class='blue'>Total</td>" & _
									"<td valign='top' class='blue' align='right'>" & formatnumber(dblGrandTotal,2) & "</td>" & _
									"</tr>"
'		end if
	end if
	
	GetCSODetailList=strResult
	
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
	strProjectID=Left(Request.Form("txthidden"),15)
	strStatus=Request.Form("fgstatus")
	if strStatus="" then strStatus="A"
	
	intCSOLevelID=Request.Form("txtID")	
	if intCSOLevelID="" then intCSOLevelID=-1	

'--------------------------------------------------
' 
'--------------------------------------------------	
	strConnect = Application("g_strConnect")
	Set objDatabase = New clsDatabase

	If Request.QueryString("act") = "save" and Request.QueryString("choose_menu")="" Then
	
		If objDatabase.dbConnect(strConnect) Then		
			
			strSql=""
			
			intCSOLevelID=Request.form("txtID")
			intStaffLevel=Cint(Request.form("lstLevel"))
			dblContingency=Request.form("txtContingency")
			dblHours=Request.form("txtHours")
			
				
			select case strStatus
			
				'For add new
				case "A"
					strSql="INSERT INTO ATC_ProjectCSOByLevel (ProjectID,StaffLevel,Contingency,Hours,InsertedDate) " & _
								"VALUES ('" & strProjectID & "'," & intStaffLevel & "," & dblContingency & "," &  dblHours & ",Getdate())"
				'For edit
				case "E"
					strSql="UPDATE ATC_ProjectCSOByLevel " & _
								"SET StaffLevel = " & intStaffLevel & _
								",Contingency = " & dblContingency & _
								",Hours = " & dblHours & _
								",InsertedDate = getdate() " & _
							"WHERE CSOLevelID=" & intCSOLevelID
				'For delete
				case "D"
					strSql="DELETE FROM ATC_ProjectCSOByLevel WHERE CSOLevelID=" & intCSOLevelID
      
					intCSOLevelID=-1
					intStaffLevel=0
					dblContingency=""
					dblHours=""
					
			end select
			if strSql<>"" then

				If objDatabase.runActionQuery(strSQL) Then
					strError = "Update successful."
				Else
					strError = objDatabase.strMessage
				End If	

			end if
			
		end if
	elseif Request.QueryString("act") = "f" OR  Request.QueryString("act") = "g" then
	
		strSearchAPK=Request.Form("txtSearch")
				
		strSql="SELECT ProjectID,ProjectName FROM ATC_Projects WHERE " &_
			" ProjectID like '%" & strSearchAPK & "%' ORDER BY ProjectID"
		
		Call GetRecordset(strSql,rsAPKSearch)
		if gMessage="" then strAPKList=GetAPKList(rsAPKSearch)
		
		if Request.QueryString("act") = "g" then strProjectID=Request.Form("lstSearch")

	end if
	
	strSql="SELECT CSOLevelID,StaffLevel,Contingency,Hours FROM ATC_ProjectCSOByLevel WHERE ProjectID ='" & strProjectID & "' ORDER BY StaffLevel"
		
	Call GetRecordset(strSql,rsCSODetails)
	if gMessage="" then strLast=GetCSODetailList(rsCSODetails,cint(intCSOLevelID))
	
	arrlstFrom(0) = selectmonth("lstmonthF",intMonthDetail , -1)
	arrlstFrom(1) = selectyear("lstYearF", intYearDetail, 2000, year(now()) + 2, 0)
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
	
	document.frmreport.txtID.value = r;
	document.frmreport.action = "Pro_CSOByLevel.asp";
	document.frmreport.submit();
}
	
function adddata()
{
	window.document.frmreport.fgstatus.value = "A"
	//window.document.frmreport.txtMonthName.value = "";
	
	//selObj.selectedIndex = num;
	
	window.document.frmreport.lstLevel.selectedIndex = 0
	document.frmreport.txtID.value ="-1";
	window.document.frmreport.txtHours.value = "";
	window.document.frmreport.txtContingency.value="1"
	window.document.frmreport.lstLevel.focus();
}

	
function savedata()
{
	if (checkdata()==true)
	{
		window.document.frmreport.action = "Pro_CSOByLevel.asp?act=save"			
		window.document.frmreport.submit();
	}
}
	
function deletedata()
{
	var answer = confirm("Are you sure you want to remove current item?");
	if (answer){
		window.document.frmreport.fgstatus.value = "D";
		window.document.frmreport.action = "Pro_CSOByLevel.asp?act=save";
		window.document.frmreport.submit();
	}	
}

function back_menu()
{
	window.document.frmreport.action = "n_projectlist.asp?b=1";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function sub_menu()
{
	window.document.frmreport.action = "Pro_CSODetails.asp";
	window.document.frmreport.target = "_self";
	window.document.frmreport.submit();
}

function checkdata()
{
	var dblHours=document.frmreport.txtHours.value
	var dblContingency=document.frmreport.txtContingency.value
	
	if (dblHours==""){
		alert("CSO hours must be required.");
		document.frmreport.txtHours.focus();
		return false;
	}
	
	if (isNaN(dblHours) ==  true) 
	{
		alert("CSO hours must be number.");
		document.frmreport.txtHours.focus(); 
		return false;
	}
	
	if (dblContingency==""){
		alert("Contingency number must be required.");
		document.frmreport.txtContingency.focus();
		return false;
	}
	
	if (isNaN(dblContingency) ==  true){
		alert("Contingency number must be number.");
		document.frmreport.txtContingency.focus();
		return false;
	}

	return true;
}

function search()
{
	var tmp = document.frmreport.txtsearch.value;

	tmp = escape(tmp);
	if (alltrim(tmp) != "") {
		document.frmreport.action = "Pro_CSOByLevel.asp?act=f"
		document.frmreport.target = "_self";
		document.frmreport.submit();
	}
}

function go()
{
	document.frmreport.action = "Pro_CSOByLevel.asp?act=g"
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
						<A href="javascript:back_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">Project List</a> | 
						<A href="javascript:sub_menu();" onMouseOver="self.status='Return main menu';return true;" onMouseOut="self.status='';return true;">CSO Detail</a></td>
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
				    <td class="title" height="50" align="center" >CSO by Level</td>
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
                              <td valign="middle" class="blue-normal" width="20%">APK </td>
                              <td valign="middle" width="30%" class="blue"><%=strProjectID%></td>
                              <td valign="top" width="25%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top"  class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Staff Level</td>
                              <td valign="middle" class="blue-normal"> 
								<select name="lstLevel" style="width:80%" class="blue-normal">
                              <%	For ii=0 to Ubound(arrCategories)%>
									<option value="<%=ii%>" <%if ii=intStaffLevel then%>selected<%end if%>><%=arrCategories(ii)%></option>                              
							  <%	Next%>
								</select>
                              </td>
                              <td valign="top" class="blue-normal" align="center"></td>
                            </tr>
                           
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Hours</td>
                              <td valign="middle" class="blue-normal"> 
                                <input type="text" name="txtHours" class="blue-normal" size="20" style="width:80%" value="<%=dblHours%>">&nbsp;
                              </td>
                              <td valign="top" class="blue-normal" align="left"></td>
                            </tr>
                             <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Contingency</td>
                              <td valign="middle" class="blue-normal"> 
                                <input type="text" name="txtContingency" class="blue-normal" size="20" style="width:80%" value="<%=dblContingency%>">&nbsp;
                              </td>
                              <td valign="top" class="blue-normal" align="left"></td>
                             
                            </tr>
                          </table>
                          <input type="hidden" name="txtID" value="<%=intCSOLevelID%>">
                           
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

                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" align="center" width="10%">No.</td>
                              <td class="blue" align="center" width="40%">Staff Level</td>
                              <td class="blue" align="center" width="15%">Contingency</td>
                              <td class="blue" align="center" width="15%">Hours</td>
                              <td class="blue" align="center" width="20%">CSO Hours</td>
                            </tr>
<%=strLast%>
                          </table>
<%if strLast<>"" then%>                          
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="20" class="blue-normal">&nbsp;&nbsp;* Click on Staff Level to update</td>
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
	Response.Write(arrTmp(1))%>
<%
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
	Response.Write(arrPageTemplate(2))%>
	
<input type="hidden" name="txthidden" value="<%=strProjectID%>">
<input type="hidden" name="fgstatus" value="<%=strStatus%>">

<input type="hidden" name="P" value="<%=Request.Form("P")%>">
<input type="hidden" name="S" value="<%=Request.Form("S")%>">


</form>

</body>
</html>