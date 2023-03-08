<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDatabase, strError, intPageSize, fgRight 'view all or Not

'****************************************
' Function: Outbody
' Description: 
' Parameters: source recordset, number of lines on one page
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************

Function Outbody(ByRef rsSrc, ByVal psize)

	dim dblExRate,strStatus,strNextStep
	'Method 3 : using 'Array' Parameter
	Dim arrStatus, arrNext
	arrStatus = Array("New","In-progress","Rejected","Done")
	strOut = ""
	
	If Not rsSrc.EOF Then
		
		rsSrc.movefirst
		i=0
		Do while not rsSrc.EOF
			
			If i Mod 2 = 0 Then
				strColor = "#E7EBF5"
			Else
				strColor = "#FFF2F2"
			End If
			
			strStatus=arrStatus(rsSrc("Status"))
			
			if Cint(rsSrc("Status"))=2 then
				if  not rsSrc("isAuthoriser1Approved") then
					strStatus = strStatus & " by " & rsSrc("Authoriser1")
				elseif rsSrc("Authoriser2_Id")<>0  AND not rsSrc("isAuthoriser2Approved") then
					strStatus = strStatus & " by " & rsSrc("Authoriser2")
				elseif rsSrc("isAuthorisedByHr")  AND not rsSrc("isHrApproved") then
					strStatus = strStatus & " by HR"
				end if
				strNextStep="Done"
			else
				if not rsSrc("isAuthoriser1Approved") then
					 strNextStep=rsSrc("Authoriser1")  
				elseif rsSrc("Authoriser2_Id")<>0  AND not rsSrc("isAuthoriser2Approved") then
					 strNextStep=rsSrc("Authoriser2")
				elseif rsSrc("isAuthorisedByHr")  AND not rsSrc("isHrApproved") then
					strNextStep="HR Authorise"
				else
					 strNextStep="Done"
				end if
			end if
			 			 
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td class='blue-normal'>" & rsSrc("ID") & "</td>" &_
			         "<td class='blue'><a class='c' href='javascript:viewRequest(" & rsSrc("ID") & ")' >" & rsSrc("Requester") & "</a></td>" &_
			         "<td class='blue-normal'>" & day(rsSrc("DateFrom")) & "/" & month(rsSrc("DateFrom")) & "/" & year(rsSrc("DateFrom")) & " " & timevalue(rsSrc("DateFrom")) &"</td>" &_
					 "<td class='blue-normal'>"& day(rsSrc("DateTo")) & "/" & month(rsSrc("DateTo")) & "/" & year(rsSrc("DateTo")) & " " & timevalue(rsSrc("DateTo")) &"</td>" &_			         
			         "<td class='blue-normal'>"& strStatus & "</td>" &_
			         "<td class='blue-normal'>" & strNextStep & "</td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			i=i+1
		loop

	End If
	Outbody = strOut
End Function

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	strDepartment = Request.Form("lbdepartment")
	fgSort = Request.Form("S")
	
	intCurPage = trim(Request.Form("P"))
	If intCurPage = "" Then
		intCurPage = 1
	End If		
	strName = Request.Form("name")
	intDepart = Request.Form("depart")
	
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
	
'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = PageSizeDefault
		Set arrPre = Nothing
	Else
		intPageSize = PageSizeDefault
	End If

'--------------------------------------------------
' Check ACCESS right
'--------------------------------------------------

	strTemp = Request.ServerVariables("URL") 
	While Instr(strTemp, "/")<>0
		strTemp = Mid(strTemp, Instr(strTemp, "/") + 1, Len(strTemp))
	Wend
	
	strFilename = strTemp
	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = strTemp Then
				fgRight=True
				Exit For
			End If
		Next
		
		Set varGetRight = Nothing		
	End If	
	If fgRight = False Then		
		Response.Redirect("../../welcome.asp")
	End If

'--------------------------------------------------
' Check VIEWALL right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "view all" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Initialize department array
'--------------------------------------------------
	
	strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

	If isEmpty(session("varDepartment")) = False Then
		varDepartment = session("varDepartment")
		intNum = Ubound(varDepartment,2)
	Else
		If objDatabase.dbConnect(strConnect) Then			
			strSQL = "SELECT * FROM ATC_Department WHERE fgActivate=1 ORDER BY Department"

			If (objDatabase.runQuery(strSQL)) Then
				If objDatabase.noRecord = False Then
					varDepartment = objDatabase.rsElement.GetRows
					intNum = Ubound(varDepartment,2)					
					session("varDepartment") = varDepartment
					objDatabase.closeRec
				End If
			Else
				Response.Write objDatabase.strMessage
			End If
		Else
			Response.Write objDatabase.strMessage		
		End If
	End If	

'--------------------------------------------------
' End Of initializing department array
'--------------------------------------------------

'--------------------------------------------------
' Analyse query and prepare staff list
'--------------------------------------------------

	strAct = Request.QueryString("act")
	If strAct = "" Then
		strAct = Request.Form("txtstatus")
	End If

	
	strSQL="SELECT b.FirstName + ' ' + b.LastName AS Requester, c.FirstName + ' ' + c.LastName AS Authoriser1, d.FirstName + ' ' + d.LastName AS Authoriser2," & _
				" a.Id, a.Type, a.StaffId, a.Authoriser1_Id, ISNULL(a.Authoriser2_Id, 0) AS Authoriser2_Id, a.DateFrom, a.DateTo, a.Note, a.isAuthorisedByHr, " & _
				" a.isAuthoriser1Approved, ISNULL(a.isAuthoriser2Approved, 0) AS isAuthoriser2Approved, ISNULL(a.isHrApproved, 0) AS isHrApproved, " & _ 
				"a.Status, a.Authoriser1Note, a.Authoriser2Note, a.HrNote " & _
				" FROM ATC_AbsenceRequests AS a INNER JOIN " & _
                         "ATC_PersonalInfo AS b ON a.StaffId = b.PersonID INNER JOIN " & _
                         "ATC_PersonalInfo AS c ON a.Authoriser1_Id = c.PersonID LEFT OUTER JOIN " & _
                         "ATC_PersonalInfo AS d ON a.Authoriser2_Id = d.PersonID" 

	' if trim(strSearch<>"") then
		' intSearchType=Request.Form("lstType")
		' if CInt(intSearchType)=1 then 
			' strSearch=" SoftwareName like '%" & trim(strSearch) & "%'"
		' elseif CInt(intSearchType)=2 then
			' strSearch="  b.Description like '%" & trim(strSearch) & "%'"
		' end if
		' strSQL=strSQL & " WHERE " & strSearch 
	' end if
	
	strSQL=strSQL & " ORDER BY a.id DESC"

'	response.write strSQl & "<br>"
	Call GetRecordset(strSQL,rsData)
	
	strLast=Outbody(rsData,intPageSize)
	
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
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

function sort(kind)
{
	document.frmreport.action = "annual_list_staff.asp?act=vps" + kind;
	document.frmreport.submit();
}

function viewpage(kind)
{
	var intpage = parseInt(window.document.frmreport.txtpage.value,10);
	var curpage = "<%=CInt(intCurPage)%>";
	var pagetotal = "<%=CInt(intTotalPage)%>";
	
	if (kind == 1)
	{
		window.document.frmreport.txtpage.value = intpage
		if ((intpage > 0) & (intpage <= pagetotal) & (intpage != curpage)) 
		{
			document.frmreport.action = "annual_list_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "annual_list_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function search(kind)
{
	if (kind == "1")
	{
		if ((document.frmreport.txtname.value != "") || (document.frmreport.lbdepartment.options[document.frmreport.lbdepartment.selectedIndex].value != "0"))
		{
			document.frmreport.action = "annual_list_staff.asp?act=vra" + kind;
			document.frmreport.submit();	
		}
	}	
	else
	{
		document.frmreport.action = "annual_list_staff.asp?act=vra" + kind;
		document.frmreport.submit();
	}	
}

function viewAnnual(varid)
{
	window.document.frmreport.txthidden.value = varid;
	window.document.frmreport.action = "annual_detail_staff.asp";
	window.document.frmreport.submit();

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
<%	If strError1 = "" Then%>        
		  <tr> 
		    <td> 
		      <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%If strError <> "" Then%>		      
		        <tr bgcolor="#E7EBF5">
				  <td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
		        </tr>
 <%End If%>		        
		        <!--<tr align="center"> 
		          <td class="blue-normal" height="30" align="right" width="10%">&nbsp;&nbsp;Name&nbsp;</td>
  				  <td class="blue" height="30" align="left" width="25%"> 
				    <input type="text" name="txtname" value="<%=showvalue(strSName)%>" class="blue-normal" size="15" style=" width:150">
				  </td>
				  <td class="blue-normal" height="30" align="right" width="15%">Department&nbsp;</td>
				  <td class="blue" height="30" align="left" width="25%"> 
				    <select id="lbdepartment" size="1" name="lbdepartment" style="width:100%" class="blue-normal">
					  <option value="" selected>&nbsp;</option>
<%
		If intNum >= 0 Then
			For ii = 0 To intNum
%>                    
					  <option <%If strDepartment=varDepartment(1,ii)Then%> selected <%End If%> value="<%=showlabel(trim(varDepartment(1,ii)))%>"><%=showlabel(varDepartment(1,ii))%></option>
<%
			Next
		End If	
%>					

				    </select>
				  </td>
				  <td class="blue-normal" height="30" align="left" width="25%"> 
				    <table width="120" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
					  <tr> 
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20"> 
						  <div align="center" class="blue"><a href="javascript:search('1');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Search</a></div>
					    </td>
					    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" class="blue" align="center">
						  <a href="javascript:search('2');" class="b" onMouseOver="self.status='';return true" onMouseOut="self.status='';return true">Show All</a>
					    </td>
					  </tr>
				    </table>
				  </td>
			    </tr>-->
			    <tr align="center"> 
				  <td class="title" height="50" align="center" colspan="5">List of booking request</td>
			    </tr>
			  </table>
		    </td>
		  </tr>
		  <tr> 
		    <td height="100%" valign="top"> 
		      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:'79%'" height="365">
		        <tr> 
		          <td bgcolor="#FFFFFF" valign="top"> 
		            <table width="100%" border="0" cellspacing="0" cellpadding="0">
		              <tr> 
		                <td bgcolor="#617DC0"> 
		                  <table width="100%" border="0" cellspacing="1" cellpadding="5">
		                    <tr bgcolor="#8CA0D1"> 
		                      <td class="blue" align="center" width="4%">ID</td>
		                      <td class="blue" align="center" width="30%">Requeter</td>
		                      <td class="blue" align="center" width="15%">From</td>
		                      <td class="blue" align="center" width="15%">To</td>
		                      <td class="blue" align="center" width="20%">Status</td>
		                      <td class="blue" align="center" width="16%">Next step</td>
		                    </tr>
<%
	Response.Write(strLast)
%>                            
		                  </table>
		                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		                    <tr> 
		                      <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;&nbsp;* Click 
		                        on the exact name to view or update Annual Leave Information.</td>
		                      <td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">&nbsp;</td>
		                    </tr>
		                    <tr> 
		                      <td bgcolor="#FFFFFF" class="blue-normal" colspan="2">&nbsp;</td>
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
          <tr> 
            <td> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
                <tr> 
                  <td align="right" bgcolor="#E7EBF5"> 
                    <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
                      <tr class="black-normal"> 
                        <td align="right" valign="middle" width="27%" class="blue-normal">Page 
                        </td>
                        <td align="center" valign="middle" width="13%" class="blue-normal"> 
                          <input type="text" name="txtpage" class="blue-normal" value="<%=intCurPage%>" size="2" style="width:50">
                        </td>
                        <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:viewpage(1);" onMouseOver="self.status='';return true"><font color="#990000">Go</font></a></td>
						<td align="right" valign="middle" width="25%" class="blue-normal"><%If CInt(intTotalPage) <> 0 Or intTotalPage <> "" Then%>Pages <%=intCurPage%>/<%=intTotalPage%><%End If%>&nbsp;&nbsp;</td>
						<td valign="middle" align="right" width="28%" class="blue-normal"><%If CInt(intCurPage) <> 1 Then%><a href="javascript:viewpage(2);" onMouseOver="self.status='Move Previous';return true" onMouseOut="self.status='';return true">Previous</a><%End If%><%If CInt(intCurPage) <> 1 And  CInt(intCurPage) <> CInt(intTotalPage) Then%>/<%End If%><%If CInt(intCurPage) <> CInt(intTotalPage) And (CInt(intTotalPage) <> 0 Or intTotalPage <> "") Then%><a href="javascript:viewpage(3);" onMouseOver="self.status='Move Next';return true" onMouseOut="self.status='';return true"> Next</a><%End If%>&nbsp;&nbsp;&nbsp;</td>
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
				<tr>
				  <td class="red">&nbsp;<%=strError%></td>
				</tr>
<%		End If%>				

		  <tr>
         	<td class="red" align="center" valign="middle"><b><%=strError1%></b></td>
		  </tr>	          

<%	End If
	Set objDatabase = Nothing
%>
          
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
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="txtstatus" value="<%=strStatus%>">
<input type="hidden" name="P" value="<%=intCurPage%>">
<input type="hidden" name="S" value="<%=fgSort%>">
<input type="hidden" name="name" value="<%=strSName%>">
<input type="hidden" name="depart" value="<%=intDepart%>">

</form>
</body>
</html>