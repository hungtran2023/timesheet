<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strError, PageSize, fgRight 'view all or Not

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
	strOut = ""
	If Not rsSrc.EOF Then
		For i = 1 To psize
			If i Mod 2 = 0 Then
				strColor = "#E7EBF5"
			Else
				strColor = "#FFF2F2"
			End If

			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' width='33%' class='blue'><a href='javascript:viewtms(" & rsSrc("StaffID") & ");' " &_
			         "class='c' OnMouseOver = 'self.status=&quot;View Annual Leave&quot; ; return true' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("Fullname")) & "</a></td>" &_
			         "<td valign='top' width='33%' class='blue-normal'>" & Showlabel(rsSrc("JobTitle")) & "</td>" &_
			         "<td valign='top' width='34%' class='blue-normal'>" & Showlabel(rsSrc("Department")) & "</td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	End If
	Outbody = strOut
End Function

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	intDepartmentID = Request.Form("lbdepartment")
	fgSort = Request.Form("S")
	strView = Request.QueryString("view")	
	intCurPage = trim(Request.Form("P"))
	If intCurPage = "" Then
		intCurPage = 1
	End If		
		
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
%>
<script LANGUAGE="javascript">
<!--
	opener.document.location = "../../message.htm";
	window.close();
//-->
</script>
<%
	End If					

	intUserID = session("USERID")

'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then intPageSize = arrPre(1, 0) Else intPageSize = 5
		Set arrPre = Nothing
	Else
		intPageSize = 5
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
			strSQL = "SELECT * FROM ATC_Department ORDER BY Department"

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

	If Request.QueryString("act") = "" Then					' Call this page the first
		fgSort = "N"
		
		strConnect = Application("g_strConnect")
		Set objDatabase = New clsDatabase
	
		If objDatabase.dbConnect(strConnect) Then
			Set rsStaff = Server.CreateObject("ADODB.Recordset")
			rsStaff.CursorLocation = adUseClient			' Set the Cursor Location to Client

			Set myCmd = Server.CreateObject("ADODB.Command")
			Set myCmd.ActiveConnection = objDatabase.cnDatabase
			myCmd.CommandType = adCmdStoredProc
			myCmd.CommandText = "sp_getListEmp"

			Set myParama = myCmd.CreateParameter("StaffID",adInteger,adParamInput)
			myCmd.Parameters.Append myParama
			Set myParamb = myCmd.CreateParameter("level",adTinyInt,adParamInput)
			myCmd.Parameters.Append myParamb
			Set myParamc = myCmd.CreateParameter("strSQL", adVarChar,adParamInput, 5000)
			myCmd.Parameters.Append myParamc
			Set myParamd = myCmd.CreateParameter("fgCheck", adTinyInt,adParamInput)
			myCmd.Parameters.Append myParamd
					
			myCmd("StaffID") = intUserID
			myCmd("level") = 0

			strSQL = "SELECT a.StaffID, FirstName + ' ' + ISNULL(MiddleName,'') + ' '+ ISNULL(LastName,'') AS FullName, a.DepartmentID, c.JobTitle, d.Department" & _
						" FROM ATC_Employees a LEFT JOIN ATC_PersonalInfo b ON a.StaffID = b.PersonID LEFT JOIN ATC_JobTitle c " & _
						" ON a.JobTitleID = c.JobTitleID LEFT JOIN ATC_Department d ON a.DepartmentID = d.DepartmentID WHERE b.fgDelete = 0"
						
			If fgRight Then						' View all		  
				myCmd("fgCheck") = 0
			Else
				strSQL = strSQL & " AND a.StaffID "
				myCmd("fgCheck") = 1 
			End If
			myCmd("strSQL") = strSQL

			On Error Resume Next	
			rsStaff.Open myCmd,,adOpenStatic,adLockBatchOptimistic
			If Err.number > 0 then
				strError = Err.Description
			End If
			Err.Clear
			If Not rsStaff.EOF Then
				intTotalPage = pageCount(rsStaff, intPageSize)
				rsStaff.MoveFirst
				rsStaff.Move (intCurPage-1)*intPageSize
				
				strLast = Outbody(rsStaff, intPageSize)

				If Not IsEmpty(session("rsStaff")) Then
					session("rsStaff") = Empty
				End If	
				Set session("rsStaff") = rsStaff
			End if
			Set myCmd = Nothing
		Else
			strError = objDatabase.strMessage
		End If
		Set objDatabase = Nothing
		
	Else															' Submit this page
	
		Set rsStaff = session("rsStaff")
		rsStaff.MoveFirst
		If recCount(rsStaff) >= 0 Then
			intTotalPage = pageCount(rsStaff, intPageSize)

			Select Case Request.QueryString("act")
				Case "vpsn"											' Sort by fullname
					If fgSort = "N" Or fgSort = "D" Then
						rsStaff.Sort = "FullName ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "FullName DESC"
						fgSort = "D"				
					End If

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vpst"											' Sort by job title
					If fgSort = "N" Or fgSort = "D" Then
						rsStaff.Sort = "JobTitle ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "JobTitle DESC"
						fgSort = "D"				
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vpsd"											' Sort by department		
					If fgSort = "N" Or fgSort = "D" Then
						rsStaff.Sort = "Department ASC"
						fgSort = "A"
					ElseIf fgSort = "A"	Then
						rsStaff.Sort = "Department DESC"
						fgSort = "D"				
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vpa1"											' When user click button "Go"
					If CInt(Request.Form("txtpage")) <= CInt(intTotalPage) Then
						intCurPage = Request.Form("txtpage")
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vpa2"											' When user click Previous link	
					If CInt(intCurPage) > 1 Then
						intCurPage = CInt(intCurPage) - 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vpa3"											' When user click Next link		

					If CInt(intCurPage) < CInt(intTotalPage) Then
						intCurPage = CInt(intCurPage) + 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
				Case "vra1"											' When user click button "Search"
					strSName = Request.Form("txtname")
					intDepart = Request.Form("lbdepartment")
					
					If strSName <> "" And CInt(intDepart) <> 0 Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%' AND DepartmentID=" & intDepart
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "# AND DepartmentID=" & intDepart
						End If		
					ElseIf strSName = "" And CInt(intDepart) <> 0 Then
						rsStaff.Filter = "DepartmentID=" & intDepart
					ElseIf strSName <> "" And CInt(intDepart) = 0 Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%'"
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "#"
						End If	
					End If
					
					If rsStaff.EOF Then
						strError = "No data for your request."
						rsStaff.Filter = ""
					End If	
				Case "vra2"											' When user click button "Show all"
					rsStaff.Filter = ""
					intDepartmentID = 0
			End Select	 
			
			strLast = Outbody(rsStaff, intPageSize)
		End If		
	End If
%>	

<html>
<head>
<title>Timesheet</title>

<style type="text/css">
<!--

-->
</style>
<link rel="stylesheet" href="../../timesheet.css" type="text/css">

<script language="javascript">
<!--

function sort(kind)
{
	document.frmreport.action = "rpt_select_staff.asp?act=vps" + kind;
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
			document.frmreport.action = "rpt_select_staff.asp?act=vpa" + kind;
			document.frmreport.submit();
		}	
	}
	else
	{	
		document.frmreport.action = "rpt_select_staff.asp?act=vpa" + kind;
		document.frmreport.submit();
	}	
}

function search(kind)
{
	if (kind == "1")
	{
		if ((document.frmreport.txtname.value != "") || (document.frmreport.lbdepartment.options[document.frmreport.lbdepartment.selectedIndex].value != "0"))
		{
			document.frmreport.action = "rpt_select_staff.asp?act=vra" + kind;
			document.frmreport.submit();	
		}
	}	
	else
	{
		document.frmreport.action = "rpt_select_staff.asp?act=vra" + kind;
		document.frmreport.submit();
	}	
}

function viewtms(varid)
{
	opener.document.frmtms.txthidden.value = varid;
	opener.document.frmtms.action = "rpt_view_leave.asp";	
	opener.document.frmtms.submit();
	window.close();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="1" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
  <table width="450" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="80"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%If strError <> "" Then%>        
          <tr bgcolor="#E7EBF5">
   		    <td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
          </tr>
 <%End If%>          
          <tr> 
            <td class="blue-normal" height="30" align="right">&nbsp;&nbsp;Name&nbsp;</td>
  			<td class="blue" height="30" align="left" width="150"> 
			  <input type="text" name="txtname" value="<%=showvalue(strSName)%>" class="blue-normal" size="12" style=" width:120">
			</td>
			<td class="blue-normal" height="30" align="right" width="65">Department&nbsp;</td>
			<td class="blue" height="30" align="left" width="70"> 
			  <select id="lbdepartment" size="1" name="lbdepartment" class="blue-normal">
			    <option value="0" selected>&nbsp;</option>
<%
		If intNum >= 0 Then
			For ii = 0 To intNum
%>                    
			    <option <%If CInt(intDepartmentID)=CInt(varDepartment(0,ii)) Then%> selected <%End If%> value="<%=varDepartment(0,ii)%>"><%=showlabel(varDepartment(1,ii))%></option>
<%
			Next
		End If	
%>					

			  </select>
			</td>
            <td class="blue" align="right" width="32%" valign="middle"> 
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
          </tr>
          <tr> 
            <td class="title" height="50" align="center" colspan="5"> List of Employees</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="5">
                <tr bgcolor="#8CA0D1"> 
                  <td class="blue" bgcolor="#8CA0D1" align="center" width="194"><a href="javascript:sort('n')" onMouseOver="self.status='Sort by Full Name';return true" onMouseOut="self.status='';return true" class="c">Full Name</a> </td>
                  <td class="blue" align="center" width="189"><a href="javascript:sort('t')" onMouseOver="self.status='Sort by Job Title';return true" onMouseOut="self.status='';return true" class="c">Job Title</a> </td>
                  <td class="blue" align="center" width="190"><a href="javascript:sort('d')" onMouseOver="self.status='Sort by Department';return true" onMouseOut="self.status='';return true" class="c">Department</a></td>
                </tr>
<%
	Response.Write(strLast)
%>                            
              </table>

            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#FFFFFF" height="20" class="blue-normal" align="center">&nbsp;</td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="100%">&nbsp;&nbsp;* Click 
              on the exact name to view timesheet.</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" height="20" class="blue-normal">&nbsp;&nbsp;* Click on each column header to sort the list by alphabetical order or hours worked. 
            </td>
          </tr>
        </table>
      </td>
    </tr>    
    <tr align="right" bgcolor="#99A89D"> 
      <td height="20" align="right" bgcolor="#E7EBF5"> 
        <table width="75%" border="0" cellspacing="1" cellpadding="0">
          <tr class="black-normal"> 
            <td align="right" valign="middle" width="27%" class="blue-normal">Page </td>
            <td align="center" valign="middle" width="10%" class="blue-normal"> 
              <input type="text" name="txtpage" size="5" class="blue-normal" value="<%=intCurPage%>">
            </td>
            <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:viewpage(1);" onMouseOver="self.status='';return true"><font color="#990000">Go</font></a></td>
			<td align="right" valign="middle" width="25%" class="blue-normal"><%If CInt(intTotalPage) <> 0 Or intTotalPage <> "" Then%>Pages <%=intCurPage%>/<%=intTotalPage%><%End If%>&nbsp;&nbsp;</td>
			<td valign="middle" align="right" width="28%" class="blue-normal"><%If CInt(intCurPage) <> 1 Then%><a href="javascript:viewpage(2);" onMouseOver="self.status='Move Previous';return true" onMouseOut="self.status='';return true">Previous</a><%End If%><%If CInt(intCurPage) <> 1 And  CInt(intCurPage) <> CInt(intTotalPage) Then%>/<%End If%><%If CInt(intCurPage) <> CInt(intTotalPage) And (CInt(intTotalPage) <> 0 Or intTotalPage <> "") Then%><a href="javascript:viewpage(3);" onMouseOver="self.status='Move Next';return true" onMouseOut="self.status='';return true"> Next</a><%End If%>&nbsp;&nbsp;&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <input type="hidden" name="txthidden" value="<%=intStaffID%>">
  <input type="hidden" name="P" value="<%=intCurPage%>">
  <input type="hidden" name="S" value="<%=fgSort%>">
</form>

</body>
</html>
