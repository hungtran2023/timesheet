<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, strError, PageSize, fgRight 'view all or Not
	Dim intUserStyle

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
			         "class='c' OnMouseOver = 'self.status=&quot;View Timesheet&quot; ; return true' OnMouseOut =" &_
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

	fgSort = Request.Form("S")
	
	intCurPage = trim(Request.Form("P"))
	If intCurPage = "" Then
		intCurPage = 1
	End If		

	strName = Request.Form("name")
	
	intUserStyle=Request.Form("lbUserStyle")
	if intUserStyle ="" then intUserStyle=1
	
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
' Analyse query and prepare staff list
'--------------------------------------------------

	strAct = Request.QueryString("act")
	If strAct = "" Then
		strAct = Request.Form("txtstatus")
	End If
		
	if isObject(rsStaff) then 
        set rsStaff = nothing 
    end if 

    strConnect = Application("g_strConnect")												' Connection string 				
	Set objDatabase = New clsDatabase 

 If (not IsObject(Session("rsStaff")) and cint(intUserStyle)=1) OR (not IsObject(Session("rsTPStaff")) and cint(intUserStyle)=3) or (strAct ="") then
   blnFirstConnect=true
 else
    blnFirstConnect=false
 end if

'Response.Write blnFirstConnect & " -- StrAct: " & strAct '(Session("rsStaff") is Nothing)

	If blnFirstConnect Then					' Call this page the first
		fgSort = "N"
		
		If objDatabase.dbConnect(strConnect) Then
		     if cint(intUserStyle)=1 then
    		 
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
    					
			    myCmd("StaffID") = session("USERID")
			    myCmd("level") = 0

		        strSQL="SELECT PersonID as StaffID,Fullname,Department,JobTitle  FROM HR_Employee a  "						
		        If fgRight Then						' View all		  
			        myCmd("fgCheck") = 0
		        Else
			        strSQL = strSQL & " WHERE a.PersonID " '& session("USERID")
			        myCmd("fgCheck") = 1 
		        End If
		        myCmd("strSQL") = strSQL
		        
			    On Error Resume Next	
			    rsStaff.Open myCmd,,adOpenStatic,adLockBatchOptimistic
			    If Err.number > 0 then
				    strError = Err.Description
			    End If
			    Err.Clear
			    If Not rsStaff.EOF Or rsStaff.RecordCount > 0 Then
				    intTotalPage = pageCount(rsStaff, intPageSize)
				    rsStaff.MoveFirst
				    rsStaff.Move (intCurPage-1)*intPageSize
				    strLast = Outbody(rsStaff, intPageSize)

				    Set session("rsStaff") = rsStaff
			    End if
			    Set myCmd = Nothing
		    Else
		         strSQL="SELECT TPUserID as StaffID, Fullname, DepartmentID,'' as JobTitle,Department FROM HR_TPStaff ORDER BY Fullname"		    
		        Call GetRecordset(strSQL ,rsStaff) 

		        If Not rsStaff.EOF Or rsStaff.RecordCount > 0 Then
				        intTotalPage = pageCount(rsStaff, intPageSize)
				        rsStaff.MoveFirst
				        rsStaff.Move (intCurPage-1)*intPageSize
				        strLast = Outbody(rsStaff, intPageSize)

				        Set session("rsTPStaff") = rsStaff
			    End if
		    End If		    
		 else
		   strError = objDatabase.strMessage   
		 end if
    	 Set objDatabase = Nothing
	Else															' Submit this page
	    
	    if cint(intUserStyle)=1 then
		    Set rsStaff = session("rsStaff")
		else
		    Set rsStaff = session("rsTPStaff")
		end if

		rsStaff.MoveFirst
		If recCount(rsStaff) >= 0 Then
			intTotalPage = pageCount(rsStaff, intPageSize)
		
			Select Case strAct
				Case "vpsn"											' Sort by fullname
					
					strStatus = strAct

'--------------------------------------------------
' This If..Then..End If to check status
' of the form when it go back					
'--------------------------------------------------

					If Request.QueryString("b") <> "" Then
						If fgSort = "A" Then
							fgSort = "D"
						ElseIf fgsort = "D" Then
							fgSort = "A"
						End If
					End If
					
'--------------------------------------------------
' End of checking		
'--------------------------------------------------								
					
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
				
					strStatus = strAct
					
'--------------------------------------------------
' This If..Then..End If to check status
' of the form when it go back					
'--------------------------------------------------

					If Request.QueryString("b") <> "" Then
						If fgSort = "A" Then
							fgSort = "D"
						ElseIf fgsort = "D" Then
							fgSort = "A"
						End If
					End If
					
'--------------------------------------------------
' End of checking		
'--------------------------------------------------								

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

					strStatus = strAct

'--------------------------------------------------
' This If..Then..End If to check status
' of the form when it go back					
'--------------------------------------------------

					If Request.QueryString("b") <> "" Then
						If fgSort = "A" Then
							fgSort = "D"
						ElseIf fgsort = "D" Then
							fgSort = "A"
						End If
					End If
					
'--------------------------------------------------
' End of checking		
'--------------------------------------------------								

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
					
					strStatus = Request.Form("txtstatus")
				Case "vpa2"											' When user click Previous link	
					If CInt(intCurPage) > 1 Then
						intCurPage = CInt(intCurPage) - 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
				Case "vpa3"											' When user click Next link		

					If CInt(intCurPage) < CInt(intTotalPage) Then
						intCurPage = CInt(intCurPage) + 1
					End If
					
					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					strStatus = Request.Form("txtstatus")
					
				Case "vra1"											' When user click button "Search"
					strSName = Request.Form("txtname")
										
					If strSName <> "" And CInt(intDepart) = 0 Then
						If InStr(1,Request.Form("txtname"),"'") = 0 Then
							rsStaff.Filter = "FullName LIKE '%" & strSName & "%'"
						Else
							rsStaff.Filter = "FullName LIKE #" & strSName & "#"
						End If	
					End If
					
					If Not rsStaff.EOF Or rsStaff.RecordCount > 0 Then
						intCurPage = 1
						intTotalPage = pageCount(rsStaff, intPageSize)

						rsStaff.MoveFirst
						rsStaff.Move CInt((intCurPage-1)*intPageSize)
					Else
						strError = "No data for your request."
						rsStaff.Filter = ""
					End If	
					
					strStatus = Request.Form("txtstatus")
				Case "vra2"	
				    										' When user click button "Show all"
					rsStaff.Filter = ""
					if cint(intUserStyle)=1 then
		                Set session("rsStaff") = rsStaff
		            else
		                Set session("rsTPStaff") = rsStaff
		            end if
					
					intTotalPage = pageCount(rsStaff, intPageSize)

					rsStaff.MoveFirst
					rsStaff.Move CInt((intCurPage-1)*intPageSize)
					
					intDepartmentID = 0
					strName = ""
					
					strStatus = Request.Form("txtstatus")
			End Select	 

			strLast = Outbody(rsStaff, intPageSize)
		End If		
	End If
		
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
	If strChoseMenu = "" Then strChoseMenu = "B"
	
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
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet - Main Menu</title>

<link rel="stylesheet" href="../../timesheet.css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

    function sort(kind) {
        document.frmreport.action = "rpt_list_staff.asp?act=vps" + kind;
        document.frmreport.submit();
    }

    function viewpage(kind) {
        var intpage = parseInt(window.document.frmreport.txtpage.value, 10);
        var curpage = "<%=CInt(intCurPage)%>";
        var pagetotal = "<%=CInt(intTotalPage)%>";

        if (kind == 1) {
            window.document.frmreport.txtpage.value = intpage
            if ((intpage > 0) & (intpage <= pagetotal) & (intpage != curpage)) {
                document.frmreport.action = "rpt_list_staff.asp?act=vpa" + kind;
                document.frmreport.submit();
            }
        }
        else {
            document.frmreport.action = "rpt_list_staff.asp?act=vpa" + kind;
            document.frmreport.submit();
        }
    }

    function search(kind) {
        

        document.frmreport.action = "rpt_list_staff.asp?act=vra" + kind;
        document.frmreport.submit();

    }

    function viewtms(varid) {
        
        var intType;
        intType = "<%=intUserStyle %>";

        window.document.frmreport.txthidden.value = varid;
        if (intType == 1)
            window.document.frmreport.action = "rpt_staff_tms.asp";
        else
            window.document.frmreport.action = "rpt_TPstaff_tms.asp";
            
        window.document.frmreport.submit();

    }

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" Topmargin="0" marginwidth="0" marginheight="0">
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
		      <table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
		        <tr <%If strError="" Then%> bgcolor="#FFFFFF" <%Else%> bgcolor="#E7EBF5" <%End If%>>
				  <td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=strError%></b></td>
		        </tr>
		        <tr align="center"> 
		          <td class="blue-normal" height="30" align="right" width="12%">&nbsp;&nbsp;Search for Name &nbsp;</td>
  				  <td class="blue" height="30" align="left" width="30%"> 
				    <input type="text" name="txtname" value="<%=showvalue(strSName)%>" class="blue-normal" size="15" style=" width:100%">
				  </td>

				  <td class="blue" height="30" align="left"  width="26%"> &nbsp;&nbsp;
				    <!--<select id="lbdepartment" size="1" name="lbdepartment" class="blue-normal">-->
					  <select id="lbUserStyle" size="1" name="lbUserStyle" class="blue-normal" style="width:97%">
					    <option value="1" <% if intUserStyle=1 then%> selected <%end if %> >Atlas Staff</option>
				        <option value="3" <% if intUserStyle=3 then%> selected <%end if %> >Contract Staff</option>
                    </select>
				  </td>
				  <td class="blue-normal" height="30" align="left" width="32%"> 
				    <table width="180" border="0" cellspacing="5" cellpadding="0" height="20" name="aa">
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
			    <tr align="center"> 
				  <td class="title" height="50" align="center" colspan="5">Timesheet Overview</td>
			    </tr>
			  </table>
		    </td>
		  </tr>
		  <tr> 
		    <td height="100%"> 
		      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:'79%'" height="365">
		        <tr> 
		          <td bgcolor="#FFFFFF" valign="top"> 
		            <table width="100%" border="0" cellspacing="0" cellpadding="0">
		              <tr> 
		                <td bgcolor="#617DC0"> 
		                  <table width="100%" border="0" cellspacing="1" cellpadding="5">		                  
		                    <tr bgcolor="#8CA0D1"> 
<%if intUserStyle=1 then %>		                    
		                      <td class="blue" align="center" width="35%"><a href="javascript:sort('n')" onMouseOver="self.status='Sort by Full Name';return true" onMouseOut="self.status='';return true" class="c">Full Name</a></td>
		                      <td class="blue" align="center" width="206"><a href="javascript:sort('t')" onMouseOver="self.status='Sort by Job Title';return true" onMouseOut="self.status='';return true" class="c">Job Title</a> </td>
		                      <td class="blue" align="center" width="168"><a href="javascript:sort('d')" onMouseOver="self.status='Sort by Department';return true" onMouseOut="self.status='';return true" class="c">Department</a></td>
<%else%>

                           	 <td class="blue" align="center" width="188"><a href="javascript:sort('n')" onMouseOver="self.status='Sort by Full Name';return true" onMouseOut="self.status='';return true" class="c">StaffID</a></td>
		                      <td class="blue" align="center" width="206"><a href="javascript:sort('t')" onMouseOver="self.status='Sort by Job Title';return true" onMouseOut="self.status='';return true" class="c">Job Title</a> </td>
		                      <td class="blue" align="center" width="168"><a href="javascript:sort('d')" onMouseOver="self.status='Sort by Department';return true" onMouseOut="self.status='';return true" class="c">Department</a></td>

<%end if %>		                      
		                    </tr>
<%
	Response.Write(strLast)
%>                            
		                  </table>
		                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		                    <tr> 
		                      <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;&nbsp;* Click on the exact name to view timesheet.</td>
		                      <td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">&nbsp;</td>
		                    </tr>
		                    <tr> 
		                      <td bgcolor="#FFFFFF" height="20" class="blue-normal" colspan="2">&nbsp;&nbsp;* Click 
		                         on each column header to sort the list by alphabetical order </td>
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
<input type="hidden" name="P" value="<%=intCurPage%>">
<input type="hidden" name="S" value="<%=fgSort%>">
<input type="hidden" name="txtstatus" value="<%=strStatus%>">
<input type="hidden" name="name" value="<%=strSName%>">
<input type="hidden" name="depart" value="<%=intDepart%>">
</form>
</body>
</html>