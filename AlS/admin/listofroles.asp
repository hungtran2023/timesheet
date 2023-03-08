<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../inc/library.asp"-->
<%
'****************************************
' function: outBody
' Description: table of list function, have 4 column (group, description, updateable, fg)
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize, ByVal whichpage)
	strOut = ""
	rsSrc.MoveFirst
	rsSrc.Move (whichpage-1)*psize
	strGroupRes = ""
	if not rsSrc.EOF then
		For i = 1 to psize
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' width='23%' class='blue'>" &_
			         "<a href='javascript:gopageRoleRelative(" & chr(34) & rsSrc("GroupID") & chr(34) & "," & chr(34) &_
			         chr(34) & "," & chr(34) & "1" & chr(34) & ");' class='c' OnMouseOver = 'self.status=&quot;Role " &_
			         "Informations&quot; ; return true;' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("GroupName")) & "</a></td>" &_
			         "<td valign='top' width='61%' class='blue'>&nbsp;" & Showlabel(rsSrc("Comment")) & "</td>" &_
			         "<td valign='top' align='center' width='15%' class='blue'><a href='javascript:gopageRoleRelative(" &_
			         chr(34) & rsSrc("GroupID") & chr(34) & "," & chr(34) & rsSrc("GroupName") & chr(34) & "," & chr(34) &_
			         "2" & chr(34) & ");' class='c' OnMouseOver = 'self.status=&quot;Role Assignment&quot; ; return true;' " &_
			         " OnMouseOut = 'self.status = &quot;&quot;'>...</a></td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function
'----------------------------------------------------------------------------------------
	Dim strFunction
	Dim objDb, gMessage, PageSize
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("Inhouse")) = False Then
		Response.Redirect("message.htm")
	End If
	
'-------------------------------
' Calculate pagesize
'-------------------------------
	if not isEmpty(session("Preferences")) then
		arrPre = session("Preferences")
		if arrPre(1, 0)>0 then PageSize = arrPre(1, 0) else PageSize = PageSizeDefault
		set arrPre = nothing
	else
		PageSize = PageSizeDefault
	end if
'------------------------------------	
' Get Full Name
'------------------------------------
	If IsEmpty(Session("strHTTP")) Then
		Call MakeHTTP
	End if
	strtmp1 = Replace(logoff, "XX", session("strHTTP") & "admin/")
	strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"
			
'----------------------------------
' Main procedure
'----------------------------------
gMessage = ""
Call freeAdmininput
Call freeRole
Call freeRoleAss

If isEmpty(session("READYROLE")) then session("READYROLE") = false

if session("READYROLE")<> True Then
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
	strQuery = "SELECT * FROM ATC_Group"
	If objDb.openRec(strQuery) Then
	  objDb.recDisConnect
	  IF not objDb.noRecord then
		set rsRole = objDb.rsElement.Clone
		session("READYROLE") = true
		rsRole.MoveFirst
		session("NumPagerole") = pageCount(rsRole, PageSize)
		if isEmpty(Session("CurPagerole")) then 
			Session("CurPagerole") = 1
		else
			if Session("CurPagerole") > Session("NumPagerole") then
				Session("CurPagerole") = Session("NumPagerole")
			elseif Session("CurPagerole") = 0 then
				Session("CurPagerole") = 1
			end if
		end if
		set session("rsRoleCache") = rsRole
	  Else
		Session("CurPagerole") = 0
		Session("NumPagerole") = 0
	  End if
	  objDb.CloseRec
	Else
	  gMessage = objDb.strMessage	  
	End if
	Set objDb = Nothing
End if


varNavi = Request.QueryString("navi")
if varNavi<>"" then
	tmpi = Session("CurPagerole")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPagerole") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPagerole")
			end if
	End select
	Session("CurPagerole") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPagerole") = CInt(varGo)

strLast=""
if not isEmpty(session("rsRoleCache")) then
	set rsRole = session("rsRoleCache")
	strLast = OutBody(rsRole, pageSize, Session("CurPagerole"))
end if

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_admin.htm")
curpage = 7
If arrPageTemplate(1)<>"" then
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(0) = Replace(arrTmp(0),"@@function", strfunction)
	for i = 1 to NumOfAdminMenu
		if i <> curpage then
			arrTmp(0) = Replace(arrTmp(0),"@@markin"&cstr(i)&"@@", "")
			arrTmp(0) = Replace(arrTmp(0),"@@markout"&cstr(i)&"@@", "")
		else
			arrTmp(0) = Replace(arrTmp(0),"@@markin"&cstr(i)&"@@", "<font color='#CD0000'>")
			arrTmp(0) = Replace(arrTmp(0),"@@markout"&cstr(i)&"@@", "</font>")
		end if
	next
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css">

<script language="javascript" src="../library/library.js"></script>
<script>
function next() {
var curpage = <%=session("CurPagerole")%>
var numpage = <%=session("NumPagerole")%>
	if (curpage < numpage) {
		document.frmrole.action = "listofroles.asp?navi=NEXT";
		document.frmrole.target = "_self";
		document.frmrole.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagerole")%>
var numpage = <%=session("NumPagerole")%>
	if (curpage > 1) {
		document.frmrole.action = "listofroles.asp?navi=PREV";
		document.frmrole.target = "_self";
		document.frmrole.submit();
	}
}

function go() {
	var numpage = <%=session("NumPagerole")%>
	var curpage = <%=session("CurPagerole")%>
	var intpage = document.frmrole.txtpage.value
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frmrole.action = "listofroles.asp?Go=" + intpage;
		document.frmrole.target = "_self";
		document.frmrole.submit();		
	}
}

function gopageRoleRelative(varid, varname, ftype) {
	window.document.frmrole.txtroleid.value = varid;
	window.document.frmrole.txthiddenrolename.value = varname;
	if(ftype=="1")
		window.document.frmrole.action = "roles.asp?outside=1";
	else
		window.document.frmrole.action = "roleassignment.asp?outside=1";
	window.document.frmrole.target = "_self";
	window.document.frmrole.submit();
}

function addnew() {
  	document.frmrole.action = "roles.asp";
	document.frmrole.target = "_self";
	document.frmrole.submit();
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmrole" method="post">
			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			'begin of @@Conntent
			%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor=<%if gMessage="" then%>"#FFFFFF"<%else%>"#E7EBF5"<%end if%>>
		 <td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
        <tr valign="middle"> 
          <td class="title" height="50" align="center"> List of Roles</td>
        </tr>
        <tr align="center"> 
          <td class="blue" height="20" align="left">&nbsp;&nbsp;<a href="javascript:gopageRoleRelative('', '', '1');" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add 
            New</a> </td>
        </tr>        
      </table>
    </td>
  </tr>
  <tr> 
    <td height="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
        <tr>
          <td bgcolor="#FFFFFF" valign="top"> 
<%if strLast<>"" then%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
                              <td class="blue" align="center" width="23%">Role Name</td>
                              <td class="blue" align="center" width="61%">Comment</td>
                              <td class="blue" align="center" width="15%">Assignment</td>
                            </tr>
	<%
	'--------------------------------------------------
	' Write the body of HTML page (menu)
	'--------------------------------------------------	
	Response.Write(strLast)
	%>					  </table>
                        </td>
                      </tr>
                    </table>
<%end if%>
		    </td>
		  </tr>
		</table>
        </td>
      </tr>
<%if strLast<>"" then%>
  <tr> 
    <td> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
		  <tr> 
		    <td align="right" bgcolor="#E7EBF5"> 
		      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
		        <tr class="black-normal"> 
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageRole")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status=''"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("CurPageRole")%>/<%=session("NumPageRole")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Go to previous page'; return true;" onMouseOut="self.status=''">Previous</a> /
		          <a href="javascript:next();" onMouseOver="self.status='Go to next page'; return true;" onMouseOut="self.status=''"> Next</a>&nbsp;&nbsp;&nbsp;</td>
		        </tr>
		      </table>
		    </td>
		  </tr>
		</table>
    </td>
  </tr>
<%end if%>
</table>
<%'end of @@content
  Response.Write(arrTmp(1))
%>
<input type="hidden" name="txtroleid">
<input type="hidden" name="txthiddenrolename">
</form>
</body>
</html>