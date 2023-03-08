<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("Inhouse")) = False Then
		Response.Redirect("message.htm")
	End If

'****************************************
' function: outBody
' Description: table of list project
' Parameters: 
'			  
' Return value: 
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc, ByVal psize)
	strOut = ""
	if not rsSrc.EOF then
		For i = 1 to psize
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' width='8%' class='blue-normal' align='center'>" & rsSrc.Bookmark & "</td>" &_
			         "<td valign='top' width='60%' class='blue-normal'><a href='javascript:_edit(" & chr(34) & rsSrc.Bookmark &_
			         chr(34) & ");' class='c' OnMouseOver = 'self.status=&quot;Edit&quot; ; return true;' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("Department")) & "</a></td>" &_
			         "<td valign='top' width='20%' class='blue-normal' align='center'>" & Showlabel(rsSrc("DepartmentCode")) & "</td>" &_
			         "<td valign='top' width='12%' class='blue-normal' align='center'><input type='checkbox' name='chk' value='" &_
			         rsSrc("DepartmentID") & "," & rsSrc.Bookmark & "'></td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function
'-------------------------------------------------
Dim gMessage

If IsEmpty(Session("strHTTP")) Then
	Call MakeHTTP
End if
strtmp1 = Replace(logoff, "XX", session("strHTTP") & "admin/")
strFunction = "<div align='right'>" & Help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
			"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"

stract = Request.QueryString("act")
if strAct = "" and Request.QueryString("navi")="" and Request.QueryString("Go")="" then 'first run
	Call freeAdmininput
	Call freeRole
	Call freeRoleAss
	Call freelistRole
end if
strNameItem = ""
strNameItem2 = ""

select case stract
case "SAVE"
	strNameItem = Request.Form("txtname")
	strNameItem2 = Request.Form("txtname2")
	varID = Request.Form("txthidden")
	gMessage = ""
	if not isEmpty(session("rsItem")) then
		set rsItem = session("rsItem")
		rsItem.MoveFirst

		if Instr(strNameItem, "'")>0 then
			tmp1 = "#" & strNameItem & "#"
		else
			tmp1 = "'" & strNameItem & "'"
		end if
		
		rsItem.Find "Department like " & tmp1,, adSearchForward
		If not rsItem.EOF then
			if rsItem.Bookmark<>int(varID) then
				gMessage = chr(34) & strNameItem & chr(34) & " has already been inputted."
				stract = "EDIT"
			else
				rsItem.MoveFirst
				tmp1 = "'" & strNameItem2 & "'"
				rsItem.Find "DepartmentCode like " & tmp1,, adSearchForward
				If not rsItem.EOF then
					if rsItem.Bookmark<>int(varID) then
						gMessage = chr(34) & strNameItem2 & chr(34) & " has already been inputted."
						stract = "EDIT"
					end if
				end if
			end if
		end if
	end if
	if gMessage = "" then
		rsItem.Bookmark = int(varID)
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			strNameItem = replace(strNameItem, "'", "''")
			strQuery = "UPDATE ATC_Department SET Department = '" & strNameItem & "', DepartmentCode = '" &_
						strNameItem2 & "' WHERE departmentID = " & rsItem("DepartmentID")
			if not objDB.runActionquery(strQuery) then
		      gMessage = objDb.strMessage
		    else
		  	  gMessage = "Saved successfully."
		  	  strNameItem = ""
		  	  strNameItem2 = ""
		    end if
		    rsItem.close()
		    session("rsItem") = empty
		    session("READYIN") = false
		    objDb.dbDisConnect
		else
		    gMessage =  objDb.strMessage
		end if
		Set objDb = Nothing
	end if
case "ADD"
	strNameItem = Request.Form("txtname")
	strNameItem2 = Request.Form("txtname2")
	gMessage = ""
	if not isEmpty(session("rsItem")) then
		set rsItem = session("rsItem")
		' testing department
		rsItem.MoveFirst
		if Instr(strNameItem, "'")>0 then
			tmp1 = "#" & strNameItem & "#"
		else
			tmp1 = "'" & strNameItem & "'"
		end if
		rsItem.Find "Department like " & tmp1,, adSearchForward
		If Not rsItem.EOF and not rsItem.BOF Then
			if rsItem.Bookmark<>int(varID) then
				gMessage = chr(34) & strNameItem & chr(34) & " has already been inputted."
			Else
				' testing department code
				rsItem.MoveFirst
				tmp1 = "'" & strNameItem2 & "'"
				rsItem.Find "DepartmentCode like " & tmp1,, adSearchForward
				If Not rsItem.EOF and not rsItem.BOF Then
					if rsItem.Bookmark<>int(varID) then
						gMessage = chr(34) & strNameItem2 & chr(34) & " has already been inputted."
					end if
				end if
			end if
		end if
	end if
	if gMessage = "" then
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			strNameItem = replace(strNameItem, "'", "''")
			strQuery = "INSERT INTO ATC_Department(Department, DepartmentCode) VALUES('" &_
						 strNameItem & "', '" & strNameItem2 & "')"
			if not objDB.runActionquery(strQuery) then
		      gMessage = objDb.strMessage
		    else
		  	  gMessage = "Added successfully."
		  	  strNameItem = ""
		  	  strNameItem2 = ""
		    end if
		    rsItem.close()
		    session("rsItem") = empty
		    session("READYIN") = false
		    objDb.dbDisConnect
		else
		    gMessage =  objDb.strMessage
		end if
		Set objDb = Nothing
	end if
case "DELETE"
	countU = Request.Form("chk").Count
	if countU>0 then
	  Set objDb = New clsDatabase
	  strConnect = Application("g_strConnect")
	  ret = objDb.dbConnect(strConnect)
	  if ret then
	    strDonot = ""
	    set rsItem = session("rsItem")
	    for ii = 1 to countU
	  		varID = Request.Form("chk")(ii)
	  		varBookmark = Mid(varID, Instr(varID, ",") + 1, len(varID))
	  		varID = Mid(varID, 1, Instr(varID, ",") - 1)
			strQuery = "DELETE ATC_Department WHERE DepartmentID = " & varID
			if not objDB.runActionquery(strQuery) then
				rsItem.Bookmark = int(varBookmark)
				strDonot = strDonot & " " & rsItem("Department") & ","
			end if
	    next
	    if strDonot<>"" then
	      strDonot = Mid(strDonot, 1, len(strDonot)- 1)
	      gMessage = chr(34) & strDonot & chr(34) & " have been used."
	    else
	  	  gMessage = "Deleted successfully."
	    end if
	    rsItem.close()
	    session("rsItem") = empty
	    session("READYIN") = false
	    Session("CurPagein") = empty
	    objDb.dbDisConnect
	  else
	    gMessage =  objDb.strMessage
	  end if
	  Set objDb = Nothing
	end if
end select

If session("READYIN")<> True Then
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)
	strQuery = "select * FROM ATC_Department ORDER BY Department"
	If objDb.openRec(strQuery) Then
	  objDb.recDisConnect
	  IF not objDb.noRecord then
		set rsItem = objDb.rsElement.Clone
		session("READYIN") = true
		rsItem.MoveFirst
		session("NumPagein") = pageCount(rsItem, PageSizeDefault)
		if isEmpty(Session("CurPagein")) then 
			Session("CurPagein") = 1
		else
			if Session("CurPagein") > Session("NumPagein") then
				Session("CurPagein") = Session("NumPagein")
			elseif Session("CurPagein") = 0 then
				Session("CurPagein") = 1
			end if
		end if
		set session("rsItem") = rsItem
	  End if
	  objDb.CloseRec
	Else
	  gMessage = objDb.strMessage	  
	End if
	Set objDb = Nothing
Else
	set rsItem = session("rsItem")
End if
if isEmpty(session("NumPagein")) then session("NumPagein") = 0
if isEmpty(Session("CurPagein")) then Session("CurPagein") = 0

if stract = "EDIT" and strNameItem="" then
	varID = Request.Form("txthidden")
	rsItem.Bookmark = int(varID)
	strNameItem = rsItem("Department")
	strNameItem2 = rsItem("DepartmentCode")
end if

varNavi = Request.QueryString("navi")
if varNavi <> "" and session("READYIN") then
	tmpi = session("CurPagein")
	select case varNavi
		case "PREV"
			if tmpi > 1 then
				tmpi = tmpi - 1
			else
				tmpi = 1
			end if
		case "NEXT"
			if tmpi < Session("NumPagein") then
				tmpi = tmpi + 1
			else
				tmpi = Session("NumPagein")
			end if
	End select
	session("CurPagein") = tmpi
end if

varGo = Request.QueryString("Go")
if varGo <> "" then Session("CurPagein") = CInt(varGo)

strLast = ""
if session("READYIN") then
	rsItem.MoveFirst
	rsItem.Move (session("CurPagein")-1)*PageSizeDefault
	strLast = Outbody(rsItem, PageSizeDefault)
end if
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_admin.htm")
curpage = 2
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

<link rel="stylesheet" href="../timesheet.css" type="text/css">
<script language="javascript" src="../library/library.js"></script>
<script>
function _edit(varid){
	document.frminput.txthidden.value = varid;
	document.frminput.action = "department.asp?act=EDIT"
	document.frminput.target = "_self"
	document.frminput.submit();
}

function chkremove() {
  fg = false;
  with (document.frminput) {
	 len = elements.length;
     for(var ii=0; ii<len; ii++) {
		if ((elements[ii].name == "chk") && (elements[ii].checked)) {
			fg = true;
			break;
		}
	}
  }
 if (fg == false) alert("No department selected.")
 return(fg)
}

function _cancel(){
  	document.frminput.action = "department.asp?act=CANCEL"
	document.frminput.target = "_self";
	document.frminput.submit();
}

function testdata() {
	var tmp = document.frminput.txtname2.value;
	var inttmp = tmp.charCodeAt(0);
	if (((inttmp>=65) && (inttmp<=90)) || ((inttmp>=97)&&(inttmp<=122)))
		return true;
	else
		alert("Please enter a character in the alphabet.")
		return false;
}

function _save(){
var varid = "<%=varID%>";
	if (testdata()==true) {
		document.frminput.txthidden.value = varid;
  		document.frminput.action = "department.asp?act=SAVE"
		document.frminput.target = "_self";
		document.frminput.submit();
	}
}

function _add(){
	var tmp = document.frminput.txtname.value;
	document.frminput.txtname.value = alltrim(tmp);
	if(document.frminput.txtname.value=="") {
		alert("Please enter a value");
	}
	else {
		if(testdata()==true) {
  			document.frminput.action = "department.asp?act=ADD"
			document.frminput.target = "_self";
			document.frminput.submit();
		}
	}
}

function _remove() {
  if (chkremove()==true) {
  	document.frminput.action = "department.asp?act=DELETE"
	document.frminput.target = "_self";
	document.frminput.submit();
  }
}

function next() {
var curpage = <%=session("CurPagein")%>
var numpage = <%=session("NumPagein")%>
	if (curpage < numpage) {
		document.frminput.action = "department.asp?navi=NEXT"
		document.frminput.target = "_self"
		document.frminput.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagein")%>
var numpage = <%=session("NumPagein")%>
	if (curpage > 1) {
		document.frminput.action = "department.asp?navi=PREV";
		document.frminput.target = "_self"
		document.frminput.submit();
	}
}

function go() {
	var numpage = <%=session("NumPagein")%>;
	var curpage = <%=session("CurPagein")%>;
	var intpage = document.frminput.txtpage.value;
	intpage = parseInt(intpage, 10)
	if ((intpage > 0) && (intpage <= numpage) && (intpage != curpage)) {
		document.frminput.action = "department.asp?Go=" + intpage;
		document.frminput.target = "_self";
		document.frminput.submit();		
	}
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frminput" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr> 
      <td align="center"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%if gMessage="" then%>#FFFFFF<%else%>#E7EBF5<%end if%>">
			<td class="red" colspan="3" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		</tr>
          <tr align="center"> 
            <td class="title" height="50" align="center" colspan="3">Department</td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" height="30" align="left" width="96"> 
              &nbsp;&nbsp;Department Name&nbsp; </td>
            <td class="blue" align="left" width="175"> 
              <input type="text" name="txtname" class="blue-normal" size="20" style=" width:165" maxlength="50" value="<%=strNameItem%>">
            </td>
            <td class="blue-normal" height="60" align="left" width="336" rowspan="2"> 
<%if stract <> "EDIT" then%>
              <table width="60" border="0" cellspacing="5" cellpadding="0" height="20">
                <tr> 
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
					<a href="javascript:_add();" onMouseOver="self.status='Add new'; return true;" onMouseOut="self.status='';" class="b">Add</a></td>
                </tr>
              </table>
<%else%>
              <table width="120" border="0" cellspacing="5" cellpadding="0" height="20">
                <tr> 
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
					<a href="javascript:_save();" onMouseOver="self.status='Save'; return true;" onMouseOut="self.status='';" class="b">Save</a></td>
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" class="blue" align="center">
					<a href="javascript:_cancel();" onMouseOver="self.status='Cancel'; return true;" onMouseOut="self.status='';" class="b">Cancel</a></td>
                </tr>
              </table>
<%end if%>
            </td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="left" width="96" nowrap> &nbsp;&nbsp;Department Code&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname2" class="blue-normal" size="5" maxlength="1" style=" width:50" value="<%=strNameItem2%>">
            </td>
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
                    <table width="100%" border="0" cellspacing="1" cellpadding="5">
                      <tr bgcolor="8CA0D1" align="left"> 
                        <td class="blue" bgcolor="8CA0D1" width="8%" align="right">&nbsp;</td>
                        <td class="blue" width="60%" align="center">Department</td>
                        <td class="blue" width="20%" align="center">Dept. Code</td>
                        <td class="blue" width="12%">&nbsp;</td>
                      </tr>
<%Response.Write(strLast)%>
                    </table>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#FFFFFF" height="20" class="blue-normal" width="76%">&nbsp;</td>
                        <td bgcolor="#FFFFFF" height="20" class="blue" width="24%" align="right">
							<a href="javascript:_remove();" onMouseOver="self.status='Delete'; return true;" onMouseOut="self.status='';">Delete</a>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
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
		          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
		          </td>
		          <td align="center" valign="middle" width="13%" class="blue-normal"> 
		            <input type="text" name="txtpage" class="blue-normal" value="<%=session("Curpagein")%>" size="2" style="width:50">
		          </td>
		          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();" onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
		          </td>
		          <td align="right" valign="middle" width="15%" class="blue-normal">Page <%=session("Curpagein")%>/<%=session("Numpagein")%>&nbsp;&nbsp;</td>
		          <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();" onMouseOver="self.status='Previous page'; return true;" onMouseOut="self.status='';">Previous</a> /
		          <a href="javascript:next();" onMouseOver="self.status='Next page'; return true;" onMouseOut="self.status='';"> Next</a>&nbsp;&nbsp;&nbsp;</td>
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
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(1))
			%>
<input type="hidden" name="txthidden" value>
</form>
</body>
</html>
