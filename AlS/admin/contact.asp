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
			         "<td valign='top' width='6%' class='blue-normal' align='center'>" & rsSrc.Bookmark & "</td>" &_
			         "<td valign='top' width='40%' class='blue-normal'><a href='javascript:_edit(" & chr(34) & rsSrc.Bookmark &_
			         chr(34) & ");' class='c' OnMouseOver = 'self.status=&quot;Edit&quot; ; return true;' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>&nbsp;" & Showlabel(rsSrc("FullName")) & "</a></td>" &_
			         "<td valign='top' width='43%' class='blue-normal'>&nbsp;" & rsSrc("EmailAddress") & "</td>" &_
			         "<td valign='top' width='11%' class='blue-normal' align='center'><input type='checkbox' name='chk' value='" &_
			         rsSrc("PersonID") & "," & rsSrc.Bookmark & "'></td>" &_
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
strFunction = "<div align='right'>" & help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
			"&nbsp;&nbsp;&nbsp" & strtmp1 & "&nbsp;&nbsp;&nbsp;</div>"

stract = Request.QueryString("act")
if stract="" and Request.QueryString("navi")="" and Request.QueryString("Go")="" then
	Call freeAdmininput
	Call freeRole
	Call freeRoleAss
	Call freelistRole
end if
strNameItem1 = ""
strNameItem2 = ""
strNameItem3 = ""
strNameItem4 = ""
strNameItem5 = ""
gMessage = ""
select case stract
case "SAVE"
	gMessage = ""
	strNameItem1 = Request.Form("txtname1")
	strNameItem2 = Request.Form("txtname2")
	strNameItem3 = Request.Form("txtname3")
	strNameItem4 = Request.Form("optgender")
	strNameItem5 = Request.Form("lstcomp")
	varID = Request.Form("txthidden")
	set rsItem = session("rsItem")

	rsItem.Bookmark = int(varID)
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		strNameItem1 = replace(strNameItem1, "'", "''")
		strNameItem2 = replace(strNameItem2, "'", "''")
		strNameItem3 = replace(strNameItem3, "'", "''")
		strQuery = "UPDATE ATC_PersonalInfo SET Firstname = '" & strNameItem1 & "', lastname = '" & strNameItem2 &_
					"', EmailAddress = '" & strNameItem3 & "', Gender = " & strNameItem4 & ", CompanyID = " &_
					strNameItem5 & " WHERE PersonID = " & rsItem("PersonID")
		if not objDB.runActionquery(strQuery) then
	      gMessage = objDb.strMessage
	      stract = "EDIT"
	    else
	      strNameItem1 = ""
		  strNameItem2 = ""
		  strNameItem3 = ""
		  strNameItem4 = ""
		  strNameItem5 = ""
	  	  gMessage = "Saved successfully."
	    end if
	    rsItem.close()
	    session("rsItem") = empty
	    session("READYIN") = false
	    set rsComp = session("rsComp")
	    rsComp.close()
	    session("rsComp") = empty
	    objDb.dbDisConnect
	else
	    gMessage =  objDb.strMessage
	end if
	Set objDb = Nothing

case "ADD"
	gMessage = ""
	strNameItem1 = Request.Form("txtname1")
	strNameItem2 = Request.Form("txtname2")
	strNameItem3 = Request.Form("txtname3")
	strNameItem4 = Request.Form("optgender")
	strNameItem5 = Request.Form("lstcomp")
	
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		strNameItem1 = replace(strNameItem1, "'", "''")
		strNameItem2 = replace(strNameItem2, "'", "''")
		strNameItem3 = replace(strNameItem3, "'", "''")
		strQuery = "INSERT INTO ATC_PersonalInfo(Firstname, Lastname, EmailAddress, Gender, CompanyID, Usertype) VALUES('" &_
					strNameItem1 & "', '" & strNameItem2 & "', '" & strNameItem3 & "', " & strNameItem4 &_
					", " & strNameItem5 & ", 0)"
		if not objDB.runActionquery(strQuery) then
	      gMessage = objDb.strMessage
	    else
	  	  gMessage = "Added successfully."
	  	  strNameItem1 = ""
	  	  strNameItem2 = ""
	  	  strNameItem3 = ""
	  	  strNameItem4 = ""
	  	  strNameItem5 = ""
	    end if
	    if not isEmpty(session("rsItem")) then 
			set rsItem = session("rsItem")
			rsItem.close()
			session("rsItem") = empty
		end if
	    if not isEmpty(session("rsComp")) then 
			set rsComp = session("rsComp")
			rsComp.close()
			session("rsComp") = empty
		end if
	    session("READYIN") = false
	    objDb.dbDisConnect
	else
	    gMessage =  objDb.strMessage
	end if
	Set objDb = Nothing
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
			strQuery = "DELETE ATC_PersonalInfo WHERE PersonID = " & varID
			if not objDB.runActionquery(strQuery) then
				rsItem.Bookmark = int(varBookmark)
				strDonot = strDonot & " " & Mid(rsItem("Fullname"), 1, 10) & " ..." & ","
			end if
	    next
	    if strDonot<>"" then
	      strDonot = Mid(strDonot, 2, len(strDonot)- 2)
	      gMessage = chr(34) & strDonot & chr(34) & " have been used."
	    else
	  	  gMessage = "Deleted successfully."
	    end if
	    rsItem.close()
	    session("rsItem") = empty
	    session("READYIN") = false
	    set rsComp = session("rsComp")
	    rsComp.Close()
	    session("rsComp") = empty
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
	strQuery = "SELECT PersonID, Firstname, Lastname, ISNULL(Firstname,'')+' '+ISNULL(Lastname,'') as Fullname, EmailAddress, " &_
				"Gender, isnull(CompanyID, 0) as companyID FROM ATC_PersonalInfo " &_
				"WHERE Usertype = 0 ORDER BY Fullname"
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
	  Else
		session("NumPagein") = 0
		Session("CurPagein") = 0
	  End if
	  objDb.CloseRec
	  'get company
	  objDb.recConnect(strConnect)
	  strQuery = "Select CompanyID, CompanyName from ATC_Companies Order by Companyname"
	  If objDb.openRec(strQuery) Then
		objDb.recDisConnect
		if not objDb.noRecord then
			set rsComp = objDb.rsElement.Clone
		else
			set rsComp = Server.CreateObject("ADODB.Recordset")
		end if
		set session("rsComp") = rsComp
		objDb.CloseRec
	  Else
		gMessage = objDb.strMessage
	  End if
	Else
	  gMessage = objDb.strMessage
	End if
	Set objDb = Nothing
Else
	set rsItem = session("rsItem")
	set rsComp = session("rsComp")
End if

if stract = "EDIT" and strNameItem1="" then
	varID = Request.Form("txthidden")
	rsItem.Bookmark = int(varID)
	strNameItem1 = rsItem("Firstname")
	strNameItem2 = rsItem("Lastname")
	strNameItem3 = rsItem("Emailaddress")
	strNameItem4 = rsItem("Gender")
	strNameItem5 = rsItem("CompanyID")
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
curpage = 10
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
<title>Atlas Industries - Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css" type="text/css">
<script language="javascript" src="../library/library.js"></script>
<script>
var fgcompemail = 0;
function _edit(varid){
	document.frminput.txthidden.value = varid;
	document.frminput.action = "contact.asp?act=EDIT"
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
 if (fg == false) alert("No country selected.")
 return(fg)
}

function _cancel(){
  	document.frminput.action = "contact.asp?act=CANCEL"
	document.frminput.target = "_self";
	document.frminput.submit();
}

function _save(){
var varid = "<%=varID%>";
	if(checkdata()==true) {
		document.frminput.txthidden.value = varid;
  		document.frminput.action = "contact.asp?act=SAVE"
		document.frminput.target = "_self";
		document.frminput.submit();
	}
}

function checkdata() {
	if(document.frminput.txtname3.value!="") {
		if((isemail(document.frminput.txtname3.value)==false)&&(fgcompemail==0)) {
			if (confirm("The specified email address does not appear to be valid, \n do you want to save it anyway?")==false) {
				document.frminput.txtname3.focus();
				return false;
			}else{
				fgcompemail = 1;
			}
		}
	}

	for(i=1;i<4;i++){
		tmp = eval("document.frminput.txtname" + i + ".value");
		tmp = alltrim(tmp);
		eval("document.frminput.txtname" + i + ".value = tmp");
		if(eval("document.frminput.txtname" + i + ".value")=="") {
			eval("document.frminput.txtname" + i + ".focus()");
			alert("Please enter a value");
			return false;
			break;
		}
	}

	if(document.frminput.lstcomp.options[document.frminput.lstcomp.selectedIndex].value=="") {
		alert("Please select a company");
		document.frminput.lstcomp.focus();
		return false;
	}
		
	return true;
}

function _add(){
	if(checkdata()==true) {
  		document.frminput.action = "contact.asp?act=ADD"
		document.frminput.target = "_self";
		document.frminput.submit();
	}
}

function _remove() {
  if (chkremove()==true) {
  	document.frminput.action = "contact.asp?act=DELETE"
	document.frminput.target = "_self";
	document.frminput.submit();
  }
}

function next() {
var curpage = <%=session("CurPagein")%>
var numpage = <%=session("NumPagein")%>
	if (curpage < numpage) {
		document.frminput.action = "contact.asp?navi=NEXT"
		document.frminput.target = "_self"
		document.frminput.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagein")%>
var numpage = <%=session("NumPagein")%>
	if (curpage > 1) {
		document.frminput.action = "contact.asp?navi=PREV";
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
		document.frminput.action = "contact.asp?Go=" + intpage;
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
            <td class="title" height="50" align="center" colspan="5">Contact</td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" nowrap> &nbsp;&nbsp;First Name&nbsp; </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname1" class="blue-normal" size="15" maxlength="15" style=" width:200" value="<%=strNameItem1%>">
            </td>
            <td class="blue-normal" height="60" align="left" rowspan="5" width="314"> 
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
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Last Name&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname2" class="blue-normal" size="15" maxlength="15" style="width:200" value="<%=strNameItem2%>">
            </td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Gender&nbsp; 
            </td>
            <td class="blue-normal" align="left" width="213"> 
               <input type="radio" name="optgender" value="1" <%if strNameItem4 = true or strNameItem4="" then%>checked<%end if%>>Male
               <input type="radio" name="optgender" value="0" <%if strNameItem4 = false then%>checked<%end if%>>Female
            </td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Email&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname3" class="blue-normal" size="15" maxlength="60" style="width:200" value="<%=strNameItem3%>">
            </td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Company&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <select name="lstcomp" style="width:200" class="blue-normal">
				<option value="">-- Select a company --</option>
<%	rsComp.MoveFirst
	Do until rsComp.EOF
		if strNameItem5 = rsComp("CompanyID") then strtmp = " selected " else strtmp = ""
%>
			<option value="<%=rsComp("CompanyID")%>" <%=strtmp%>><%=rsComp("CompanyName")%></option>
<%	rsComp.MoveNext
	Loop%>
              </select>
            </td>
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
                      <tr bgcolor="8CA0D1" align="left"> 
                        <td class="blue" bgcolor="8CA0D1" width="6%" align="right">&nbsp;</td>
                        <td class="blue" width="40%" align="center">Full Name</td>
                        <td class="blue" width="43%" align="center">Email</td>
                        <td class="blue" width="11%">&nbsp;</td>
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
