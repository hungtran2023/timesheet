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
			         "<td valign='top' width='32%' class='blue-normal'><a href='javascript:_edit(" & chr(34) & rsSrc.Bookmark &_
			         chr(34) & ");' class='c' OnMouseOver = 'self.status=&quot;Edit&quot; ; return true;' OnMouseOut =" &_
			         " 'self.status = &quot;&quot;'>&nbsp;" & Showlabel(rsSrc("CompanyName")) & "</a></td>" &_
			         "<td valign='top' width='33%' class='blue-normal'>&nbsp;" & rsSrc("EmailAddress") & "</td>" &_
			         "<td valign='top' width='8%' class='blue-normal'>&nbsp;" & rsSrc("Charcode") & "</td>" &_
			         "<td valign='top' width='10%' class='blue-normal'>&nbsp;" & rsSrc("Numcode") & "</td>" &_
			         "<td valign='top' width='11%' class='blue-normal' align='center'><input type='checkbox' name='chk' value='" &_
			         rsSrc("CompanyID") & "," & rsSrc.Bookmark & "'></td>" &_
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
fg = "0"
if Request.QueryString("txtfg")<>"" then fg = Request.QueryString("txtfg")
select case stract
case "SAVE"
	gMessage = ""
	strNameItem1 = Request.Form("txtname1")
	strNameItem2 = Request.Form("txtname2")
	strNameItem3 = Request.Form("txtname3")
	strNameItem4 = Request.Form("txtname4")
	strNameItem5 = Request.Form("txtname5")
	varID = Request.Form("txthidden")
	set rsItem = session("rsItem")
	if fg = "0" then
		rsItem.MoveFirst
		if Instr(strNameItem4, "'")>0 then
			tmp1 = "#" & strNameItem4 & "#"
		else
			tmp1 = "'" & strNameItem4 & "'"
		end if
		if Instr(strNameItem5, "'")>0 then
			tmp2 = "#" & strNameItem5 & "#"
		else
			tmp2 = "'" & strNameItem5 & "'"
		end if

		rsItem.Find "Charcode like " & tmp1,, adSearchForward
		If not rsItem.EOF then
			if rsItem.Bookmark<>int(varID) then
				gMessage = chr(34) & strNameItem4 & chr(34) & " has already been inputted."
				stract = "EDIT"
			else
				rsItem.MoveFirst
				rsItem.Find "Numcode like " & tmp2,, adSearchForward
				If not rsItem.EOF then
					if rsItem.Bookmark<>int(varID) then
						gMessage = chr(34) & strNameItem5 & chr(34) & " has already been inputted."
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
			strNameItem1 = replace(strNameItem1, "'", "''")
			strNameItem2 = replace(strNameItem2, "'", "''")
			strNameItem3 = replace(strNameItem3, "'", "''")
			strNameItem4 = replace(strNameItem4, "'", "''")
			strNameItem5 = replace(strNameItem5, "'", "''")
			if fg = "0" then
				strQuery = "UPDATE ATC_Companies SET CompanyName = '" & strNameItem1 & "', Address = '" & strNameItem2 &_
							"', EmailAddress = '" & strNameItem3 & "', Charcode = '" & strNameItem4 & "', Numcode = '" &_
							strNameItem5 & "' WHERE CompanyID = " & rsItem("CompanyID")
			else
				strQuery = "UPDATE ATC_Companies SET CompanyName = '" & strNameItem1 & "', Address = '" & strNameItem2 &_
							"', EmailAddress = '" & strNameItem3 & "' WHERE CompanyID = " & rsItem("CompanyID")
			end if
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
		    objDb.dbDisConnect
		else
		    gMessage =  objDb.strMessage
		end if
		Set objDb = Nothing
	end if
case "ADD"
	gMessage = ""
	strNameItem1 = Request.Form("txtname1")
	strNameItem2 = Request.Form("txtname2")
	strNameItem3 = Request.Form("txtname3")
	strNameItem4 = Request.Form("txtname4")
	strNameItem5 = Request.Form("txtname5")
	
	if not isEmpty(session("rsItem")) then
		set rsItem = session("rsItem")
		rsItem.MoveFirst
		if Instr(strNameItem4, "'")>0 then
			tmp1 = "#" & strNameItem4 & "#"
		else
			tmp1 = "'" & strNameItem4 & "'"
		end if
		if Instr(strNameItem5, "'")>0 then
			tmp2 = "#" & strNameItem5 & "#"
		else
			tmp2 = "'" & strNameItem5 & "'"
		end if

		rsItem.Find "Charcode like " & tmp1,, adSearchForward
		If not rsItem.EOF then
			if rsItem.Bookmark<>int(varID) then
				gMessage = chr(34) & strNameItem4 & chr(34) & " has already been inputted."
				'stract = "EDIT"
			else
				rsItem.MoveFirst
				rsItem.Find "Numcode like " & tmp2,, adSearchForward
				If not rsItem.EOF then
					if rsItem.Bookmark<>int(varID) then
						gMessage = chr(34) & strNameItem5 & chr(34) & " has already been inputted."
						'stract = "EDIT"
					end if
				end if
			end if
		end if		
	end if

	If gMessage = "" then
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			strNameItem1 = replace(strNameItem1, "'", "''")
			strNameItem2 = replace(strNameItem2, "'", "''")
			strNameItem3 = replace(strNameItem3, "'", "''")
			strNameItem4 = replace(strNameItem4, "'", "''")
			strNameItem5 = replace(strNameItem5, "'", "''")
			strQuery = "INSERT INTO ATC_Companies(CompanyName, Address, EmailAddress, Charcode, Numcode) VALUES('" &_
						strNameItem1 & "', '" & strNameItem2 & "', '" & strNameItem3 & "', '" & strNameItem4 &_
						"', '" & strNameItem5 & "')"
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
			strQuery = "DELETE ATC_Companies WHERE CompanyID = " & varID
			if not objDB.runActionquery(strQuery) then
				rsItem.Bookmark = int(varBookmark)
				strDonot = strDonot & " " & Mid(rsItem("Companyname"), 1, 10) & " ..." & ","
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
	strQuery = "select a.CompanyID, a.Companyname, a.Address, a.EmailAddress, a.charcode, a.Numcode, ISNULL(b.CompanyID, 0) fg FROM ATC_Companies a " &_
				"LEFT JOIN (Select distinct CompanyID From ATC_Projects) b ON a.CompanyID=b.CompanyID " &_
				"ORDER BY a.CompanyID DESC"
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
	Else
	  gMessage = objDb.strMessage	  
	End if
	Set objDb = Nothing
Else
	set rsItem = session("rsItem")
End if

if stract = "EDIT" and strNameItem1="" then
	varID = Request.Form("txthidden")
	rsItem.Bookmark = int(varID)
	strNameItem1 = rsItem("Companyname")
	strNameItem2 = rsItem("Address")
	strNameItem3 = rsItem("Emailaddress")
'	if rsItem("fg") = 0 then
'		fg = "0"
		strNameItem4 = rsItem("Charcode")
		strNameItem5 = rsItem("Numcode")
'	else
'		fg = "1"
'	end if
end if

fg = "0" 'temporary

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
curpage = 9
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
function _edit(varid){
	document.frminput.txthidden.value = varid;
	document.frminput.action = "company.asp?act=EDIT"
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
  	document.frminput.action = "company.asp?act=CANCEL"
	document.frminput.target = "_self";
	document.frminput.submit();
}

function _save(){
var varid = "<%=varID%>";
	if(checkdata()==true) {
		document.frminput.txthidden.value = varid;
  		document.frminput.action = "company.asp?act=SAVE"
		document.frminput.target = "_self";
		document.frminput.submit();
	}
}

function checkdata() {
	if(document.frminput.txtname3.value!="") {
		if(!emailCheck(document.frminput.txtname3.value)) {
			document.frminput.txtname3.focus();
			return false;
		}
	}

	if(document.frminput.txtname5.value!="") {
		var tmp = document.frminput.txtname5.value;
		if(tmp.indexOf("0") != 0) {
			alert("The first figure must be zero.");
			document.frminput.txtname5.focus();
			return false;
		}
		else {
			if(tmp.indexOf("0", 1) == 1) {
			alert("The second figure must not be zero.");
			document.frminput.txtname5.focus();
			return false;
		}
		}

	}
	
	for(i=1;i<6;i++){
		tmp = eval("document.frminput.txtname" + i + ".value");
		tmp = alltrim(tmp)
		eval("document.frminput.txtname" + i + ".value = tmp");
		if(eval("document.frminput.txtname" + i + ".value")=="") {
			eval("document.frminput.txtname" + i + ".focus()");
			alert("Please enter a value");
			return false;
			break;
		}
	}
	
	return true;
}

function _add(){
	if(checkdata()==true) {
  		document.frminput.action = "company.asp?act=ADD"
		document.frminput.target = "_self";
		document.frminput.submit();
	}
}

function _remove() {
  if (chkremove()==true) {
  	document.frminput.action = "company.asp?act=DELETE"
	document.frminput.target = "_self";
	document.frminput.submit();
  }
}

function next() {
var curpage = <%=session("CurPagein")%>
var numpage = <%=session("NumPagein")%>
	if (curpage < numpage) {
		document.frminput.action = "company.asp?navi=NEXT"
		document.frminput.target = "_self"
		document.frminput.submit();
	}
}

function prev() {
var curpage = <%=session("CurPagein")%>
var numpage = <%=session("NumPagein")%>
	if (curpage > 1) {
		document.frminput.action = "company.asp?navi=PREV";
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
		document.frminput.action = "company.asp?Go=" + intpage;
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
            <td class="title" height="50" align="center" colspan="5">Company</td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" nowrap> &nbsp;&nbsp;Company name&nbsp; </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname1" class="blue-normal" size="15" maxlength="100" style=" width:200" value="<%=strNameItem1%>">
            </td>
            <td class="blue-normal" height="60" align="left" rowspan="3" width="314"> 
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
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Address&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname2" class="blue-normal" size="15" maxlength="100" style=" width:200" value="<%=strNameItem2%>">
            </td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Email&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname3" class="blue-normal" size="15" maxlength="60" style=" width:200" value="<%=strNameItem3%>">
            </td>
          </tr>
<%if fg = "0" then%>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Char Code&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname4" class="blue-normal" size="15" maxlength="3" style=" width:50" value="<%=strNameItem4%>">
            </td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;Num Code&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              <input type="text" name="txtname5" class="blue-normal" size="15" maxlength="5" style=" width:50" value="<%=strNameItem5%>">
            </td>
          </tr>
<%else%>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;&nbsp; 
            </td>
            <td class="blue" align="left" width="213">&nbsp;</td>
          </tr>
          <tr align="center"> 
            <td class="blue-normal" align="right" width="82" nowrap> &nbsp;&nbsp;&nbsp; 
            </td>
            <td class="blue" align="left" width="213"> 
              &nbsp;
            </td>
          </tr>
<%end if%>
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
                        <td class="blue" bgcolor="8CA0D1" width="6%" align="right">&nbsp;</td>
                        <td class="blue" width="32%" align="center">Company</td>
                        <td class="blue" width="33%" align="center">Email</td>
                        <td class="blue" width="8%" align="center">Char Code</td>
                        <td class="blue" width="10%" align="center">Num Code</td>
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
<input type="hidden" name="txtfg" value="<%=fg%>">
</form>
</body>
</html>
