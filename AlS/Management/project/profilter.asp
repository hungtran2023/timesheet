<!-- #include file = "../../class/CDatabase.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If
	
	varAct = Request.QueryString("act")
	if varAct = "CLEAR" then Session("filter") = empty
	filter_criteria = ""
	if varAct = "APPLY" then		
		filter_criteria = "fgActivate =" & Request.Form("rdostatus")
		criteria = Request.Form("lsttype")
		if criteria <>"" then filter_criteria = filter_criteria & " AND ProjectType =" & criteria
		criteria = Request.Form("chkinhouse")
		if criteria <> "" then 
			filter_criteria = filter_criteria & " AND Projectkey2=7"
		else
			filter_criteria = filter_criteria & " AND ProjectKey2<>7"
		end if
		Session("filter") = filter_criteria
	end if
	if filter_criteria <> "" or varAct = "CLEAR" then
	
%>
<SCRIPT LANGUAGE=javascript>
<!--
	var tmp = window.opener.document.location;
	tmp = tmp.toString();
	var i2 = tmp.indexOf("?");
	if(i2==-1) { 
		i2 = tmp.length;
	}
	var i1 = tmp.lastIndexOf("/");
	scriptname = tmp.substring(0, i2);//tmp.substring(i1 + 1, i2);
	window.opener.document.forms[0].action = scriptname + "?filter=1";
	window.opener.document.forms[0].submit();
	//-->
</SCRIPT>

<%	
	end if
	'-------------------------------
	' get default
	'-------------------------------
'		if not isEmpty(session("Preferences")) then
'			arrPre = session("Preferences")
'			if arrPre(2, 0)<>"" then str
'			set arrPre = nothing
'		else
'			PageSize = PageSizeDefault
'		end if
	
	gMessage=""
	Set objDb = New clsDatabase
	strConnect = Application("g_strConnect")
	ret = objDb.dbConnect(strConnect)
	if ret then
		ret = objDb.runQuery("SELECT * FROM ATC_ProjectTypes")
		if not ret then
			gMessage = objDb.strMessage
		else
			strOut = "<select name='lsttype' class='blue-normal' style='HEIGHT: 22px; WIDTH: 130px'>"
			strOut = strOut & "<option value=''>All</option>"
			if not objDb.noRecord then
			  Do Until objDb.rsElement.EOF
			    strOut = strOut & "<option value='" & objDb.rsElement(0) & "'>" & showlabel(objDb.rsElement(1)) & "</option>"              
			    objDb.MoveNext
			  Loop
			end if
			strOut = strOut & "</select>"
			objDb.CloseRec
		end if
	else
		gMessage = objDb.strMessage
	end if	
	objDb.dbdisConnect
	set objDb = nothing
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script>
function apply() {
	document.filter.action = "profilter.asp?act=APPLY";
	document.filter.target = "_self";
	document.filter.submit();
}

function clear() {
	document.filter.action = "profilter.asp?act=CLEAR";
	document.filter.target = "_self";
	document.filter.submit();
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="filter" method="post">
<table width="252" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#003399" height="210">
  <tr> 
    <td valign="middle"> 
      <table width="250" border="0" cellspacing="0" cellpadding="0" align="center">
        <form name="form1" >
		  <tr bgcolor=<%if gMessage="" then%>"#C0CAE6"<%else%>"#E7EBF5"<%end if%>>
            <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
          </tr>        
          <tr bgcolor="C0CAE6" align="center"> 
            <td colspan="2" height="50" class="title">Filter Criteria</td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="99" class="blue-normal" height="26"> 
              <div align="right" class="c"> Project Status&nbsp; </div>
            </td>
            <td width="151" class="blue-normal" bgcolor="C0CAE6"> 
              <input type="radio" name="rdostatus" value="1" checked>Activated 
              <input type="radio" name="rdostatus" value="0">Deactivated</td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="99" class="blue-normal" height="26"> 
              <div align="right"> Project Type&nbsp; </div>
            </td>
            <td width="151" class="blue-normal" bgcolor="C0CAE6"> 
<%Response.Write strOut%>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="99" class="blue-normal" align="right" height="26">In-house 
              Project &nbsp; </td>
            <td width="151" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="checkbox" name="chkinhouse" value="1">
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td height="60" colspan="2"> 
              <table width="180" border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa">
                <tr> 
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue">
                      <a href="javascript:apply();" class="b" onMouseOver="self.status='Apply filter'; return true;" onMouseOut="self.status=''">Apply</a></td>
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center"> 
                    <a href="javascript:clear();" class="b" onMouseOver="self.status='Clear filter'; return true;" onMouseOut="self.status=''">Clear</a></td>
                  <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center">
					<a href="javascript:void(0);" class="b" onClick="window.close()" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
                </tr>
              </table>
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>