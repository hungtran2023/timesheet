<!-- #include file = "../../class/CDatabase.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					
	
	gMessage = ""
	varKind = Request.QueryString("type")
	if varKind = "" then varKind = Request.Form("txtkind")
	select case varKind
	case 1
		pagetitle = "job title"
		firstlabel = "Job title"
		secondlabel = ""
	case 2
		pagetitle = "country"
		firstlabel = "Country"
		secondlabel = "Nationality"
	case 3
		pagetitle = "working hour"
		firstlabel = "Office hour"
		secondlabel = "Description"
	case 4
		pagetitle = "department"
		firstlabel = "Department"
		secondlabel = ""
	end select

	varAct = Request.QueryString("act")
	if varAct = "APPLY" then
		varKind = Request.Form("txtkind")
		varinfo1 = trim(Request.Form("txtname1"))
		varinfo1 = replace(varinfo1, "'", "''")
		varinfo1 = replace(varinfo1, chr(34), "''")
		if varKind = 2 or varKind = 3 then
			varinfo2 = trim(Request.Form("txtname2"))
			varinfo2 = replace(varinfo2, "'", "''")
			varinfo2 = replace(varinfo2, chr(34), "''")
		end if
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			select case varKind			
			case 1
				strQuery = "INSERT INTO ATC_JobTitle(JobTitle) " &_
							"VALUES('" & varinfo1 & "')"
			case 2
				strQuery = "INSERT INTO ATC_Countries(CountryName, Nationality) " &_
							"VALUES('" & varinfo1 & "', '" & varinfo2 & "')"
			case 3
				strQuery = "INSERT INTO ATC_WorkingHours(Hours, Description) " &_
							"VALUES('" & varinfo1 & "', '" & varinfo2 & "')"
			case 4
				strQuery = "INSERT INTO ATC_Department(Department) " &_
							"VALUES('" & varinfo1 & "')"
			end select
			ret = objDb.runActionQuery(strQuery)
			if not ret then gMessage = objDb.strMessage
			objDb.dbdisConnect
		else
			gMessage = objDb.strMessage
		end if		
		set objDb = nothing
		if gMessage = "" then	
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
				window.opener.document.forms[0].action= scriptname + "?addmore=1";
				window.opener.document.forms[0].submit();
				//-->
			</SCRIPT>
			<%	
		end if
	end if
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>
<script>
function check() {
	var tmp =document.sub.txtname1.value;
	document.sub.txtname1.value = alltrim(tmp);
	if (document.sub.txtname1.value == "") {
		alert("Please enter value for this field.");
		document.sub.txtname1.focus();
		return false;
	}
	if(document.sub.txtname2){
		var tmp =document.sub.txtname2.value;
		document.sub.txtname2.value = alltrim(tmp);
		if (document.sub.txtname2.value == "") {
			alert("Please enter value for this field.");
			document.sub.txtname2.focus();
			return false;
		}
	}
	return true;
}

function add() {
	if (check()==true) {
		document.sub.action = "addmore.asp?act=APPLY"
		document.sub.target = "_self"
		document.sub.submit();
	}	
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="265" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#003399" height="184">
  <tr> 
    <td valign="top"> 
      <table width="263" border="0" cellspacing="0" cellpadding="0" align="center">
        <form name="sub" method="post">
		  <tr bgcolor=<%if gMessage="" then%>"#C0CAE6"<%else%>"#E7EBF5"<%end if%>>
            <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
          </tr>
          <tr bgcolor="C0CAE6" align="center"> 
            <td colspan="2" height="50" class="title">Add <%=pagetitle%></td>
          </tr>

          <tr bgcolor="C0CAE6"> 
            <td width="25%" class="blue-normal" height="26" align="right"> 
              <%=firstlabel%>&nbsp;
            </td>
            <td width="75%" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="text" name="txtname1" maxlength="30" class="blue-normal" size="18" style='HEIGHT: 22px; WIDTH: 180px'>
            </td>
          </tr>
<%if secondlabel<>"" then%>
          <tr bgcolor="C0CAE6"> 
            <td width="25%" class="blue-normal" height="26" align="right"> 
              <%=secondlabel%>&nbsp;
            </td>
            <td width="75%" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="text" name="txtname2" maxlength="30" class="blue-normal" size="18" style='HEIGHT: 22px; WIDTH: 180px'>
            </td>
          </tr>          
<%end if%>
          <tr bgcolor="C0CAE6"> 
            <td height="60" colspan="2"> 
              <table width="120" border="0" cellspacing="5" cellpadding="0" align="center" height="20">
                <tr> 
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                      <a href="javascript:add();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Add</a>
                  </td>
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center" >
                  <a href="javascript:window.close();" class="b" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
                </tr>
              </table>
            </td>
          </tr>
          <input type="hidden" name="txtkind" value="<%=varKind%>">
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>