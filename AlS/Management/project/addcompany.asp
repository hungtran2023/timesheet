<!-- #include file = "../../class/CDatabase.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	dim strCompanyName,strAddress,strEmail,strCharCode,strNum,gMessage,varAct,strConnect,objDb

	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					
	
	gMessage = ""
	strCharCode = Request.QueryString("charcode")
	varAct = Request.QueryString("act")
	if varAct = "APPLY" then
	
		strCharCode = Request.Form("txtcharcode")
		strCompanyName = replace(trim(Request.Form("txtcompname")),"'","''")
		strAddress = replace(trim(Request.Form("txtcompadd")), "'", "''")
		strEmail = replace(trim(Request.Form("txtcompemail")),"'","''")

		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		
		if objDb.dbConnect(strConnect) then
			'Get the automatical numcode for a new Charcode
			strQuery = "SELECT max(numcode) + 1 as mCode from ATC_Companies"
			ret = objDb.runQuery(strQuery)
			if not ret then
				gMessage = objDb.strMessage
			else
				strNum = objDb.rsElement("mCode")
				if len(strNum) < 5 then strNum = "0" & strNum
				objDb.rsElement.Close()
			end if
			
			if gMessage = "" then
				strQuery1 = "INSERT INTO ATC_Companies(CompanyName, Website, EmailAddress, Charcode, Numcode) VALUES('" &_
						strCompanyName & "',"
				if strAddress="" then
					strQuery1=strQuery1	& "null,"
				else
					strQuery1=strQuery1	& "'" & strAddress & "'," 
				end if
				
				if strEmail="" then
					strQuery1=strQuery1	& "null,"
				else
					strQuery1=strQuery1	& "'" & strEmail & "'," 
				end if
				strQuery1 =strQuery1 & "'" & strCharCode & "', '" & strNum & "')"
				ret = objDb.runActionQuery(strQuery1)

				if not ret then gMessage = objDb.strMessage
				'end insert
				objDb.dbdisConnect
			end if
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
				window.opener.document.forms[0].action= scriptname + "?act=sc";
				window.opener.document.forms[0].submit();
				//-->
			</SCRIPT>
			<%	
		end if
	end if
%>
<html>
<head>
<title>Atlas Industries - Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>
<script>
var fgcompemail = 0;
function check() {
var tmp;
	if(document.sub.txtcompname) { //for input company
		tmp = document.sub.txtcompname.value;
		document.sub.txtcompname.value = alltrim(tmp);
		if (document.sub.txtcompname.value == "") {
			alert("Please enter value for company name.");
			document.sub.txtcompname.focus();
			return false;
		}

	}

	return true;
}

function add() {
	if (check()==true) {
		document.sub.action = "addcompany.asp?act=APPLY"
		document.sub.target = "_self"
		document.sub.submit();
	}	
}
</script>
</head>
<body bgcolor="#C0CAE6" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="302" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#003399">
  <tr> 
    <td valign="middle"> 
      <table width="300" border="0" cellspacing="0" cellpadding="0" align="center">
        <form name="sub" method="post">
        <%if gMessage<>"" then%>
		  <tr bgcolor="#C0CAE6">
            <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
          </tr>
        <%end if%>
          <tr bgcolor="C0CAE6" align="center"> 
            <td colspan="2" height="30" class="title">Client</td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="100%" colspan="2" class="blue-normal" height="26" > 
              <b>&nbsp;&nbsp;&nbsp;Company information</b>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="35%" class="blue-normal" height="26" align="right"> 
              Company name*&nbsp;
            </td>
            <td width="65%" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="text" name="txtcompname" class="blue-normal" value="<%=strCompanyName%>" size="20" style='HEIGHT: 22px; WIDTH: 160px' maxlength="100">
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="35%" class="blue-normal" height="26" align="right"> 
              Website&nbsp;
            </td>
            <td width="65%" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="text" name="txtcompadd" class="blue-normal" value="<%=strAddress%>" size="20" style='HEIGHT: 22px; WIDTH: 160px' maxlength="100">
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="35%" class="blue-normal" height="26" align="right"> 
              Email domain&nbsp;
            </td>
            <td width="65%" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="text" name="txtcompemail" class="blue-normal" value="<%=strEmail%>" size="20" style='HEIGHT: 22px; WIDTH: 160px' maxlength="60">
            </td>
          </tr>
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
          <input type="hidden" value="<%=strCharCode%>" name="txtcharcode">
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>