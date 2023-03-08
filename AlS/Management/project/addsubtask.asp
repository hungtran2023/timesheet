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
	
	gMessage = ""
'main procedure
	varAct = Request.QueryString("act")
	set rsTask = session("rsTaskCache")
	proID = Request.QueryString("proid")
	if varAct = "APPLY" then
		varTaskParent = Request.Form("lsttask")
		varTask = Request.Form("txtname")
		varTask = replace(varTask, "'", "''")
		varTask = replace(varTask, chr(34), "''")
		strChainID = Mid(varTaskParent, InStr(varTaskParent,"@") + 1,len(varTaskParent))
		varTaskParent = Mid(varTaskParent, 1, InStr(varTaskParent,"@") - 1)
		strChainID = strChainID & varTaskParent & ","
		Set objDb = New clsDatabase
		strConnect = Application("g_strConnect")
		ret = objDb.dbConnect(strConnect)
		if ret then
			'testing for this task have assignment or not
			strQuery = "SELECT count(*) FROM ATC_Assignments WHERE SubtaskID = " & varTaskParent		
			if objDb.runQuery(strQuery) then
				if objDb.rsElement(0)>0 then
					rsTask.MoveFirst
					rsTask.Find "sID = " & varTaskParent, , adSearchForward
					gMessage = "Task '" & rsTask("sName") & "'  was assigned."
					rsTask.MoveFirst
				end if
			else
				gMessage = objDb.strMessage
			end if
			
			if gMessage = "" then
				strQuery = "INSERT INTO ATC_Tasks(ProjectID, SubtaskName, TaskID, ChainID, OwnerID) " &_
							"VALUES('" & proID & "', '" & varTask & "', " & varTaskParent & ", '" & strChainID & "', " & session("USERID") & ")"
				ret = objDb.runActionQuery(strQuery)
				if not ret then gMessage = objDb.strMessage
			end if
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
				window.opener.document.forms[0].action= scriptname + "?addsub=1";
				window.opener.document.forms[0].submit();
				//-->
			</SCRIPT>
			<%	
		end if
	end if
	if varAct <> "APPLY" or gMessage<>"" then
		taskname = Request.Form("txtname")
		Dim arrRs(4)
		rsTask.MoveFirst
		strListTask = makeList(rsTask, 1) '1: for add subtask(owner or righton and no assigned), 2: for assignment (leaf)
	end if
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script>
function check() {
	if (document.sub.txtname.value == "") {
		alert("Please enter value for this field.");
		document.sub.txtname.focus();
		return false;
	}
	var tmp = document.sub.lsttask.selectedIndex
	if (tmp == -1) {
			alert("Please choose an item.");
			document.sub.lsttask.focus();
			return false;
	}
	if (document.sub.lsttask.options[tmp].value == "") {
			alert("This item can't be chosen.");
			document.sub.lsttask.focus();
			return false;
	}
	return true;
}

function add() {
var strproid = "<%=proID%>";
	if (check()==true) {
		document.sub.action = "addsubtask.asp?act=APPLY&proid=" + strproid;
		document.sub.target = "_self"
		document.sub.submit();
	}	
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="265" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" bgcolor="#003399" height="184">
  <tr> 
    <td valign="middle"> 
      <table width="263" border="0" cellspacing="0" cellpadding="0" align="center">
        <form name="sub" method="post">
		  <tr bgcolor=<%if gMessage="" then%>"#C0CAE6"<%else%>"#E7EBF5"<%end if%>>
            <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
          </tr>
          <tr bgcolor="C0CAE6" align="center"> 
            <td colspan="2" height="50" class="title">Add Sub-task</td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="25%" class="blue-normal" height="26" align="right"> 
              Name&nbsp;
            </td>
            <td width="75%" class="text-blue01" bgcolor="C0CAE6"> 
              <input type="text" name="txtname" maxlength="50" class="blue-normal" size="20" style='HEIGHT: 22px; WIDTH: 180px' value="<%=taskname%>">
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="25%" class="blue-normal" height="26" align="right"> 
              Sub-task of&nbsp;
            </td>
            <td width="75%" class="text-blue01" bgcolor="C0CAE6"> 
<%
	Response.Write (strListTask)
%>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td height="60" colspan="2"> 
              <table width="120" border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa">
                <tr> 
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                      <a href="javascript:add();" class="b" onMouseOver="self.status='Submit'; return true;" onMouseOut="self.status=''">Add</a>
                  </td>
                  <td bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center" >
                  <a href="javascript:void(0);" class="b" onClick="window.close();" onMouseOver="self.status='Close window'; return true;" onMouseOut="self.status=''">Close</a></td>
                </tr>
              </table>
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>