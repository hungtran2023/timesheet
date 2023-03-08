<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/library.asp"-->
<!-- #include file = "fileupload_class.inc"-->
<%
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
'	If checkSession(session("Inhouse")) = False Then
'		Response.Redirect("message.htm")
'	End If
	
	gMessage = ""
	varAct = Request.QueryString("act")
	if varAct = "APPLY" then
		fgComplete = false
		
		'Upload logo image
		set fo = new fileupload
		Response.Write("commnet: " & fo.FormInput("txtCommnent"))
		Response.End
		With fo
			.UploadDirectory = Server.MapPath("../images")
			.UploadSize = 1000000
			.allowoverwrite = true
			.extensions = array(".jpg", ".gif")
			.upload
			
			if .uploadsuccessful then
				filepath = .absolutepathtouploadedfile
				'Write the filename into database
				Set objDb = New clsDatabase
				strConnect = Application("g_strConnect")
				ret = objDb.dbConnect(strConnect)
				if ret then
					filename = Right(filepath ,Len(filepath)-InstrRev(filepath, "\"))
					'strQuery = "UPDATE ATC_CompanyProfile SET logo = '" & filename & "' WHERE CompanyID =" & session("Inhouse")
'					ret = objDb.runActionQuery(strQuery)
					if not ret then gMessage = objDb.strMessage
					objDb.dbdisConnect
				else
					gMessage = objDb.strMessage
				end if		
				set objDb = nothing
				if gMessage="" then fgComplete = true
			else
				gMessage = "Can't upload this file."
				fgComplete = false
			end if
		end With
		set fo = nothing
		if fgComplete then
			%>
			<SCRIPT LANGUAGE=javascript>
			<!--
/*				window.close();
				var tmp = window.opener.document.location;
				tmp = tmp.toString();
				var i2 = tmp.indexOf("?");
				if(i2==-1) { 
					i2 = tmp.length;
				}
				var i1 = tmp.lastIndexOf("/");
				scriptname = tmp.substring(0, i2);//tmp.substring(i1 + 1, i2);
				window.opener.document.forms[0].action= scriptname;// + "?logo=1";
				window.opener.document.forms[0].submit();
				//-->*/
			</SCRIPT>
			<%

		end if
	end if
%>
<html>
<head>
<title>Atlas Industries Time Sheet System</title>
<link rel="stylesheet" href="../timesheet.css" type="text/css">
<script language="javascript" src="../library/library.js"></script>
<script>
function check() {
	if (document.sub.blob.value == "") {
		alert("Please enter value for this field.");
		document.sub.blob.focus();
		return false;
	}
	return true;
}
function add() {
	if(check()==true) {
		document.sub.action = "upload.asp?act=APPLY"
		document.sub.target = "_self"
		document.sub.submit();
	}
}
</script>
</head>
<body style="margin:0;color:#000000; background-color:#FFFFFF;">
<form name="sub" method="post"  enctype="multipart/form-data">
<table style="width:256px; height:158px; padding:0; border:none; background-color:#003399">
  <tr> 
    <td valign="middle"> 
      <table width="263" border="0" cellspacing="0" cellpadding="0" align="center">
      
        
		  <tr bgcolor="<%if gMessage="" then%>#C0CAE6<%else%>#E7EBF5<%end if%>">
            <td class="red" colspan="2" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
          </tr>
          <tr bgcolor="C0CAE6" align="center"> 
            <td colspan="2" height="50" class="title">Upload Logo</td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td width="15%" class="blue-normal" height="26" align="right"> Type&nbsp;</td>
            <td width="85%" bgcolor="C0CAE6">
            <input type="file" name="blob" class="blue-normal" size="13" style='HEIGHT: 22px; WIDTH: 200px'/>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td class="blue-normal" height="26" align="right"> Comment&nbsp;</td>
            <td bgcolor="C0CAE6">
                <input type="text" name="txtCommnent" id="txtCommnent" class="blue-normal" style="height:22px; width:200px;"/>
            </td>
          </tr>
          <tr bgcolor="C0CAE6"> 
            <td height="60" colspan="2"> 
              <table width="120" border="0" cellspacing="5" cellpadding="0" align="center" height="20">
                <tr> 
                  <td bgcolor="8CA0D1" onmouseover="this.style.backgroundColor='7791D1';" onmouseout="this.style.backgroundColor='8CA0D1';" height="20" align="center" class="blue"> 
                      <a href="javascript:add();" class="b" onmouseover="self.status='Upload'; return true;" onmouseout="self.status=''">Upload</a>
                  </td>
                  <td bgcolor="8CA0D1" onmouseover="this.style.backgroundColor='7791D1';" onmouseout="this.style.backgroundColor='8CA0D1';" class="blue" height="20" align="center" >
                  <a href="javascript:window.close();" class="b" onmouseover="self.status='Close window'; return true;" onmouseout="self.status=''">Close</a></td>
                </tr>
              </table>
            </td>
          </tr>
        
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>

