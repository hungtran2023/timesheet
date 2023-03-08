<%@ Language=VBScript %>
<!-- #include file = "class/CDatabase.asp"-->
<!-- #include file="class/clsSHA-1.asp" -->
<!-- #include file = "inc/createtemplate.inc"-->
<!-- #include file = "inc/getmenu.asp"-->
<%

Response.Expires = - 1441
Response.Buffer = true

Dim SessionSharing
Set SessionSharing = server.CreateObject("SessionMgr.Session2")

strURL=request.querystring("ReturnUrl")
if strURL="" then strURL=request.form("txtReturnUrl")

If Request.Form("txtusername") <> "" Then
   	Dim objDatabase, objSHA1
	Dim strConnect, strDigest, strPassword, strError

	strConnect = Application("g_strConnect")

' Connect to SQL database
	Set objDatabase = New clsDatabase

	If objDatabase.dbConnect(strConnect) Then
		strPassword = Request.Form("txtpwd")

		Set objSHA1 = New clsSHA1
		strDigest = ObjSHA1.SecureHash(strPassword)
'DA39A3EE5E6B4B0D3255BFEF95601890AFD80709
		'Response.Write strDigest
		'Response.End

		Set ObjSHA1 = Nothing

		strQuery = "SELECT a.UserID, a.UserName, b.UserType FROM ATC_Users a INNER JOIN ATC_PersonalInfo b ON a.UserID = b.PersonID " &_
					"WHERE b.fgDelete = 0 AND a.UserName = '" & replace(trim(Request.Form("txtusername")),"'","""") & "'"
'Response.Write strQuery
		If (objDatabase.runQuery(strQuery)) Then
			If objDatabase.noRecord = False Then
			    Session("UserType")=objDatabase.getColumn_by_name("UserType")
'if trim(Request.Form("txtusername")) = "chintu"  Then

				If (objDatabase.runQuery("SELECT UserID, UserName, Password, fgChangePass FROM ATC_Users WHERE UserName = '" & replace(trim(Request.Form("txtusername")),"'","""") & "' AND Password = '" & strDigest & "'")) Then
				'If (objDatabase.runQuery("SELECT UserID, UserName, Password FROM ATC_Users WHERE UserName = '" & replace(trim(Request.Form("txtusername")),"'","""") & "'")) Then
					If objDatabase.noRecord = False Then
							'strError = "The system is being upgraded. Please try to log-in again after 11:00 AM."
						    blnChangePass = objDatabase.getColumn_by_name("fgChangePass")
						    if not blnChangePass then
						        strURL="tools/confirm.asp?id=" & objDatabase.getColumn_by_name("UserID")
						        Response.Redirect(strURL)
						    end if

						'end if
						'if fgChangePass=0 then

						'else
							Session("USERID") = objDatabase.getColumn_by_name("UserID")
                            SessionSharing("USERID") =  Session("USERID")
							Session("USERNAME") = objDatabase.getColumn_by_name("UserName")

							session("GroupManager")=-1
							If (objDatabase.runQuery("SELECT UserID,GroupID FROM ATC_UserGroup WHERE GroupID=33 AND UserID=" & Session("USERID") )) Then
								If not objDatabase.noRecord then
									session("GroupManager") = objDatabase.getColumn_by_name("GroupID")
								end if
							End if

							'Checking for daily Timesheet
							'session("GroupManager") = objDatabase.getColumn_by_name("GroupManager")
							'get companyID
							session("InHouse") = 0
							If (objDatabase.runQuery("SELECT CompanyID FROM ATC_CompanyProfile")) Then
								If not objDatabase.noRecord then
									session("InHouse") = objDatabase.getColumn_by_name("CompanyID")
								end if
							End if
							objDatabase.dbDisConnect
							Set objDatabase = Nothing

							'Get list of right
							session("Righton") = empty
							If isEmpty(session("Righton")) then
								getRight = getarrRight(session("USERID"))
								if isArray(getRight) then session("Righton") = getRight
							End if

							'Get list of preferences
							session("Preferences") = empty
							If isEmpty(session("Preferences")) then
								getPre = getarrPreference(session("USERID"))
								if isArray(getPre) then session("Preferences") = getPre
							End if

							'Make list of menu
							session("Menu") = empty
							If isEmpty(session("Menu")) then
								getRes = getarrMenu(session("USERID"))
								session("Menu") = getRes
							End if

							Response.Clear

							if strURL="" then
								if isEmpty(session("Preferences")) then
									strURL = "welcome.asp"
								else
								
										arrPre = session("Preferences")
										if arrPre(0, 0)<>"" and arrPre(0, 0)<>"true" then strURL = arrPre(0, 0)
										if strURL = "" then strURL = "welcome.asp"
										set arrPre = nothing
								
								end if
							end if
							Response.Redirect(strURL)

					Else
						strError = "Password is invalid. Please try again."
					End If
				Else
					Response.Redirect("message.htm?strerror=" & objDatabase.strMessage)
				End If
'else
'    strError = "Timesheet is in maintenance process, please come back after 15:00"
'    Response.Redirect("Upgrade.htm")
'end if
			Else
				strError = "Username is invalid. <br>Timesheet username is the same with your login PC."
			End If

		Else
			Response.Redirect("message.htm?strerror=" & objDatabase.strMessage)
		End If
	Else
		Response.Redirect("message.htm?strerror=objDatabase.strMessage")

		objDatabase.dbDisConnect
		Set objDatabase = Nothing
   End If
End If

'Response.Cookies("templatepath") = "templates/template1/"

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplate("&nbsp;", "&nbsp;", arrPageTemplate, "templates/template1/")

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Atlas Information System</title>
<link rel="stylesheet" href="timesheet.css" type="text/css">

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background: #ffffff;
}
.lable {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #495860;
	font-weight: bold;
}
.clock {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #575757;
	font-weight: normal;
}
.new_a:active
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    FONT-FAMILY: Arial, Helvetica, sans-serif;
    TEXT-DECORATION: none
}
.new_a:link
{
	FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    FONT-FAMILY: Arial;
    TEXT-DECORATION: none

}
.new_a:visited
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    FONT-FAMILY: Arial;
    TEXT-DECORATION: none
}
.new_a:hover
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    BACKGROUND-REPEAT: repeat;
    FONT-FAMILY: Arial;
    TEXT-DECORATION: underline
}
-->

.popup{
	position: absolute;
	background: #E7EBF5;
	border: 1px solid #032e9b;
	border-style:outset;
	z-index: 10000;
}

#background{
	position: absolute;
	background: gray;
	left: 0px;
	top: 0px;
}
.lable {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #ffffff;
	font-weight: bold;
}

.new_a:active
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    FONT-FAMILY: Arial, Helvetica, sans-serif;
    TEXT-DECORATION: none
}
.new_a:link
{
	FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    FONT-FAMILY: Arial;
    TEXT-DECORATION: none

}
.new_a:visited
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    FONT-FAMILY: Arial;
    TEXT-DECORATION: none
}
.new_a:hover
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 10pt;
    COLOR: #ffffff;
    BACKGROUND-REPEAT: repeat;
    FONT-FAMILY: Arial;
    TEXT-DECORATION: underline
}

</style>

<script type="text/javascript" src="library/library.js"></script>


<link href="jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>

<script type="text/javascript" src="jQuery/jquery.min.js"></script>
<script type="text/javascript" src="jQuery/jquery-ui.min.js"></script>
<link href="jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<script ID="clientEventHandlersJS" language="Javascript">
<!--

var ns, ie;

ns = (document.layers)? true:false
ie = (document.all)? true:false

function checkin()
{
	debugger;
	if (isempty(window.document.frmlogin.txtusername.value))
	{
		alert("Please enter user name.");
		window.document.frmlogin.txtusername.focus();
	}
	else
	{
		window.document.frmlogin.action ="initial.asp";
		window.document.frmlogin.target = "_self";
		//window.document.frmlogin.submit();
	}
}

/*function document_onkeypress()
{
var keycode = event.keyCode;
	if (keycode == 13)
	{
		event.keyCode = 0;
		checkin();
	}
}*/

function window_onload()
{
var strError = "<%=strError%>";
	if (strError == "Invalid password!")
		window.document.frmlogin.txtpwd.focus();
	else{
		window.document.frmlogin.txtusername.focus();

//WorldClock()}
}
//-->

zone=0;

isitlocal=true;
ampm='';

function updateclock(z){
	zone=z.options[z.selectedIndex].value;
	isitlocal=(z.options[0].selected)?true:false;
}

function WorldClock(){
	//London
	zone = 0;
	//zone = 1;
	isitlocal = false;

	now=new Date();
	ofst=now.getTimezoneOffset()/60;
	secs=now.getSeconds();
	sec=-1.57+Math.PI*secs/30;
	mins=now.getMinutes();
	min=-1.57+Math.PI*mins/30;
	hr=(isitlocal)?now.getHours():(now.getHours() + parseInt(ofst)) + parseInt(zone);
	hrs=-1.575+Math.PI*hr/6+Math.PI*parseInt(now.getMinutes())/360;
	if (hr < 0) hr+=24;
	if (hr > 23) hr-=24;
	ampm = (hr > 11)?"PM":"AM";
	statusampm = ampm.toLowerCase();

	hr2 = hr;
	if (hr2 == 0) hr2=12;
	(hr2 < 13)?hr2:hr2 %= 12;
	if (hr2<10) hr2="0"+hr2

	var finaltime=hr2+':'+((mins < 10)?"0"+mins:mins)+' '+statusampm + ' London';
	if (document.all)
		worldclockLondon.innerHTML=finaltime
	else if (document.getElementById)
		document.getElementById("worldclockLondon").innerHTML=finaltime
	else if (document.layers){

		document.worldclockns.document.worldclockLondon.document.write(finaltime)
		document.worldclockns.document.worldclockLondon.document.close()
	}

	//Dubai
	zone = 4;
	isitlocal = false;

	now=new Date();
	ofst=now.getTimezoneOffset()/60;
	secs=now.getSeconds();
	sec=-1.57+Math.PI*secs/30;
	mins=now.getMinutes();
	min=-1.57+Math.PI*mins/30;
	hr=(isitlocal)?now.getHours():(now.getHours() + parseInt(ofst)) + parseInt(zone);
	hrs=-1.575+Math.PI*hr/6+Math.PI*parseInt(now.getMinutes())/360;
	if (hr < 0) hr+=24;
	if (hr > 23) hr-=24;
	ampm = (hr > 11)?"PM":"AM";
	statusampm = ampm.toLowerCase();

	hr2 = hr;
	if (hr2 == 0) hr2=12;
	(hr2 < 13)?hr2:hr2 %= 12;
	if (hr2<10) hr2="0"+hr2

	var finaltime=hr2+':'+((mins < 10)?"0"+mins:mins)+' '+statusampm + ' Dubai';

	if (document.all)
	worldclockDubai.innerHTML=finaltime
	else if (document.getElementById)
	document.getElementById("worldclockDubai").innerHTML=finaltime
	else if (document.layers){
	document.worldclockns.document.worldclockDubai.document.write(finaltime)
	document.worldclockns.document.worldclockDubai.document.close()
	}

	//Saigon
	zone = 7;
	isitlocal = false;

	now=new Date();
	ofst=now.getTimezoneOffset()/60;
	secs=now.getSeconds();
	sec=-1.57+Math.PI*secs/30;
	mins=now.getMinutes();
	min=-1.57+Math.PI*mins/30;
	hr=(isitlocal)?now.getHours():(now.getHours() + parseInt(ofst)) + parseInt(zone);
	hrs=-1.575+Math.PI*hr/6+Math.PI*parseInt(now.getMinutes())/360;
	if (hr < 0) hr+=24;
	if (hr > 23) hr-=24;
	ampm = (hr > 11)?"PM":"AM";
	statusampm = ampm.toLowerCase();

	hr2 = hr;
	if (hr2 == 0) hr2=12;
	(hr2 < 13)?hr2:hr2 %= 12;
	if (hr2<10) hr2="0"+hr2

	//var finaltime=hr2+':'+((mins < 10)?"0"+mins:mins)+':'+((secs < 10)?"0"+secs:secs)+' '+statusampm;
	var finaltime=hr2+':'+((mins < 10)?"0"+mins:mins)+' '+statusampm + ' Saigon';

	if (document.all)
	worldclockSaiGon.innerHTML=finaltime
	else if (document.getElementById)
	document.getElementById("worldclockSaiGon").innerHTML=finaltime
	else if (document.layers){
	document.worldclockns.document.worldclockSaiGon.document.write(finaltime)
	document.worldclockns.document.worldclockSaiGon.document.close()
	}

	//Sedney
	zone = 11;
	//zone = 10;
	isitlocal = false;

	now=new Date();
	ofst=now.getTimezoneOffset()/60;
	secs=now.getSeconds();
	sec=-1.57+Math.PI*secs/30;
	mins=now.getMinutes();
	min=-1.57+Math.PI*mins/30;
	hr=(isitlocal)?now.getHours():(now.getHours() + parseInt(ofst)) + parseInt(zone);
	hrs=-1.575+Math.PI*hr/6+Math.PI*parseInt(now.getMinutes())/360;
	if (hr < 0) hr+=24;
	if (hr > 23) hr-=24;
	ampm = (hr > 11)?"PM":"AM";
	statusampm = ampm.toLowerCase();

	hr2 = hr;
	if (hr2 == 0) hr2=12;
	(hr2 < 13)?hr2:hr2 %= 12;
	if (hr2<10) hr2="0"+hr2

	var finaltime= hr2+':'+((mins < 10)?"0"+mins:mins)+' '+statusampm + ' Melbourne';

	if (document.all)
	worldclockSydney.innerHTML=finaltime
	else if (document.getElementById)
	document.getElementById("worldclockSydney").innerHTML=finaltime
	else if (document.layers){
	document.worldclockns.document.worldclockSydney.document.write(finaltime)
	document.worldclockns.document.worldclockSydney.document.close()
	}
	
	//Hong Kong
	zone = 8;
	isitlocal = false;

	now=new Date();
	ofst=now.getTimezoneOffset()/60;
	secs=now.getSeconds();
	sec=-1.57+Math.PI*secs/30;
	mins=now.getMinutes();
	min=-1.57+Math.PI*mins/30;
	hr=(isitlocal)?now.getHours():(now.getHours() + parseInt(ofst)) + parseInt(zone);
	hrs=-1.575+Math.PI*hr/6+Math.PI*parseInt(now.getMinutes())/360;
	if (hr < 0) hr+=24;
	if (hr > 23) hr-=24;
	ampm = (hr > 11)?"PM":"AM";
	statusampm = ampm.toLowerCase();

	hr2 = hr;
	if (hr2 == 0) hr2=12;
	(hr2 < 13)?hr2:hr2 %= 12;
	if (hr2<10) hr2="0"+hr2

	//var finaltime=hr2+':'+((mins < 10)?"0"+mins:mins)+':'+((secs < 10)?"0"+secs:secs)+' '+statusampm;
	var finaltime=hr2+':'+((mins < 10)?"0"+mins:mins)+' '+statusampm + ' Hong Kong';

	if (document.all)
	worldclockHongKong.innerHTML=finaltime
	else if (document.getElementById)
	document.getElementById("worldclockHongKong").innerHTML=finaltime
	else if (document.layers){
	document.worldclockns.document.worldclockHongKong.document.write(finaltime)
	document.worldclockns.document.worldclockHongKong.document.close()
	}


	setTimeout('WorldClock()',1000);
}

window.onload=WorldClock
</script>

<script type="text/javascript">

    $(document).ready(function() {
        openUpload();

        $("#background").click(function(e) {
        closePopup();
        e.preventDefault();
         });

         $("#closeButton").click(function(e) {
        closePopup();
        e.preventDefault();
         });

    })

function openUpload() {
    var dheight = getBrowserHeight();
    var dwidth = getBrowserWidth();

    $("#background").width(dwidth).height(dheight);
    $("#background").fadeTo("slow", 0.8);

    var $divUpload = $("#divUpload");
    $divUpload.css("top", (dheight - $divUpload.height()) / 2);
    $divUpload.css("left", (dwidth - $divUpload.width()) / 2);

    $divUpload.fadeIn();
}

function closePopup() {
    $("#background").fadeOut();
    $("#divUpload").hide();
}

function getBrowserWidth() {
    if (window.innerWidth) {
        return window.innerWidth;
    }
    else if (document.documentElement && document.documentElement.clientWidth != 0) {
        return document.documentElement.clientWidth;
    }
    else if (document.body) {
        return document.body.clientWidth;
    }

    return 0;
};

function getBrowserHeight() {
    if (window.innerHeight) {
        return window.innerHeight;
    }
    else if (document.documentElement && document.documentElement.clientHeight != 0) {
        return document.documentElement.clientHeight;
    }
    else if (document.body) {
        return document.body.clientHeight;
    }
    return 0;
};

</script>



</head>

<body language="javascript" cellspacing="0" cellpadding="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
    <td width="1400">
    <div id="worldclockns">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="images/clock.gif" width="29" height="29" /></td>
          </tr>
          <tr>
            <td class="clock"><span id="worldclockLondon"></span><br />
          <!--  <span id="worldclockDubai"></span><br />
            <span id="worldclockSaiGon"></span><br />
			<span id="worldclockHongKong"></span><br />
            <span id="worldclockSydney"></span><br />-->

            </td>
          </tr>

        </table></div></td>
     <!--   <td align="right" valign="bottom">

        		<a href="https://forms.office.com/Pages/ResponsePage.aspx?id=qfx0Qb8XHUqmI9Yz1rA4aZGKO7rJ3GVMvJZritlXMMhUQkgyMUszMEMwMldGRUcxUlpQMVo1VEJDUS4u" target="_blank"><img src="images/covid_testresult.jpg"/></a>
	            <a href="https://forms.office.com/pages/responsepage.aspx?id=qfx0Qb8XHUqmI9Yz1rA4aZGKO7rJ3GVMvJZritlXMMhURFg4R1FNQkZNT0xXQTdPM0lKRTJROFY2WS4u" target="_blank"><img src="images/covid_vaccine.jpg"/></a>
	            <a href="https://kbytcq.khambenh.gov.vn/#tokhai_yte/model?&o=a685bd5f-b700-4aac-b8eb-b29f39893c4f&dd_id=d803079c-0383-46e6-af09-a3988e37b605&n=C%C3%94NG%20TY%20TNHH%20C%C3%94NG%20NGHI%E1%BB%86P%20TO%C3%80N%20C%E1%BA%A6U&stt=1" target="_blank"><img src="images/covid_declaration.jpg"/></a>
	            <a href="https://forms.office.com/pages/responsepage.aspx?id=qfx0Qb8XHUqmI9Yz1rA4aYLe65iDhWBKgJDCE4RxkzJUMU9IVDNZVjNaQU5WNzk4N0ZMNU5BSEEySy4u" target="_blank"><img src="images/covid_WFH.jpg"/></a> 
	            <a href="https://forms.office.com/pages/responsepage.aspx?id=qfx0Qb8XHUqmI9Yz1rA4aYLe65iDhWBKgJDCE4RxkzJUOERaODM4MThJWk9XSFBMQlRYWjZFV1hYWS4u" target="_blank"><img src="images/covid_RTO.jpg"/></a>
        </td>-->
      </tr>
    </table></td>
    <td></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="top" background="images/left_2022.jpg" width="497px">
		<form id="frmlogin" name="frmlogin" method="post" action>
		<table width="90%" border="0" cellspacing="0" cellpadding="3">
		<tr>
		  <td width="5%"></td>
            <td width="95%">&nbsp;</td>
          </tr>
          <tr>
		  <td></td>
            <td><span class="lable">Username</span></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
			<td>
			    <input name="txtusername" type="text" id="txtusername" tabindex="1" <%If strError <> "" Then%> value="<%=Request.Form("txtusername")%>" <%End If%> style="width:80%" />			</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td class="lable">Password</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><input name="txtpwd" type="password" tabindex="2" id="txtpwd" style="width:80%" /></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td></td>
                <td align="left" class="lable"><a href="javascript:checkin()" class="new_a">LOGIN</a> | <a href="tools/retrievepassword.asp" class="new_a"> Forgotten your password </a></td>
              </tr>
            </table></td>
          </tr>
		<tr>
            <td>&nbsp;</td>
            <td class="red"><input name="txtReturnUrl" type="hidden" id="txtReturnUrl" value="<%=strURL%>"/><%=strError%></td>
        </tr>

        </table>
		</form></td>
       <!-- <td><img src="images/right_2022.jpg"/></td>-->
      </tr>
      <tr>
        <td><img src="images/t_dot.gif" height="1" /></td>
        <td align="right">
		</td>
      </tr>
      <tr>

        <td>
			<span class="style3">
				<a href="http://ais.atlasindustries.com/staff/" class="style3">Atlas Staff</a> 
						</span>

        </td>
        <td align="right">

		<table>
			<tr>

				<!--<td align="right"><span class="style3">
					<a href="http://www.atlasindustries.com" class="style3">Atlas Industries</a> |
					<a href="https://atlasindustries.sharefile.com" class="style3">Atlas ShareFile</a> |
					<a href="https://atlasindustries.sharepoint.com" class="style3">Atlas Connect</a> |
<!--					<a href="https://login.salesforce.com/" class="style3">Salesforce</a>-->
					</span>
				</td>-->
				<td align="right"><img src="images/partnersInDesign.jpg" height="21" /></td>
			</tr>
			</table>
		</td>
      </tr>
    </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
<script language="JavaScript1.2">
var hotkey=13
if (document.layers)
document.captureEvents(Event.KEYPRESS)
function backhome(e){
	if (document.layers){
		if (e.which==hotkey)
			checkin()
	}
	else {
		if (event.keyCode==hotkey){
			event.keyCode = 0;
			checkin()
			}
		}
}
document.onkeypress=backhome
</script>



<!--<div id="divUpload" class="popup" style="width:722px; height:559px">
    <a id="closeButton" style="display:block; position:absolute;top: 10px; right: 10px; padding:5px ; background-color:RGB(195,197,195)" href="#">Close</a>
    <img src="images/popup/CardCompetition.jpg" />
</div>
<div id="background"></div>-->

</body>
</html>
