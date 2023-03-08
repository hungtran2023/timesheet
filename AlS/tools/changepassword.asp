<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/constants.inc"-->
<!-- #include file="../class/clsSHA-1.asp" -->
<!-- #include file = "../inc/library.asp"-->
<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb
	Dim gMessage
	
'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If					

'----------------------------------
' Get Full Name and Job Title
'----------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
'----------------------------------	
' Make list of menu
'----------------------------------
	If isEmpty(session("Menu")) then 
		getRes = getarrMenu(session("USERID"))
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End if	
	
	'current URL
	if Request.ServerVariables("QUERY_STRING")<>"" then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	else
		strURL = Request.ServerVariables("URL")
	end if
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "C"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../")
'-----------------------------------
' Analyse query string
'-----------------------------------

if Request.QueryString("fgMenu") <> "" then
	fgExecute = false
else
	fgExecute = true
end if
strAct = Request.QueryString("act")
gMessage=""
if fgExecute then
	if strAct = "SAVE" then
		strOld = Request.Form("txtold")
		strNew = Request.Form("txtnew")
		strCon = Request.Form("txtconfirm")
		Set objSHA1 = New clsSHA1	
		strDigest = ObjSHA1.SecureHash(strOld)
		strConnect = Application("g_strConnect") 
		Set objDb = New clsDatabase
		If objDb.dbConnect(strConnect) then
		  strQuery = "Select count(*) as mysum From ATC_Users Where UserID = " & session("USERID") & "and Password = '" & strDigest & "'"
		  ret = objDb.runQuery(strQuery)
		  if ret then
			if objDb.rsElement("mysum")=1 then '--------------starting update
				objDb.cnDatabase.BeginTrans
				strDigest = ObjSHA1.SecureHash(strNew)
				strQuery = "UPDATE ATC_Users SET Password = '" & strDigest & "' WHERE UserID = " & session("USERID")
				ret = objDb.runActionQuery(strQuery)
				if ret=false then				
					objDb.cnDatabase.RollbackTrans
					gMessage = objDb.strMessage
				else
					objDb.cnDatabase.CommitTrans
					gMessage = "Your password has been successfully updated."
					objDb.closerec
				end if
			else
			  gMessage = "The old password is incorrect. Be sure you are using password for Atlas Information System"
			end if
		  else
		    gMessage = objDb.strMessage
		  end if
		Else
		  gMessage = objDb.strMessage
		End if
		objDb.dbdisConnect
		set objDb = nothing
		Set ObjSHA1 = Nothing
	Elseif strAct = "LIST" then
	    
	    Set objSHA1 = New clsSHA1	
		strDigest = ObjSHA1.SecureHash(strOld)
		strConnect = Application("g_strConnect") 
		Set objDb = New clsDatabase
		
		If objDb.dbConnect(strConnect) then
		  strQuery = "SELECT * FROM ATC_Users WHERE NewPassword IS NOT NULL"
		  ret = objDb.runQuery(strQuery)
		  if ret then
	        if not objDb.noRecord then
	        
	            strDigest = ObjSHA1.SecureHash(strNew)
		        arrlistUsers = objDb.rsElement.GetRows
		        
		        strDigest = ObjSHA1.SecureHash(strOld)
		        
		        for i=0 to UBound(arrlistUsers,2)
		            strDigest = ObjSHA1.SecureHash(arrlistUsers(7,i))
		            
		            strSql="UPDATE ATC_Users SET EncrypNewPassword = '" & strDigest & "' WHERE UserID = " & arrlistUsers(0,i)
		            Response.Write strSql & "<br>"
		            
		        next
		            
		        objDb.CloseRec
	        else
		        arrlistUsers = ""
	        end if        
			
		  else
		    gMessage = objDb.strMessage
		  end if
		Else
		  gMessage = objDb.strMessage
		End if
		objDb.dbdisConnect
		set objDb = nothing
		Set ObjSHA1 = Nothing
	    
	end if
	
end if
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../timesheet.css"/>
<link href="../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>
<link href="../jQuery/atlasJquery.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../library/library.js"></script>
<script type="text/javascript" src="../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../jQuery/jquery-ui.min.js"></script>

<style type="text/css">

#Container 
{
	width:60%;
	background-color:#C0CAE6;
	border: 1px solid #003399;
	padding: 10px;
}   

#Container ul, li {
	margin:0;
	padding:3;
	list-style-type:none;	
}
#Container label {
	width:100;
	text-align:left
}

.submit ul
{
      list-style:none;
      text-align:center;
      padding:0;
}

.submit ul li
{       
    display:inline;
}

.submit a
{
    padding-top:3px;
    display:inline-block;
    width:60px;
    height:20px;
    background-color:#8CA0D1;
    text-align:center;
    font-weight: bold;
    text-decoration:none;
}

.submit a:hover
{
    background-color:#7791D1;
    color:white;
}


#errMsg {
	width:60%;
	padding: 10px;
	background:#fefefe;
	font-size:.8em;
	text-align:left;
}
#errMsg h5 {
	margin: 0;
	padding:0;
	font-weight:normal;
}

#errMsg li {
	margin:0;
	padding:0;
	list-style-type:circle;	
	}
	
</style>

<script type="text/javascript">

    $(document).ready(function() {

        //$("#pswd_info").hide();


        $("#txtNew").focus();
        $("a.checkdata").click(function(e) {

            //var errMsg = "<b>Password must meet the following requirements:</b><ul>";
            var errMsg = ""
            var oldPass = $("input#txtOld").val();
            var newPass = $("input#txtNew").val();
            var retypePass = $("input#txtRetype").val();

            //alert(newPass.match(/[`~!@#$%^&*()-+=|_{}:;"'<>?,.]/));
            //alert(newPass.match(/\W/));

            if (newPass.length < 8)
                errMsg = errMsg + "<li>Be at least <strong>8 characters</strong></li>";

            if (newPass.match(/[a-zA-Z]/) == null)
                errMsg = errMsg + "<li>At least <strong>one English uppercase or lowercase </strong></li>";
            if (newPass.match(/\d/) == null)
                errMsg = errMsg + "<li>At least <strong>one number</strong></li>";

            if (newPass.match(/\W/) == null)
                errMsg = errMsg + "<li>At least <strong>one nonalphanumeric character such as [@#$%^&*...]</strong></li>";

            if (errMsg != "")
                errMsg = " <h5>Password must meet the following requirements:</h5><ul>" + errMsg + "</ul>";

            if ((errMsg == "") && (retypePass != newPass))
                errMsg = errMsg + "Password does not match the confirm password.";

            if ((errMsg == "") && (oldPass == ""))
                errMsg = errMsg + "Please enter old password.";

            if (errMsg == "")
                save();
            else
                $("div#errMsg").html(errMsg);

        })

    });

    function CheckDataDone(){
        return false;
    }
    
    
    
</script>

<script type="text/javascript">

    function checkdata() {
        if (document.frmdetail.txtnew.value == "") {
            alert("Please enter your new password.");
            document.frmdetail.txtnew.focus();
            return false;
        }
        if (document.frmdetail.txtconfirm.value == "") {
            alert("Please re-enter your new password.");
            document.frmdetail.txtconfirm.focus();
            return false;
        }
        var strtmp1 = document.frmdetail.txtnew.value;
        var strtmp2 = document.frmdetail.txtconfirm.value;
        if ((strtmp1 != "") && (strtmp2 != "") && (strtmp1 != strtmp2)) {
            alert("New Password and Confirmation are not consistent!");
            document.frmdetail.txtconfirm.value = "";
            document.frmdetail.txtconfirm.focus();
            return false;
        }
        return true;
    }
    

    function save() {
            document.frmdetail.action = "changepassword.asp?act=SAVE";
            document.frmdetail.target = "_self";
            document.frmdetail.submit();
    }

    function List() {

        document.frmdetail.action = "changepassword.asp?act=LIST";
        document.frmdetail.target = "_self";
        document.frmdetail.submit();

    }
</script>
</head>

<body style="background-color:White; margin:0;">
<form name="frmdetail" method="post" action="">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>

			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(0))
			%>		
        <table width="100%">
          <tr> 
            <td  align="center">

                 
                <div class="title" style="padding:10px;">Change password</div>
                <div id="Container" class="blue-normal">
                    <ul>
                       
                        <li>
                            <label for="txtNew">New Password <span class="red">*</span></label>
                            <span><input id="txtNew" type="password" name="txtNew" /></span>
                        </li>
                        <li>
                            <label for="txtRetype">Confirm Password <span class="red">*</span></label>
                            <span><input id="txtRetype" type="password" name="txtRetype" /></span>
                        </li>
                        <li>
                        
                        </li>
                         <li>
                            <label for="txtOld">Old Password <span class="red">*</span></label>
                            <span><input id="txtOld" name="txtOld" type="password" /></span>
                        </li>
                    </ul>
                </div>
                <div class="submit">
                    <ul>
                        <li><a href="#" class="checkdata" name="save">Save</a></li>
                        
                    </ul>


               </div>
                
                <div id="errMsg" class="red"><%=gMessage%></div> 
               
            </td><!--End Container-->
          </tr>
        </table>
  			<%
			'--------------------------------------------------
			' Write the body of HTML page
			'--------------------------------------------------
			Response.Write(arrTmp(1))
			%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>
</body>
</html>