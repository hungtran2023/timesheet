<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/library.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<%
'StatusList =Array("In used", "Broken", "Loss", "Liquidate/Charity", "Stock")

'****************************************
' Function: Outbody
' Description: 
' Parameters: source recordset, number of lines on one page
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Outbody(ByRef rsSrc,byval intPage, ByVal psize)
	dim intStart,intFinish, strWarning
	
	strOut = ""

	if not rsSrc.EOF then
		
		rsSrc.AbsolutePage = intPage
		intStart = rsSrc.AbsolutePosition
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsData.PageSize - 1)
		End if
	
		For i = intStart to intFinish
			if i mod 2 = 0 then
				strColor = "#E7EBF5"
			else
				strColor = "#FFF2F2"
			end if
			strWarning=""
			if rsSrc("Spare") - rsSrc("AlmostEmpty")<=0 then strWarning="<div class='red'><strong>Warning:</strong>&nbsp;&nbsp; This cartridge is running out!</div>"
			
			strCartridgeCode=Showlabel(rsSrc("CartridgeCode"))
			
			strHref="<button type='button' class='clssEmpty'>" & Showlabel(rsSrc("AlmostEmpty")) & "</button>" & chr(13)
			strHref=strHref & "<div style='display: none;'><input type='text' class='blue-normal' style='width:50px' value='" & Showlabel(rsSrc("AlmostEmpty")) & "'>" & chr(13)
			strHref=strHref & "<a href='#' idCartridge=" & rsSrc("CartridgeID") & " class='lnkSubmit' style='padding:5px'>Submit</a><a href='#' class='lnkCancel'>Cancel</a></div>"
			
			strOut = strOut & "<tr bgcolor=" & strColor & ">" &_
			         "<td valign='top' align='center'><a href='javascript:getdetail(0," & rsSrc("CartridgeID") & ");'><img style='border: 0;' src='../../images/plus.png'></a>" & _
									"<a href='javascript:getdetail(1," & rsSrc("CartridgeID") & ");'><img style='border: 0;'  src='../../images/minus.png'></a></td>" &_
			         "<td valign='top'  class='blue-normal'>" & strCartridgeCode & "</td>" &_
			         "<td valign='top' align='center' class='blue-normal'>" & rsSrc("Spare") & "</td>" &_
			         "<td valign='top' align='center' class='blue-normal'>" & strHref & "</td>" &_
			         "<td valign='top' >" & strWarning & "</td>" &_
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function

'***************************************************************
'
'***************************************************************
function ExecuteSQL(strSql,prefix)

	dim strConnect,ret,strMessage
	dim objDb	

	strConnect = Application("g_strConnect") 
	Set objDb = New clsDatabase
		
	If objDb.dbConnect(strConnect) then
			
		ret = objDb.runActionQuery(strSql)
				
		if ret=false then				
			strMessage = objDb.strMessage
		else
			strMessage=prefix & " successfully."
		end if
			  
	else
		strMessage=objDb.strMessage
	end if
	
	ExecuteSQL=strMessage
end function
'------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, PageSize, fgUpdate, fgRight

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If					

'-----------------------------------
'Check ACCESS right
'-----------------------------------
	tmp = Request.ServerVariables("URL") 
	while Instr(tmp, "/")<>0
		tmp = mid(tmp, Instr(tmp, "/") + 1, len(tmp))
	Wend
	strFilename = tmp
	if isEmpty(session("Righton")) then
		fgRight = false
	else
		getRight = session("Righton")
		fgRight = false
		for ii = 0 to Ubound(getRight, 2)
			if getRight(0, ii) = tmp then
				fgRight=true
				fgUpdate = false
				if getRight(1, ii) = 1 then fgUpdate = true	'updateable right
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if

'-------------------------------
' Calculate pagesize
'-------------------------------
	if not isEmpty(session("Preferences")) then
		arrPre = session("Preferences")
		if arrPre(1, 0)>0 then PageSize = arrPre(1, 0) else PageSize = PageSizeDefault
		set arrPre = nothing
	else
		PageSize = PageSizeDefault
	end if
	
'-------------------------------
' Get Fullname and Job Title
'-------------------------------
	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
'-----------------------------
' Make list of menu
'-----------------------------
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
	if strChoseMenu = "" then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")
'-----------------------------
' read data
'-----------------------------

	if Request.QueryString("act")="u" then
		intCategoryID=Request.Form("txtID")
		intAlmostEmpty=Request.Form("txtAlmost")
		
		 strSql="UPDATE ATC_Cartridges SET " & _
					"AlmostEmpty=" & intAlmostEmpty & _
				" WHERE CartridgeID= " & intCategoryID
					
		gMessage = ExecuteSQL(strSql,"Update AlmostEmpty value ")	
	end if

	strSearch=Request.Form("txtSearch")

	strSQL="SELECT b.CartridgeID , b.CartridgeCode, b.CartridgeNote, b.AlmostEmpty, ISNULL(a.Total,0) as Spare " & _
				"FROM (SELECT CartridgeID, SUM(Qty) as Total FROM ATC_InkStock GROUP BY CartridgeID ) AS a RIGHT OUTER JOIN  " & _
                      "ATC_Cartridges AS b ON a.CartridgeID = b.CartridgeID " & _
			"WHERE b.fgActivate=1 "

	Call GetRecordset(strSQL,rsData)
	
'--------------------------------------------------
'Start Paging
'--------------------------------------------------

' Set the PageSize, CacheSize and populate the intPageCount
	rsData.PageSize=PageSize
' The Cachesize property sets the number of records that will be cached locally in memory	
	rsData.CacheSize=rsData.PageSize	
	intPageCount=rsData.PageCount
	intRecordCount=rsData.RecordCount
	
' Checking to make sure that we are not before the start or beyond end of the recordset
' If we are beyond the end, set the current page equal the last page of the recordset.
' If we are before the start, set the current page equal the start of the recordset
	
	intPage=Request.QueryString("Navi")

	if intPage="" then intPage=1
	
	if cint(intPage)>Cint(intPageCount) then intPage=intPageCount
	if cint(intPage)<=0 then intPage=1

'--------------------------------------------------
'End Paging	
'--------------------------------------------------
	
	strLast=Outbody(rsData,intPage,PageSize)

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_pro.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intPageCount)	
End if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link rel="stylesheet" href="../../timesheet.css">


<style>
.clssEmpty {
    background-color: #8CA0D1; /* black */
    border: none;
    color: white;
    padding: 2px 20px;
    text-align: center;
    text-decoration: none;
    display: inline-block;
    margin: 1px 1px;
    cursor: pointer;
}

</style>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="navi" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			Response.Write(arrTmp(0))
			%>
          <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor=<%if gMessage="" then%>"FFFFFF"<%else%>"#E7EBF5"<%end if%>>
					<td class="red" colspan="5" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
				</tr>
               
                <tr> 
                  <td class="title" height="50" align="center" colspan="5">Cartridge List</td>
                </tr>  
			</table>
            </td>
          </tr>
          
          <tr> 
            <td height="100%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style=height:"79%" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="8CA0D1"> 
							  <td class="blue" align="center"  width="15%">Action</td>
                              <td class="blue" align="center"  width="35%">Cartridge Name</td>
                              <td class="blue" align="center" width="10%">Spare</td>
                              <td class="blue" align="center" width="15%">Almost empty</td>
                              <td class="blue" align="center" width="25%">Status</td>
                            </tr>
<%
	Response.Write(strLast)
%>                            
                          </table>
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal"> 
                                &nbsp;&nbsp;</td>
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
<input type="hidden" name="txtID" value="">
<input type="hidden" name="txtStt" value="">
<input type="hidden" name="txtAlmost" value="">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">

<script language="javascript" src="../../library/library.js"></script>

<script type="text/javascript" src="../../js/jquery.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>
<script type="text/javascript" src="../../js/bootstrap-datepicker.js" charset="UTF-8"></script>
<script type="text/javascript" src="../../js/bootstrap-table.js"></script>
<script type="text/javascript" src="../../js/js-control.js"></script>
<script type="text/javascript" src="../../js/formValidation.min.js"></script>
<script type="text/javascript" src="../../js/framework/bootstrap.min.js"></script>

<script>
<!--
$(document).ready(function() {
//alert ("test");
	$("button").click(function(){
		$(this).hide();
		$(this).next().show();
		
	});
	
	$(".lnkCancel").click(function(){
		$(this).parent().hide();
		$(this).parent().prev().show();
		//alert($(this).attr("idCartridge"));
	});
	
	$(".lnkSubmit").click(function(){
		
		updateAlmostempty($(this).attr("idCartridge"),$(this).prev().val());
		//alert($(this).prev().val());
		//alert($(this).attr("idCartridge"));
	});
	
 });
 
function getdetail(stt,varid){
	
	document.navi.txtStt.value = stt;	
	document.navi.txtID.value = varid;	
	document.navi.action = "CartridgeInOut.asp";
	document.navi.target = "_self";
	document.navi.submit();
}

function updateAlmostempty(varid,intvalue){
	
	document.navi.txtID.value = varid;	
	document.navi.txtAlmost.value = intvalue;	
	document.navi.action = "CartridgeList.asp?act=u";
	document.navi.target = "_self";
	document.navi.submit();
}

-->
</script>

</form>

</html>