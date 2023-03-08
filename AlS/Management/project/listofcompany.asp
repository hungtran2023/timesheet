<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->
<%
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
	dim intStart,intFinish
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

			strOut = strOut & "<tr bgcolor=" & strColor & " idValue='" & rsSrc("CompanyID") &"'>" &_
			         "<td class='blue-normal'>" &  UCase(Showlabel(rsSrc("CharCode"))) & "</td>" &_
			         "<td class='blue-normal'>" & Showlabel(rsSrc("CompanyName")) & "</td>" &_
					 "<td class='blue-normal'>" & Showlabel(rsSrc("ClientType")) & "</td>" &_
			         "<td class='blue-normal'>" & Showlabel(rsSrc("SeverPath")) & "</td>" &_
					 "<td class='blue-normal'>" & Showlabel(rsSrc("Website")) & "</td>" &_
			         "<td class='blue-normal' >" & Showlabel(rsSrc("EmailAddress")) & "</td>" &_	
			         "</tr>" & chr(13)
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	end if
	Outbody = strOut
end function

'****************************************
' Function: Pagination
' Description: 
' Parameters: source recordset, number of lines on one page
'			  
' Return value: rows of table
' Author: 
' Date: 
' Note:
'****************************************
function Pagination()
	dim num,i
	dim strDisabledPre, strDisabledNext
	strDisabledPre=""
	strDisabledNext=""
	
	num=intPageCount
	
	if intPage=1 then strDisabledPre="disabled"
	if cint(intPage)=cint(intPageCount) then strDisabledNext="disabled"
	
	
	
	strOut = "<nav>" & _
			  "<ul class='pagination'>" & _
					"<li class ='pag_prev " & strDisabledPre &"'><a class='page-link' href='#' tabindex='-1'>Previous</a></li>" 
	for i=1 to intPageCount
		strOut=strOut &	"<li class= 'numeros "
		if i=cint(intPage) then	strOut=strOut &	"active"
		strOut=strOut & "'><a class='page-link' href='#'>" & i & "</a></li>"
	next
	
	strOut=strOut &	"<li class =' pag_next " & strDisabledNext &"'> <a class='page-link' href='#'>Next</a></li></ul></nav>"
	
	Pagination = strOut
end function

'***************************************************************
'
'***************************************************************
function ExecuteSQL(strSql)

	dim objDatabase
	dim strCnn
	dim blnReturn
	
	blnReturn = false	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase 	
	
	If objDatabase.dbConnect(strCnn) then		
		blnReturn= (objDatabase.runActionQuery(strSql))	
		strError=" Update successful."
		if not blnReturn then strError=objDatabase.strMessage		
	else
		strError=objDatabase.strMessage
	end if
	
	Set objDatabase = nothing
	ExecuteSQL=strError
	
end function
'------------------------------------------------------------------------------
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage, PageSize, fgUpdate, fgRight
	Dim intPageCount,intPage
	Dim strClientName, intClientType, strWebsite,strEmaildomain,strServerPath,strNote,blnTP
	

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
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
'Get list of data
'--------------------------------------------------
	strAct=Request.QueryString("act")
	
	If strAct="save" then
	
		intCompanyID=request.form("txtID")
		strClientName=request.form("txtClientName")
		intClientType=request.form("lstClientType")
		strWebsite=request.form("txtWebsite")
		strEmaildomain=request.form("txtEmaildomain")
		strServerPath=request.form("txtServerPath")
		strNote=request.form("txtNote")
		intType=1
		if request.form("radTP") = 1 then intType=2		
		
		strSql="UPDATE  ATC_Companies " & _
                    "SET  CompanyName  = '" & replace(strClientName,"'","''") & "'"& _
					", EmailAddress  = " & IIF(strEmaildomain="", "NULL", "'" & replace(strEmaildomain,"'","''") & "'") & _
                    ", Website  = " & IIF(strWebSite="", "NULL", "'" & replace(strWebsite,"'","''") & "'") & _
                    ", Note  = " & IIF(strNote="","NULL", "'" & Replace(strNote,"'","''") & "'") & _
                    ", SeverPath  =" & IIF(strServerPath="","NULL", "'" & Replace(strServerPath,"'","''") & "'") & _
                    ", [type]=" & intType & _
					", [ClientTypeID]=" & IIF(intClientType="","NULL", intClientType) & _ 
                " WHERE CompanyID=" & intCompanyID
		strError= ExecuteSQL(strSql)
		
	end if
	
	intCompanyID=Request.QueryString("id")	
	if intCompanyID="" then intCompanyID=-1

	strSearch=Request.Form("txtSearch")

	strSQL="SELECT [CompanyID],[CompanyName],[EmailAddress],[Website],[Note],[CharCode],[SeverPath],[Type],ISNULL(a.ClientTypeID,-1) as [ClientTypeID], ISNULL(b.Descriptions,'') as ClientType FROM ATC_Companies a  " & _
				" LEFT JOIN ATC_ClientTypes b ON a.ClientTypeID=b.ClientTypeID "

	if trim(strSearch<>"") then	
		strSQL=strSQL & " WHERE " & " CompanyName like '%" & trim(strSearch) & "%'"   & " OR  CharCode like '%" & trim(strSearch) & "%'"
	end if
	
	strSQL=strSQL & " ORDER BY CharCode"

	Call GetRecordset(strSQL,rsData)

	if cdbl(intCompanyID)<>-1 then
		rsData.filter="CompanyID=" & intCompanyID
		strClientName=rsData("CompanyName")
		intClientType=rsData("ClientTypeID")
		strWebsite=rsData("Website")
		strEmaildomain=rsData("EmailAddress")
		strServerPath=rsData("SeverPath")
		strNote=rsData("Note")
		blnTP=(rsData("Type")=2)
		
		strSql="SELECT * FROM ATC_ClientTypes WHERE fgActivate=1"
		Call GetRecordset(strSQL,rsClientTyle)
		strClientTyle= PopulateDataToListWithoutSelectTag(rsClientTyle,"ClientTypeID", "Descriptions",cdbl(intClientType))
		
		rsData.filter=""
		rsData.Movefirst
	end if

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

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
End if
%>	

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">
<head>
<meta http-equiv="Content-type" content="text/html;charset=UTF-8"/>
<meta http-equiv="Content-Language" content="en"/>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>


<title>Atlas Industries Time Sheet System</title>


    <link href="../../bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
    <link href="../../css/timesheet.css" rel="stylesheet" >
    <link href="../../css/style.css" rel="stylesheet" type="text/css">    
	<link href="../../css/bootstrapValidator.min.css" rel="stylesheet" type="text/css" />
</head>

<body data-pinterest-extension-installed="cr1.39.1">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------
Response.Write(arrPageTemplate(0))
'--------------------------------------------------
' Write the body of HTML page
'--------------------------------------------------
Response.Write(arrTmp(0))
%>
<!--style="background-color: #00ff00;"-->
<div class="container-fluid" >
	<%If strError <> "" Then%>  
		<div class="row">	
			<div class="<%if strError<>" Update successful." then %>alert alert-danger<%else%>alert alert-success<%end if%>">
				<strong><%if strError<>" Update successful." then %>Error!<%else%>Success!<%end if%></strong><%=strError%>
			</div>
		</div>
	<% End If%>	
	<div class="row" style="padding:20px 0px 0px 20px;">   <h3>        List of Clients</h3>  </div>
	<div class="row">
        <div class="col-sm-6 col-sm-offset-3">
            <div id="imaginary_container"> 
			<form name="searchform" method="post" action="listofcompany.asp">
				<div class="form-group">
					<div class="input-group stylish-input-group">
						<input type="text" name="txtSearch" id="txtSearch" onkeyup="myFunction()" class="form-control"  placeholder="Search by Client name or Client code" >
						<span class="input-group-addon">
							<button type="submit" >
								<span class="glyphicon glyphicon-search"></span>
							</button>  
							
						</span>
					</div>
				</div>
			</form>
            </div>
        </div>
	</div>
	<div class="row">
		<div class="table-responsive">	
			<table class="table table-hover" id="tblListClient">
				<thead class="thead-inverse tableheaderblue" >
					<tr>						
						
						<th>Client Code</th>
						<th>Client name</th>
						<th>Client Type</th>						
						<th>Server Path</th>
						<th>Website</th>
						<th >Email Domain</th>						
					</tr>
				</thead>
				<tbody>
					<%=strLast%>					
				</tbody>
				
			</table>			
		</div>			
    </div> 
	<div class="row">
		<div class="text-center">
			<%=Pagination()%>
		</div>
	</div>
 <!-- Modal HTML -->
    <div id="myModal" class="modal">
        <div class="modal-dialog">
            <div class="modal-content">
			<form id="detailform" class="form-horizontal" action="listofcompany.asp?act=save" method="post">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                    <h4 class="modal-title">View detail of client</h4>
                </div>
                <div class="modal-body">
                    
						<div class="form-group">
							<label class="col-md-12">Client Name</label>
							<div class="col-md-12">
								<input type="text" id="txtClientName" name="txtClientName" class="form-control" method="post"  value="<%=strClientName%>">
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Client Type</label>
							<div class="col-md-12">
								 <select id="lstClientType" name="lstClientType" class="form-control" >
									<option value=""></option>
									<%=strClientTyle%>
								</select>
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Website</label>
							<div class="col-md-12">
								<input type="text" id="txtWebsite" name="txtWebsite" class="form-control"  value="<%=strWebsite%>">
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Email domain</label>
							<div class="col-md-12">
								<input type="text" id="txtEmaildomain" name="txtEmaildomain" class="form-control"  value="<%=strEmaildomain%>">
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Server Path</label>
							<div class="col-md-12">
								<input type="text" id="txtServerPath" name="txtServerPath" class="form-control"  value="<%=strServerPath%>">
							</div>
						</div>
						<div class="form-group">
							<label class="col-md-12">Note</label>
							<div class="col-md-12">
								<textarea id="txtNote" name="txtNote"  cols="30" rows="3" class="form-control"><%=strNote%></textarea>
							</div>
						</div>
						
						<div class="form-group">
							<div class="col-md-12">
								<div class="col-md-1 no-padding width-auto">
									<input type="checkbox" name="radTP" id="radTP" value="1" class="no-padding" <%if blnTP then%>checked<%end if%>>
								</div>
								<label class="col-md-3 padding-left5 no-blod" for="radTP">Third Party</label>
							</div>
						</div>
						<input type="hidden" id="txtID" name="txtID" value="<%=intCompanyID%>">
					
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                    <button type="submit" class="btn btn-primary" id="btnSave">Save changes</button>
                </div>
				</form>
            </div>
        </div>
    </div>
</div>
<!--Content-->		    	
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

<script type="text/javascript" src="../../js/jquery.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>
<script type="text/javascript" src="../../js/formValidation.min.js"></script>
<script type="text/javascript" src="../../js/framework/bootstrap.min.js"></script>
<script>
var currentPage = <%=intPage%>;
var clientID = <%=intCompanyID%>;

$( document ).ready(function() { 
	

	$(window).load(function(){
		if (clientID!=-1)
			$('#myModal').modal({backdrop: "static"});
     });
	
    $(".pagination li.pag_prev").click(function() {
        if($(this).next().is('.active')) return;
        showPage(currentPage-1);
        
        //showPage();
    });
	
	$(".pagination li.pag_next").click(function() {
        if($(this).prev().is('.active')) return;
        showPage(currentPage+1);
    });
	
	$(".pagination li.numeros").click(function() {       
        showPage(parseInt($(this).text()));
    });
	
	
    // Align modal when it is displayed
    $(".modal").on("shown.bs.modal", alignModal);
    
    // Align modal when user resize the window
    $(window).on("resize", function(){
        $(".modal:visible").each(alignModal);
    });   
	
	$("#tblListClient tbody tr").click(function(e) {
		e.preventDefault();
		var id = $(this).attr("idValue");
		//alert(id);
		window.location = "listofcompany.asp?navi=" + currentPage + "&id=" + id ;
        
    });
	
	 $('#detailform').formValidation({
		framework: 'bootstrap',
        icon: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },
        fields: {
			txtClientName: {
	
                validators: {
                    notEmpty: {
                        message: 'The Client name is required'
                    }
                }
            }
		}
	 });
	
});

function alignModal(){
	var modalDialog = $(this).find(".modal-dialog");
	
	// Applying the top margin on modal dialog to align it vertically center
	modalDialog.css("margin-top", Math.max(0, ($(window).height() - modalDialog.height()) / 2));
}
	
function showPage(intpage){
	window.location = "listofcompany.asp?navi=" + intpage;
}

function myFunction() {
   var input, filter, table, tr, td, i;
  input = document.getElementById("txtSearch");
  filter = input.value.toUpperCase();
  table = document.getElementById("tblListClient");
  tr = table.getElementsByTagName("tr");

  // Loop through all table rows, and hide those who don't match the search query
  for (i = 0; i < tr.length; i++) {
    td = tr[i].getElementsByTagName("td")[0];
    if (td) {
      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
        tr[i].style.display = "";
      } else {
        tr[i].style.display = "none";
      }
    } 
  }
}
</script>
</script>

</body>
</html>