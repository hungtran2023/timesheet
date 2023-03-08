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

Function Outbody(ByRef rsSrc, ByVal psize)
	strOut = ""
	dim strLinkPro,strLinkInvoice,strColor,strDate,i
	dim strFileLink,strServerPath
	arrStatus = Array("","Live","Lost","Closed")
	strServerPath="..\..\data\CSO\"
	rsSrc.AbsolutePage = intPage
	intStart = rsSrc.AbsolutePosition
	
	If CInt(intPage) = CInt(intPageCount) Then
		intFinish = intRecordCount
	Else
		intFinish = intStart + (rsSrc.PageSize - 1)
	End if
	
	If Not rsSrc.EOF Then
		For i = intStart To intFinish
			strColor = "#FFF2F2"
			If i Mod 2 = 0 Then	strColor = "#E7EBF5"
			strDate=ConvertDate(cdate(rsSrc("DateTransfer")))
			strLinkPro="<a href='javascript:viewpro(&quot;" & rsSrc("ProjectKey") & "&quot;,&quot;" & rsSrc("fgStatus") & "&quot;,&quot;" & rsSrc("DateTransfer") & "&quot;," & rsSrc("ManagerID") & ");' " &_
					         "class='c' OnMouseOver = 'self.status=&quot;Project Detail&quot; ; return true' OnMouseOut =" &_
					         " 'self.status = &quot;&quot;'>" & Showlabel(rsSrc("ProjectKey")) & "</a>"
			If rsSrc("fgStatus") = "New" and not fgApproval then strLinkPro=Showlabel(rsSrc("ProjectKey"))
			
			strSignContract=""	
			if rsSrc("SignContract")=1 then 
				strSignContract="<img src='../../images/notyet.gif'>"
			elseif rsSrc("SignContract")=2 then
				strSignContract="<img src='../../images/icon_doc_download.gif' border=0>"
			end if
			
			if rsSrc("CSOFileName")<>"" then strSignContract="<a href='#' path='" & strServerPath & rsSrc("CSOFileName") & "'  class='cso'>" & strSignContract & "</a>"
			
			strUtilised=IIF(cint(rsSrc("ProjectKey2"))=1,"<img src='../../images/yes.gif'>","")
			strBillable=IIF(rsSrc("Billable"),"<img src='../../images/yes.gif'>","")
			
			if cdbl(rsSrc("invValue"))=0 then
				strLinkInvoice="<a href='#' class='inv c'>--</a>"
			else
				strLinkInvoice="<a href='#' class='inv c'>"& FormatNumber(cdbl(rsSrc("invValue"))) '& " " & rsSrc("CurrencyCode") &"</a>"
			end if
			
			
			if cdbl(rsSrc("ProposalValue"))=0 then 
				strProposalValue=""
			else
				strProposalValue=FormatNumber(rsSrc("ProposalValue")) '& " " & rsSrc("CurrencyCode")
			end if
			
			
			if cdbl(rsSrc("AwardedValue"))=0 then 
				strAwardedValue=""
			else
				strAwardedValue=FormatNumber(rsSrc("AwardedValue"))'& " " & rsSrc("CurrencyCode")
			end if
'response.write 			rsSrc("ProjectName")
			strOut = strOut & "<tr bgcolor=" & strColor & " status='" & rsSrc("fgStatus")&"'>" &_
			         "<td valign='top' class='blue-normal-small'>" & Showlabel(rsSrc("ProjectKey")) & "</td>" &_
			         "<td valign='top' class='blue-normal-small'>" & rsSrc("ProjectName") & "</td>" &_
					 "<td valign='top' class='blue-normal-small'>" & day(rsSrc("DateTransfer")) & "/" & month(rsSrc("DateTransfer")) & "/" & Year(rsSrc("DateTransfer")) & "</td>" &_
					 "<td valign='top' class='blue-normal-small'>" & Showlabel(rsSrc("Department")) & "</td>" &_
					 "<td valign='top' class='blue-normal-small'>" & Showlabel(rsSrc("BDM")) & "</td>" &_
					  "<td valign='top' class='blue-normal-small' idValue="& rsSrc("ManagerID") &">" & Showlabel(rsSrc("Manager")) & "</td>" &_
					  "<td valign='top' class='blue-normal-small' align='center'>" & arrStatus(rsSrc("ProjStatus")) & "</td>" &_
					  "<td valign='top' class='blue-normal-small' align='center'>" & rsSrc("CurrencyCode") & "</td>" &_
					  "<td valign='top' class='blue-normal-small' align='right'>" & strProposalValue & "</td>" &_
					  "<td valign='top' class='blue-normal-small' align='right'>" & strAwardedValue & "</td>" &_
			         "<td valign='top' class='blue-normal-small' align='right'>" & strLinkInvoice & "</td>" &_
			         "<td valign='top' class='blue' align='center'>" & strSignContract & "</td>" &_
			         "</tr>" & chr(13)                                                                                                                                                                                                                                                                                                                                                                                                    
					         
			rsSrc.MoveNext
			If rsSrc.EOF Then Exit For
		Next
	End If
	Outbody = strOut
End Function
'****************************************
' Function: Header Sort Column
' Description: 
' Parameters: 
'			  
' Return value: Header Sort Column
' Author: 
' Date: 
' Note:
'****************************************
function HeaderSortColumn
	dim strHeader,strSortIcon
	dim arrColumns,i
	arrColumns=array("<th @icon@><span class='sort-by 0'>APK </span></th>","<th @icon@><span class='sort-by 1'>Project Name</span></th>" ,"<th @icon@><span class='sort-by 2'>Register Date</span></th>")
	strHeader=""
	strSortIcon="class='glyphicon glyphicon-triangle-" & IIF(cint(right(varSortType,1))=1,"bottom'","top'")
	for i=0 to Ubound(arrColumns)
		if (cint(left(varSortType,1))=i) then
			strHeader=strHeader & replace(arrColumns(i),"@icon@",strSortIcon)
		else
			strHeader=strHeader & replace(arrColumns(i),"@icon@","")
		end if 
	next	
	HeaderSortColumn=strHeader
end function
'--------------------------------------------------------------------------------
'Retrive data from project
'--------------------------------------------------------------------------------
Sub GetProjectData(byval intSearchType, byval intBillable,byval strSearch,byval intBooked,byref rsReturn)
	dim strConnect,objDb,strQuery
	dim arrfields,arrsortType
	
	arrfields = Array("b.ProjectID","ProjectName","SortDate")
	arrsortType=Array("ASC", "DESC")
	
	strConnect = Application("g_strConnect")
	Set objDb = New clsDatabase
	objDb.recConnect(strConnect)	
	if strSearch<>"" then
		strSearch = replace(strSearch, "%", "")
		strSearch = Replace(strSearch, "'", "''")
	end if
	strQuery = "SELECT b.ProjectID AS ProjectKey,Projectkey2, " & _
				" ProjectName, ISNULL(CSOFilename,'') as CSOFilename, DateTransfer, (CASE WHEN CHARINDEX('___',a.ProjectID,7) > 1 THEN 'New' ELSE 'Issued' END) AS fgStatus, " & _
				" CSOApproval,SignContract,CSOCompleted,ManagerID,billable,(c.FirstName + ' ' + c.LastName) as Manager , e.Department, f.Fullname as BDM, " & _ 
				" ISNULL(g.[ProjStatus],0) as ProjStatus,ISNULL([ProposalValue],0) as ProposalValue,ISNULL([AwardedValue],0) as AwardedValue, b.CurrencyCode, ISNULL(h.invValue,0) as invValue" &_
				" FROM ATC_ProjectStage a INNER JOIN ATC_Projects b ON a.ProjectID=b.ProjectID " & _
				" INNER JOIN ATC_Department e ON b.DepartmentID=e.DepartmentID " &_
				" LEFT JOIN ATC_Companies d ON b.CompanyID=d.CompanyID " & _
				" LEFT JOIN ATC_PersonalInfo c ON ManagerID=c.PersonID " & _
				" LEFT JOIN HR_BDM f ON f.BDMID=b.BDMID " & _
				" LEFT JOIN Pro_GetLastTracking g ON g.ProjectID=a.ProjectID " & _
				" LEFT JOIN (SELECT projectID, sum(InvoiceValue) as invValue FROM ATC_ProjectInvoices GROUP BY projectID) h ON h.ProjectID=a.ProjectID " & _
				" WHERE b.fgDelete = 0"
'response.write 			strQuery	

	strQuerySearch=""
	
	if not fgRight then strQuery = strQuery & " AND " & getWherePhase("b",session("USERID"))
	
	strQuerySearch= strQuerySearch & " AND (b.ProjectID Like '%" & strSearch & "%' OR ProjectName Like '%" & strSearch & "%')"
	
	strQuery=strQuery & strQuerySearch & " ORDER BY " & arrfields(cint(left(varSortType,1))) & " "  & arrsortType(cint(right(varSortType,1)))
	
	If objDb.openRec(strQuery) Then
		objDb.recDisConnect
		set rsReturn = objDb.rsElement.Clone
		if not objDb.noRecord then 
			session("NumPage")=pageCount(rsReturn.Clone, PageSize)
		else
			gMessage = "No results found."
		end if
		objDb.CloseRec
	Else
		gMessage = objDb.strMessage	  
	End if
	
	Set objDb = Nothing
end sub

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
	
	dim intstart, intfinish
	
	strDisabledPre=""
	strDisabledNext=""
	
	num=10
	if intPage mod num<>0 then
		intstart=(intPage\num)*num + 1
	else
		intstart=((intPage-1)\num)*num + 1
	end if

	intfinish=intstart+num-1
'response.write intPage & ":" & intstart & "->" & intfinish	
	if intfinish>intPageCount then intfinish=intPageCount
		
	if intPage=1 then strDisabledPre="disabled"
	if cint(intPage)=cint(intPageCount) then strDisabledNext="disabled"
		
	strOut = "<nav>" & _
			  "<ul class='pagination'>" & _
					"<li class ='pag_prev " & strDisabledPre &"'><a class='page-link' href='#' tabindex='-1'>Previous</a></li>" 
	for i=intstart to intfinish
		strOut=strOut &	"<li class= 'numeros "
		if i=cint(intPage) then	strOut=strOut &	"active"
		strOut=strOut & "'><a class='page-link' href='#'>" & i & "</a></li>"
	next
	
	strOut=strOut &	"<li class =' pag_next " & strDisabledNext &"'> <a class='page-link' href='#'>Next</a></li></ul></nav>"
	
	Pagination = strOut
end function

'===============================================================================
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------

	Dim rsProject, gMessage, PageSize
	dim varSearch,varSearchType,varBillable,varPage,varSortType,varUpdate
	Dim strProjectID, strHeaderSort

	varSearch = trim(Request.QueryString("search"))
	varSearch = trim(Request.Form("txtSearch"))
	
	varPage=Request.QueryString("Page")
	
	varSortType=Request.Form("txtsort")

	if varSortType="" then varSortType="21" 'ORDER BY Sortdate DESC
	
	
	varUpdate=Request.QueryString("active")
	
'--------------------------------------------------
' Check session variable If it was expired or Not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					

	intUserID = session("USERID")
	
'--------------------------------------------------
' Calculate pagesize
'--------------------------------------------------

	If Not isEmpty(session("Preferences")) Then
		arrPre = session("Preferences")
		If arrPre(1, 0)>0 Then PageSize = arrPre(1, 0) Else PageSize = PageSizeDefault
		Set arrPre = Nothing
	Else
		PageSize = PageSizeDefault
	End If
	
	PageSize=300
	
'--------------------------------------------------
' Check ACCESS right
'--------------------------------------------------

	strTemp = Request.ServerVariables("URL") 
	While Instr(strTemp, "/")<>0
		strTemp = Mid(strTemp, Instr(strTemp, "/") + 1, Len(strTemp))
	Wend
	
	strFilename = strTemp
	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = strTemp Then
				fgRight=True
				Exit For
			End If
		Next
		
		Set varGetRight = Nothing		
	End If	
	If fgRight = False Then		
		Response.Redirect("../../welcome.asp")
	End If

'--------------------------------------------------
' Check VIEWALL project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRight = False
	Else
		varGetRight = session("RightOn")
		fgRight = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "View all projects" Then
				fgRight = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Check Approving Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgApproval = False
	Else
		varGetRight = session("RightOn")
		fgApproval = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "approving project" Then
				fgApproval = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Check Registration Project right
'--------------------------------------------------

	If isEmpty(session("RightOn")) Then
		fgRegister = False
	Else
		varGetRight = session("RightOn")
		fgRegister = False
		For ii = 0 To Ubound(varGetRight, 2)
			If varGetRight(0, ii) = "registration" Then
				fgRegister = True
				Exit For
			End If
		Next
		Set varGetRight = Nothing
	End If

'--------------------------------------------------
' Analyse query and prepare project list
'--------------------------------------------------
'varPage="" --> didn't retrive data from database
if varPage="" then	
	varPage=1
	call GetProjectData(varSearchType,varBillable,varSearch,varBooked,rsProject)	
	if rsProject.RecordCount>0 then
		if not isEmpty(session("rsProject")) then session("rsProject") = empty
		set session("rsProject")=rsProject.Clone
	else
		if not IsEmpty(Session("rsProject")) then set rsProject = session("rsProject")
		if rsProject.Recordcount>0 then rsProject.MoveFirst
	end if
else
	set rsProject = session("rsProject")
	rsProject.MoveFirst
	if varSortType<>"" then
		rsProject.Sort = varCol & " " & varSortType		
	end if
	rsProject.Move (cint(varPage)-1)* PageSize
	
end if


' Set the PageSize, CacheSize and populate the intPageCount

	rsProject.PageSize=PageSize
' The Cachesize property sets the number of records that will be cached locally in memory	
	rsProject.CacheSize=rsProject.PageSize	
	intPageCount=rsProject.PageCount
	intRecordCount=rsProject.RecordCount
	
' Checking to make sure that we are not before the start or beyond end of the recordset
' If we are beyond the end, set the current page equal the last page of the recordset.
' If we are before the start, set the current page equal the start of the recordset
	
	intPage=Request.QueryString("Navi")

	if intPage="" then intPage=1
	
	if cint(intPage)>Cint(intPageCount) then intPage=intPageCount
	if cint(intPage)<=0 then intPage=1	

	strLast = Outbody(rsProject, PageSize)
	
'--------------------------------------------------
' Get Fullname and Job Title
'--------------------------------------------------

	Set objEmployee = New clsEmployee	
	objEmployee.SetFullName(intUserID)
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)
	
	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'>" & strtmp1 & "&nbsp;&nbsp;&nbsp;" &_
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing
	
'--------------------------------------------------
' Make list of menu
'--------------------------------------------------

	If isEmpty(session("Menu")) Then 
		getRes = getarrMenu(intUserID)
		session("Menu") = getRes
	Else
		getRes = session("Menu")
	End If	
	
	'current URL
	If Request.ServerVariables("QUERY_STRING")<>"" Then
		strURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		strURL = Request.ServerVariables("URL")
	End If
	
	strChoseMenu = Request.QueryString("choose_menu")
	if strChoseMenu = "" then strChoseMenu = "AC"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

'--------------------------------------------------
' Read template page from file
'--------------------------------------------------
Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/",  "ats_menu.htm")

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

<meta http-equiv="Content-Language" content="en"/>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Atlas Industries Timesheet System</title>

<link href="../../bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../bootstrap/css/dataTables.bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../../css/timesheet.css" rel="stylesheet" >
<link href="../../css/style.css" rel="stylesheet" type="text/css">    
<style>

</style>
</head>
<body data-pinterest-extension-installed="cr1.39.1">

<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
	
%>

<div class="container-fluid" >
<%If strError <> "" Then%>  
		<div class="row">	
			<div class="<%if strError="Update successfull." then %>alert alert-danger<%else%>alert alert-success<%end if%>">
				<strong>Error:</strong><%=strError%>
			</div>
		</div>
	<% End If%>	
	<div class="row" style="padding:20px 0px 0px 20px;"><h3> List of Projects</h3></div>
	<div class="row">
        <div class="col-sm-6 col-sm-offset-3">
            <div id="imaginary_container"> 
			<form id="frmSearch" method="post" action="n_projectlist.asp">
				<div class="form-group">
					<div class="input-group stylish-input-group">
						<input type="text" name="txtSearch" id="txtSearch" onkeyup="myFunction()" class="form-control"  placeholder="Search by APK or Project Name" value="<%=varSearch%>">
						<span class="input-group-addon">
							<button type="submit" >
								<span class="glyphicon glyphicon-search"></span>
							</button>  
							
						</span>
					</div>
				</div>
				<input type="hidden" id="txtsort" name="txtsort" value="<%=varSortType%>">
			</form>
            </div>
        </div>
	</div>
	<div class="row">
		<form id="frmList" method="post" action="project_register.asp">
			<div class="table-responsive">	
				<div class="form-group" style="padding-left:15px">
					<button class="btn  btn-default btnNext" id="btnNew" type="button">Add New Project</button>
				</div>
				<table class="table table-hover" id="tblList">
					<thead class="thead-inverse tableheaderblue" >
						<tr>
							<%=HeaderSortColumn()%>
							<th>Department</th>
							<th>BDM</th>
							<th>Manager</th>
							<th>Status</th>
							<th></th>
							<th>Proposal</th>
							<th>Adwarded</th>
							<th>Invoice</th>
							<th >CSO</th>						
						</tr>
					</thead>
					<tbody>
						<%=strLast%>					
					</tbody>
					
				</table>	
				<input type="hidden" id="txthidden" name="txthidden" value="">
				<input type="hidden" name="txtpreviouspage" value="n_projectlist.asp">
			
			</div>		
		</form>		
    </div> 
	<div class="row">
		<div class="text-center">
			<%=Pagination()%>
		</div>
	</div>
<div>
    
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
<input type="hidden" name="txthidden" value="">
<input type="hidden" name="txtstatus" value="<%=strStatus%>">
<input type="hidden" name="P" value="<%=intCurPage%>">
<input type="hidden" name="S" value="<%=fgSort%>">
<input type="hidden" name="name" value="<%=strSName%>">
<input type="hidden" name="txtpreviouspage" value="<%=strFilename%>">

</form>

<script type="text/javascript" src="../../js/jquery-3.2.1.min.js"></script>

<script type="text/javascript" src="../../bootstrap/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="../../bootstrap/js/dataTables.bootstrap.min.js"></script>

<script type="text/javascript" src="../../js/library.js"></script>


<script language="javascript">
<!--
var objNewWindow;
var currentPage = <%=intPage%>;

$( document ).ready(function() { 
	

    $(".pagination li.pag_prev").click(function() {
        //if($(this).next().is('.active')) return;
        showPage(currentPage-1);
        
    });
	
	$(".pagination li.pag_next").click(function() {
        //if($(this).prev().is('.active')) return;
        showPage(currentPage+1);
    });
	
	$(".pagination li.numeros").click(function() {       
        showPage(parseInt($(this).text()));
    });
	
	
	$("#tblList tbody").on('click', 'tr td:not(:first-child) ', function (e) {	
		e.preventDefault();
		var status, proDate,marID
		var tds=$(this).parent().find('td');
		//alert(tds[0].innerHTML);
		status=$(this).parent().attr("status");
		proDate=tds[2].innerHTML.split("/");
		marID=tds[5].getAttribute("idValue");
		var temp=tds[0].innerHTML + ";" + status + ";" + proDate[1]+"/"+proDate[0]+ "/" + proDate[2]  + ";" + marID  ;
				
		$("#txthidden").val(temp);
		$("#frmList" ).submit(); 
        
    });
	
	$("#btnNew").click(function() {
        $("#txthidden").val('');		
		$("#frmList" ).submit(); 
       
    });
	
	$(".inv").click(function(e) {
		e.preventDefault();
		e.stopPropagation(); 
		
		var row = $(this).closest("tr");
		var tds=row.find('td');
		//alert(tds[0].innerHTML);
		
		$("#txthidden").val(tds[0].innerHTML);
		$('#frmList').attr("action","pro_invoice.asp");		
		$("#frmList" ).submit(); 
       
    });
	
	$(".cso").click(function(e) {
		e.preventDefault();
		e.stopPropagation(); 
		window.location=$(this).attr("path");
       
    });
	
	$(".sort-by").click(function(e) {
	
		e.preventDefault();
		var sortType="0";
		var FieldId = $(this).attr('class').replace('sort-by ', '');
		var curSortType = $(this).closest("th").attr('class');
		if (curSortType)
		{
			if (curSortType.indexOf("top") > -1) 
				sortType="1";
		}

		$("#txtsort").val(FieldId + sortType);
		
		$('#frmSearch').attr("action","n_projectlist.asp?s="+FieldId + sortType);
		$("#frmSearch" ).submit(); 
       
    });
});

function showPage(intpage){
	$('#frmSearch').attr("action","n_projectlist.asp?navi=" + intpage);
	$("#frmSearch" ).submit(); 
}

function myFunction() {
   var input, filter, table, tr, tdAPK, tdName, i;
  input = document.getElementById("txtSearch");
  filter = input.value.toUpperCase();
  table = document.getElementById("tblList");
  tr = table.getElementsByTagName("tr");

  // Loop through all table rows, and hide those who don't match the search query
  for (i = 0; i < tr.length; i++) {
    tdAPK = tr[i].getElementsByTagName("td")[0];
	tdName = tr[i].getElementsByTagName("td")[1];
    if (tdAPK && tdName) {
      if ((tdAPK.innerHTML.toUpperCase().indexOf(filter) > -1)||(tdName.innerHTML.toUpperCase().indexOf(filter) > -1)) {
        tr[i].style.display = "";
      } else {
        tr[i].style.display = "none";
      }
    } 
  }
}


//-->
</script>

</body>
</html>