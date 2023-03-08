<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->

<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim strPCCode,strSupplier,	dateBuying,strNote,strCPU,strRAM,strSysMem,strVideo,intStatus
	dim intPCID,intDetailPCID
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon,arrCategories,intCategoryType
	
	arrCategories =Array("Computer Status", "Computer Types", "Software Categories", "Licence Types")
	Dim SessionSharing
   Set SessionSharing = server.CreateObject("SessionMgr.Session2")
'***************************************************************
'
'***************************************************************
function OutBody(rsSrc)
	dim strOut
	dim i,j,subTotal
	dim grandTotal(15)
	
	strOut=""
	i=0
	For j=0 to ubound(grandTotal)-1
		grandTotal(i)=0
	next
	
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
					
			For j=0 to ubound(grandTotal)-1
				grandTotal(j)=grandTotal(j) + rsSrc.Fields(j+1)
			next
			
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' align='center' class='blue-normal'><a href='javascript:ViewData(" & rsSrc("year_") & ",0);' class='c'>" & rsSrc("year_") & "</a></td>"
			strOut=strOut & "<td valign='top'align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Direct"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",1);' class='c'>" &  rsSrc("Direct") & "</a>") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Indirect"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",2);' class='c'>" &  rsSrc("Indirect") & "</a>") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
            strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("TP"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",13);' class='c'>" &  rsSrc("TP") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"	
				
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Server"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",3);' class='c'>" &  rsSrc("Server") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("ATCTrainning"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",4);' class='c'>" &  rsSrc("ATCTrainning") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("MeetingRoom"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",5);' class='c'>" &  rsSrc("MeetingRoom") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"			
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Hotdesk"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",6);' class='c'>" &  rsSrc("Hotdesk") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"

            strOut = strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Laptop"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",7);' class='c'>" &  rsSrc("Laptop") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"			
					
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Stock"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",9);' class='c'>" &  rsSrc("Stock") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("OffStock"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",15);' class='c'>" &  rsSrc("OffStock") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Network"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",11);' class='c'>" &  rsSrc("Network") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"	
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Printer"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",12);' class='c'>" &  rsSrc("Printer") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"	
															
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("Other"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",10);' class='c'>" &  rsSrc("Other") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"																
			strOut=strOut & "<td valign='top' align='right' bgcolor='#D2DAEC' class='blue-normal'>" & rsSrc("total") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "<td valign='top' align='right' class='blue-normal'>" & _
								IIF(cint(rsSrc("WFH"))=0,"", "<a href='javascript:ViewData(" & rsSrc("year_") & ",14);' class='c'>" &  rsSrc("WFH") & "</a>") &_
								"&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			strOut=strOut & "</tr>"
			i=i+1
			rsSrc.MoveNext
		loop
		
		strOut=strOut & "<tr bgcolor='#D2DAEC'>"
		strOut=strOut & "<td valign='top' align='right' class='blue'>Grand Total</td>"
		
		dblTotal=0
		For j=0 to ubound(grandTotal)-1
			strOut=strOut & "<td valign='top' align='right' class='blue'>" &  IIF(grandTotal(j)>0,grandTotal(j),"") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			dblTotal=dblTotal + grandTotal(j)
		next
		
		'strOut=strOut & "<td valign='top' align='right' class='blue'>" & "" & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
		strOut=strOut & "</tr>"
		
	end if
	
	OutBody=strOut
End Function

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


'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------

	If Not checkSession(session("USERID")) Then
		Response.Redirect("../../message.htm")
	End If					
	SessionSharing("USERID") =  session("USERID")
	intUserID = session("USERID")
	print(session("USERID"))
	''SessionSharing("USERID") =  Session("USERID")
'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	intCategoryType=Request.Form("txtCategoryType")
	if intCategoryType="" then intCategoryType=0	
	'strTile=arrCategories(intCategoryType)
	
	strTile="Hardware Count by Age"
	
'--------------------------------------------------
' 
'--------------------------------------------------
'arrCategories =Array("Computer Status", "Computer Types", "Software Categories", "Licence Types")	
	'Select Case intCategoryType
	'	Case 0
	'	  strSql="SELECT StatusID as CategoryID,StatusDescription as CategoryName,StatusNote as CategoryNote,fgActivate as fgActivate FROM ATC_ComputerStatus"
	'	case 1
	'	  strSql="SELECT AtlasPCTypeID as CategoryID ,Description as CategoryName ,NoteType as CategoryNote,fgActivate as fgActivate FROM ATC_ComputerType"
	'	case 2
	'	  strSql="SELECT [SoftTypeID] as CategoryID ,[Description] as CategoryName ,[Note] as CategoryNote,fgActivate as fgActivate FROM [ATC_SoftwareType]"
	'	case 3
	'	  strSql="SELECT [LicenceTypeID] as CategoryID ,[LicenceTypeDescription] as CategoryName ,[LicenceTypeNote] as CategoryNote,fgActivate as fgActivate FROM ATC_SoftwareLicenceType"
	'end select
	
	
	strSql="SELECT * FROM IT_HardwareAgeStatistics ORDER BY Year_"

	Call GetRecordset(strSql,rsSrc)
	
	strLast=OutBody(rsSrc)	
	'strLast="aa"

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
	objEmployee.SetFullName(intStaffID)
	varFullName = split(objEmployee.GetFullName,";")
	strFullName = varFullName(0)					
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
	If strChoseMenu = "" Then strChoseMenu = "AF"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) Then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

	arrlstFrom(0) = selectmonth("lstmonthF",month(dateBuying) , -1)
	arrlstFrom(1) = selectday("lstdayF", day(dateBuying), -1)
	arrlstFrom(2) = selectyear("lstyearF", year(dateBuying), 1999, year(date())+2, 0)
		
	
	
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

Call ReadFromTemplateAll(arrPageTemplate, "../../templates/template1/", "ats_menu.htm")


arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
If arrPageTemplate(1)<>"" Then
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@menu", strMenu)
	arrTmp = split(arrPageTemplate(1), "@@content", -1)
	arrTmp(1) = Replace(arrTmp(1), "@@curpage", intCurPage)
	arrTmp(1) = Replace(arrTmp(1), "@@numpage", intTotalPage)	
End If
%>	

<html>
<head>
<title>Atlas Industries - Timesheet</title>

<link rel="stylesheet" type="text/css" href="../../library/DropDownMenu/anylinkcssmenu.css" />
<script type="text/javascript" src="../../library/DropDownMenu/anylinkcssmenu.js"></script>

<link rel="stylesheet" href="../../timesheet.css" type="text/css">
<script language="javascript" src="../../library/library.js"></script>

<script language="javascript">
<!--

anylinkcssmenu.init("anchorclass")

function ViewData(year,type)
{
	window.document.frmreport.year_.value = year
	window.document.frmreport.type_.value = type
	
	window.document.frmreport.action = "ComputerList.asp?act=out"			
	window.document.frmreport.submit();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmreport" method="post">
<%
'--------------------------------------------------
' Write the header of HTML page
'--------------------------------------------------

	Response.Write(arrPageTemplate(0))
	Response.Write(arrTmp(0))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
	      <tr> 
            <td> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
<%		If strError <> "" Then%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red" colspan="2">&nbsp;<b><%=strError%></b></td>
				</tr>
<%		End If%>				
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> </td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
                
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2"><%=strTile%></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign=top>
            
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
						  <table width="100%" border="0" cellspacing="1" cellpadding="5">
<%if strLast<>"" then %>						  
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center">Year</td>
                              <td class="blue" align="center" >Direct</td>                                  
                              <td class="blue" align="center" >Indirect</td>
                              <td class="blue" align="center" >TP</td>
							  
                              <td class="blue" align="center" >Server</td>
                              <td class="blue" align="center" >ATC</td>
                              <td class="blue" align="center" >Meeting<br>room</td>                                  
                              <td class="blue" align="center" >Hotdesk</td>
                              <td class="blue" align="center" >Laptop</td> 
                              <td class="blue" align="center" >Stock</td> 
							  <td class="blue" align="center" >Off-site<br>Stock</td>							  
                              <td class="blue" align="center">Network</td>
                              <td class="blue" align="center" >Printer</td>
                              <td class="blue" align="center" >Other</td>
                              <td class="blue" align="center">Total</td>
                              <td class="blue" align="center" >WFH</td>
                                
                            </tr>
<%Response.Write strLast%>
                          </table> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="20" class="blue-normal"> 
                                &nbsp;&nbsp;*Click on each number to view 
                                the list hardware with year by row and status by column</td>
                            </tr>
                          </table>  
<%end if%>
                          
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
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
<input type="hidden" name="year_" value="">
<input type="hidden" name="type_" value="">

<div id="submenu1" class="anylinkcss">
<ul>
<%for i=0 to UBound(arrCategories)%>
<li><a href="javascript:Category(<%=i%>)"><%=arrCategories(i)%></a></li>
<%next%>
</ul>
</div>

</form>

</body>
</html>
