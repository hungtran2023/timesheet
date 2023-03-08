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
	
	arrCategories =Array("Computer Status", "Computer Types", "Software Categories", "Installation Types", "Licence Types","Monitor Brand","Monitor Size", "Printer", "Cartridge", "Accessories")
	
'***************************************************************
'
'***************************************************************
function OutBody(rsSrc)
	dim strOut
	dim i
	
	strOut=""
	i=0
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
			
			strActivate="<img src='../../images/yes.gif'>"
			if not rsSrc("fgActivate") then strActivate=""
			
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & i+1 & "</td>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:CategoryDetail(" & rsSrc("CategoryID") & ");' " &_
						"class='c'>" & rsSrc("CategoryName") & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("CategoryNote") & "</td>"
			strOut=strOut & "<td valign='top' align='center' class='blue-normal'>" & strActivate & "</td>"
			strOut=strOut & "</tr>"
			i=i+1
			rsSrc.MoveNext
		loop
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

	intUserID = session("USERID")

'--------------------------------------------------
' Initialize variables
'--------------------------------------------------
	intCategoryType=Request.Form("txtCategoryType")
	if intCategoryType="" then intCategoryType=0	
	strTile=arrCategories(intCategoryType)
	
		
	intPCID = Request.Form("txtID")
	fgDel=Request.Form("fgstatus")

	if Request.QueryString("act") = "save" then
		intCategoryID=Request.Form("txtCategoryID")
		
		if fgDel<>"D" then
			strName=Request.Form("txtName")
			strNote=Request.Form("txtNote")
			if intCategoryType=8 then 
				intAlmostEmpty=Request.Form("txtAlmostEmpty")
				if intAlmostEmpty="" then intAlmostEmpty=0
			end if
			fgActivate=Request.Form("optActivate")			
			if fgActivate="" then fgActivate=0
				
			if cint(intCategoryID)=-1 then
				Select Case intCategoryType
					Case 0
					  strSql="INSERT INTO ATC_ComputerStatus (StatusDescription,StatusNote ,fgActivate) VALUES('" & _
								Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
					case 1
					  strSql= "INSERT INTO ATC_ComputerType (Description ,NoteType ,fgActivate) VALUES('" & _
								Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
					case 2
					  strSql="INSERT INTO ATC_SoftwareType(Description,Note,fgActivate) VALUES('" & _
								Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
					case 3
					  strSql="INSERT INTO ATC_SoftwareLicenceType(LicenceTypeDescription,LicenceTypeNote,fgActivate) VALUES('" & _
								Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
                    case 4
					  strSql="INSERT INTO ATC_LicenceTypes(TypeName,NoteType,fgActivated) VALUES('" & _
                            Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
                    case 5
					  strSql="INSERT INTO ATC_MonitorBrands(BrandName,BrandNote,fgActivate) VALUES('" & _
                            Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
                    case 6
					  strSql="INSERT INTO ATC_MonitorSize(SizeDetail,SizeNote,fgActivate) VALUES('" & _
                            Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"
					case 7 
						strSql="INSERT INTO ATC_Printers (PrinterName,Description,fgActivate)  VALUES   ('" & _
							Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"							
				    case 8 
						strSql="INSERT INTO ATC_Cartridges (CartridgeCode,CartridgeNote,AlmostEmpty,fgActivate)  VALUES   ('" & _
							Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & intAlmostEmpty & "," & fgActivate & ")"		
					 case 9 
						strSql="INSERT INTO [ATC_ITAccessories] ([Description],[NoteAccessories],[fgActivate]) VALUES  ('" & _
							Replace(strName,"'","''") & "'," & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & "," & fgActivate & ")"					
				end select
				strPrefix="Add new "

			else
				Select Case intCategoryType
					Case 0
					  strSql="UPDATE ATC_ComputerStatus SET " & _
								"StatusDescription = '" & Replace(strName,"'","''") & "'" & _
								",StatusNote = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
							" WHERE StatusID=" & intCategoryID
					case 1				  
					  strSql="UPDATE ATC_ComputerType SET " & _
								"Description = '" & Replace(strName,"'","''") & "'" & _
								",NoteType = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
								" WHERE AtlasPCTypeID= " & intCategoryID
					case 2
					
					  strSql="UPDATE ATC_SoftwareType SET " & _
								"Description = '" & Replace(strName,"'","''") & "'" & _
								",Note = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
								" WHERE SoftTypeID= " & intCategoryID
					case 3
					  strSql="UPDATE ATC_SoftwareLicenceType SET " & _
								"LicenceTypeDescription = '" & Replace(strName,"'","''") & "'" & _
								",LicenceTypeNote = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
							" WHERE LicenceTypeID= " & intCategoryID
                    case 4
					  strSql="UPDATE ATC_LicenceTypes SET " & _
								"TypeName = '" & Replace(strName,"'","''") & "'" & _
								",NoteType = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivated = " & fgActivate & _
							" WHERE LicenceTypeID= " & intCategoryID							
				    case 5
					  strSql="UPDATE ATC_MonitorBrands SET " & _
								"BrandName = '" & Replace(strName,"'","''") & "'" & _
								",BrandNote = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
							" WHERE BrandID= " & intCategoryID	
					case 6
					  strSql="UPDATE ATC_MonitorSize SET " & _
								"SizeDetail = '" & Replace(strName,"'","''") & "'" & _
								",SizeNote = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
							" WHERE SizeID= " & intCategoryID	
					case 7
					  strSql="UPDATE ATC_Printers SET " & _
								"PrinterName = '" & Replace(strName,"'","''") & "'" & _
								",Description = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
							" WHERE PrinterID= " & intCategoryID	
					case 8
					  strSql="UPDATE ATC_Cartridges SET " & _
								"CartridgeCode = '" & Replace(strName,"'","''") & "'" & _
								",CartridgeNote = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",AlmostEmpty=" & intAlmostEmpty & _
								",fgActivate = " & fgActivate & _
							" WHERE CartridgeID= " & intCategoryID	
					case 9
					  strSql="UPDATE [ATC_ITAccessories] SET " & _
								"[Description] = '" & Replace(strName,"'","''") & "'" & _
								",[NoteAccessories] = " & IIF(strNote="","NULL","'" & Replace(strNote,"'","''") & "'") & _
								",fgActivate = " & fgActivate & _
							" WHERE AccessoriesID= " & intCategoryID							
				end select
				strPrefix="Update "
			end if
		else
			Select Case intCategoryType
				Case 0
				  strSql="DELETE ATC_ComputerStatus WHERE StatusID=" & intCategoryID
				case 1				  
				  strSql="DELETE ATC_ComputerStatus WHERE AtlasPCTypeID= " & intCategoryID
				case 2					
				  strSql="DELETE ATC_SoftwareType WHERE SoftTypeID= " & intCategoryID
				case 3
				  strSql="DELETE ATC_SoftwareLicenceType WHERE LicenceTypeID= " & intCategoryID
                case 4
				  strSql="DELETE ATC_LicenceTypes WHERE LicenceTypeID= " & intCategoryID	
                case 5
				  strSql="DELETE ATC_MonitorBrands WHERE BrandID= " & intCategoryID
                case 6
				  strSql="DELETE ATC_MonitorSize WHERE SizeID= " & intCategoryID		
                case 7
				  strSql="DELETE ATC_Printers WHERE PrinterID= " & intCategoryID	
                case 8
				  strSql="DELETE ATC_Cartridges WHERE CartridgeID= " & intCategoryID	
				case 9
				  strSql="DELETE ATC_ITAccessories WHERE AccessoriesID= " & intCategoryID

			end select
			strPrefix="Delete "
			fgDel=""
		end if
		strError=ExecuteSQL(strSql,strPrefix)		
	else
		intCategoryID=-1
		strName=""
		strNote=""
		fgActivate=false
	End If
	
'--------------------------------------------------
' 
'--------------------------------------------------
'arrCategories =Array("Computer Status", "Computer Types", "Software Categories", "Licence Types")	
	Select Case intCategoryType
		Case 0
		  strSql="SELECT StatusID as CategoryID,StatusDescription as CategoryName,StatusNote as CategoryNote,fgActivate as fgActivate FROM ATC_ComputerStatus"
		case 1
		  strSql="SELECT AtlasPCTypeID as CategoryID ,Description as CategoryName ,NoteType as CategoryNote,fgActivate as fgActivate FROM ATC_ComputerType"
		case 2
		  strSql="SELECT [SoftTypeID] as CategoryID ,[Description] as CategoryName ,[Note] as CategoryNote,fgActivate as fgActivate FROM [ATC_SoftwareType]"
		case 3
		  strSql="SELECT [LicenceTypeID] as CategoryID ,[LicenceTypeDescription] as CategoryName ,[LicenceTypeNote] as CategoryNote,fgActivate as fgActivate FROM ATC_SoftwareLicenceType"
        case 4
		  strSql="SELECT [LicenceTypeID] as CategoryID ,[TypeName] as CategoryName ,[NoteType] as CategoryNote,fgActivated as fgActivate FROM ATC_LicenceTypes"		
		case 5
		  strSql="SELECT [BrandID] as CategoryID ,[BrandName] as CategoryName ,[BrandNote] as CategoryNote,fgActivate as fgActivate FROM ATC_MonitorBrands"
        case 6
		  strSql="SELECT [SizeID] as CategoryID ,[SizeDetail] as CategoryName ,[SizeNote] as CategoryNote,fgActivate as fgActivate FROM ATC_MonitorSize"				   
		case 7
		  strSql="SELECT [PrinterID] as CategoryID ,[PrinterName] as CategoryName ,[Description] as CategoryNote,fgActivate as fgActivate FROM [ATC_Printers]"				   
		case 8
		  strSql="SELECT [CartridgeID] as CategoryID ,[CartridgeCode] as CategoryName ,[CartridgeNote] as CategoryNote,fgActivate as fgActivate, AlmostEmpty FROM [ATC_Cartridges]"	
		 case 9
		  strSql="SELECT [AccessoriesID] as CategoryID,[Description] as CategoryName,[NoteAccessories] as CategoryNote,[fgActivate] as fgActivate  FROM [ATC_ITAccessories] "
		end select
	
	Call GetRecordset(strSql,rsSrc)
	
	strLast=OutBody(rsSrc)	
	
	If Request.QueryString("act") = "EDIT" Then			
		if rsSrc.RecordCount>0 then
		
			intCategoryID=Request.Form("txtCategoryID")
			rsSrc.Filter = "CategoryID=" & intCategoryID
			
			strName=rsSrc("CategoryName")
			strNote=rsSrc("CategoryNote")
			fgActivate=rsSrc("fgActivate")
			if intCategoryType=8 then intAlmostEmpty=rsSrc("AlmostEmpty")
		end if			
	end if
	
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

function savedata()
{
	if (checkdata())
	{
		window.document.frmreport.action = "Categories.asp?act=save"			
		window.document.frmreport.submit();
	}
}

function Category(type)
{
	window.document.frmreport.txtCategoryType.value = type
	window.document.frmreport.action = "Categories.asp"			
	window.document.frmreport.submit();
}

function CategoryDetail(id)
{
	window.document.frmreport.txtCategoryID.value = id
	window.document.frmreport.action = "Categories.asp?act=EDIT"			
	window.document.frmreport.submit();
}

function deletedata()
{
	window.document.frmreport.fgstatus.value = "D"
	window.document.frmreport.action = "Categories.asp?act=save"
	window.document.frmreport.submit();
}

function checkdata()
{
	if (window.document.frmreport.txtName.value=="")
	{
		alert("Please enter <%=strTile%> name.");
		document.frmreport.txtName.focus();
		return false	
	}	
	
	return true	
}

function Add() {
	document.frmreport.txtName.value = "";
	document.frmreport.txtNote.value = "";
	document.frmreport.optActivate.checked=false
	document.frmreport.txtCategoryID.value=-1
	<%if intCategoryType=8 then%>
	document.frmreport.txtAlmostEmpty.value = "";
	<%end if%>
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
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;</td>
                  <td class="blue" height="30" align="right" width="77%">
					<table width="120" border="0" cellspacing="2" cellpadding="0" align="right" height="20" name="aa">
                      <tr> 
                       
                        <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20">
                          <div align="center" class="blue">
							<a href="javascript:Category(0)" class="anchorclass" rel="submenu1">Select Category</a>							
                        </td>						                     
                      </tr>                    
                    </table>
                  </td>
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
                          <table width="100%" border="0" cellspacing="0" cellpadding="2">
                                 
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="20%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%">Name *</td>
                              <td valign="middle" width="45%" class="blue">
								<input type="text" name="txtName" maxlength="100"  class="blue-normal" style="width:95%;" value="<%=strName%>"></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>  
<%if intCategoryType=8 then%>							
							<tr bgcolor="#FFFFFF"> 
                              <td valign="top"  class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" ">Almost empty </td>
                              <td valign="middle"  class="blue">
								<input type="text" name="txtAlmostEmpty" maxlength="20" class="blue-normal" style="width:95%;" value="<%if cint(intAlmostEmpty)>0 then response.write(intAlmostEmpty) end if%>"</td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr> 							
<%end if%>							
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Note</td>
                              <td valign="middle" class="blue">
                              
                              <TEXTAREA NAME="txtNote" ROWS=2 style="width:95%" class="blue-normal"><%=strNote%></TEXTAREA>

							</td>
                             <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr>  
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal"></td>
                              <td valign="middle" class="blue">
								<input type="checkbox" name="optActivate" value="1" <%if fgActivate then%>checked<%end if%>>Activate</td>
                             <td valign="top" class="blue-normal" align="center">&nbsp;</td>                                
                            </tr> 

                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr> 
                              <td height="50"> 
                                <table width="120" border="0" cellspacing="2" cellpadding="0" align="center" height="20" name="aa">
                                  <tr> 
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:savedata()" onMouseOver="self.status='Please click here to save changes';return true" onMouseOut="self.status='';return true" class="b">Save</a></div>
                                    </td>
                                    <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" height="20" width="60">
                                      <div align="center" class="blue"><a href="javascript:deletedata()" onMouseOver="self.status='Please click here to delete this record';return true" onMouseOut="self.status='';return true" class="b">Delete</a></div>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
							<tr> 
							  <td class="blue" height="20" align="left">&nbsp;&nbsp;
								<a href="javascript:Add();" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add New</a></td>
							  </tr>
                          </table>

<%if strLast<>"" then %>	
						  <table width="100%" border="0" cellspacing="1" cellpadding="5">
					  
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="10%">No.</td>
                              <td class="blue" align="center" width="35%">Name</td>  
                              <td class="blue" align="center" width="40%">Note</td>  
                              <td class="blue" align="center" width="15%">Activate</td>  
                            </tr>
<%Response.Write strLast%>
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
<input type="hidden" name="fgstatus" value="<%=fgDel%>">
<input type="hidden" name="txtCategoryType" value="<%=intCategoryType%>">
<input type="hidden" name="txtCategoryID" value="<%=intCategoryID%>">

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
