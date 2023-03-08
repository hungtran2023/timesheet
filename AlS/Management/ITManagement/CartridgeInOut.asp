<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file = "../../inc/library.asp"-->


<%
	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	dim strPCCode,strSupplier,	dateBuying,strNote,strCPU,strRAM,strHDD,strVideo,intStatus, strSN
	dim intPCID,intDetailPCID
	Dim objEmployee, objDatabase, strError,rsData
	Dim arrlstFrom(2),arrlongmon, strArrayDisk1,strArrayDisk2
	Dim arrDisk1, arrDisk2

'***************************************************************
'
'***************************************************************
function OutBody(rsSrc)
	dim strOut
	dim i

	i=0
	if (rsSrc.RecordCount>0) then	
		rsSrc.MoveFirst
		Do while not rsSrc.EOF
			

			if cint(intInkStockID)=cint(rsSrc("InkStockID")) then
				intQty=abs(rsSrc("Qty") )
				intPrinterID=rsSrc("PrinterID") 
				
				dateStock=rsSrc("StockDate")
				strNote=rsSrc("StockNote") 
			end if

			
			strColor = "#FFF2F2"
			if i mod 2 = 0 then	strColor = "#E7EBF5"
			
			strOut=strOut & "<tr bgcolor='" & strColor & "'>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & i+1 & "</td>"
			strOut=strOut & "<td valign='top' class='blue'>" & _
						"<a href='javascript:InkStock(" & rsSrc("InkStockID") & ");' " &_
						"class='c'>" & day(rsSrc("StockDate")) & "/" & Month(rsSrc("StockDate")) & "/" & year(rsSrc("StockDate")) & "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & Abs(rsSrc("Qty")) &  "</td>"
			strOut=strOut & "<td valign='top' class='blue-normal'>" & rsSrc("StockNote") & "</td>"		
			strOut=strOut & "</tr>"
			i=i+1	
			rsSrc.MoveNext
		loop
		
	end if
	
	OutBody=strOut
End Function

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

	'strConnect = Application("g_strConnect")
	'Set objDatabase = New clsDatabase
	
	intCartridgeID = Request.Form("txtID")
	intInOut= Request.Form("txtStt")
	fgDel=Request.Form("fgstatus")
	intInkStockID=Request.Form("txtInkStockID")
	if intInkStockID="" then intInkStockID=-1
	
	
	strSql="SELECT CartridgeCode FROM ATC_Cartridges  WHERE [CartridgeID]=" & intCartridgeID
	Call GetRecordset(strSql,rsData)
	
	strCartridgeCode=rsData("CartridgeCode")
	
	if Request.QueryString("act") = "save" then

		
		intQty=Request.Form("txtQty")
		
		intPrinterID="null"
		if cint(intInOut)=1 then
			intPrinterID=Request.Form("lstPrinter")	
			if cint(intPrinterID)=-1 then intPrinterID="null"
			intQty=(-1) * intQty
		end if
		
		dateStock=cdate(Request.Form("lstMonthF") & "/" & Request.Form("lstDayF") & "/" & Request.Form("lstYearF"))
		
		strNote=Request.Form("txtNote")
		

		if fgDel<>"D" then
			
			if Cint(intInkStockID)=-1 then
				'Add new				
				strSQL="INSERT INTO ATC_InkStock(CartridgeID ,Qty,StockDate,isAdded,PrinterID,StockNote,StaffID) VALUES ( " &_
						intCartridgeID & "," & intQty & ",'" & dateStock & "'," & intInOut & "," & intPrinterID & "," & IIF(strNote="","NULL","'" & strNote & "'" ) & "," & intUserID & ")"
			else
				
				'Update
				
				strSQL="UPDATE ATC_InkStock " & _
							"SET   Qty =" &  intQty  & _
							  ",StockDate = '" & dateStock & "'" & _
							  ",PrinterID = " & intPrinterID  & _
							  ",StockNote = " & IIF(trim(strNote)="","NULL","'" & strNote & "'" )  & _
							  ",StaffID = " & intUserID & _
						" WHERE InkStockID=" & intInkStockID
			end if
		else
		
			strSQL="DELETE FROM ATC_InkStock WHERE InkStockID=" & intInkStockID
			fgDel=""
						
		end if
'response.write strSQL
		strCnn = Application("g_strConnect")	
		Set objDatabase = New clsDatabase     
		strError=""
		
		If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strSQL) then 
			   strError = objDatabase.strMessage
			Else
				strError = "Updated successfully"
			end if			  
		end if		
	End If
'--------------------------------------------------
' 
'--------------------------------------------------
	strOut=""
	intQty=""
	intPrinterID=-1
	dateStock=Date()
	strNote=""
	strSql ="SELECT a.InkStockID, a.CartridgeID, a.Qty, a.StockDate, a.isAdded, ISNULL(a.PrinterID,-1) as PrinterID, a.StockNote, a.StaffID FROM ATC_InkStock a WHERE CartridgeID=" & intCartridgeID & " AND isAdded=" & intInOut

	Call GetRecordset(strSql,rsSrc)
	
	strLast=OutBody(rsSrc)
	
	strSql="SELECT * FROM ATC_Printers WHERE fgActivate=1"
	Call GetRecordset(strSql,rsPrinter)

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

	arrlstFrom(0) = selectmonth("lstmonthF",month(dateStock) , -1)
	arrlstFrom(1) = selectday("lstdayF", day(dateStock), -1)
	arrlstFrom(2) = selectyear("lstyearF", year(dateStock), 1999, year(date())+2, 0)
	
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

<link rel="stylesheet" href="../../timesheet.css" type="text/css">

<script type="text/javascript" src="../../library/library.js"></script>
<link href="../../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>

<script type="text/javascript" src="../../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../../jQuery/jquery-ui.min.js"></script>

<link href="../../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>

<script language="javascript">
<!--
	
	
	 function InkStock(id) {
        window.document.frmreport.txtInkStockID.value = id
        window.document.frmreport.action = "CartridgeInOut.asp?act=edit"
        window.document.frmreport.submit();
    }

    function savedata() {
        if (checkdata()) {
            window.document.frmreport.action = "CartridgeInOut.asp?act=save"
            window.document.frmreport.submit();
        }
    }

    function deletedata()
{
	var answer = confirm("Do you want to remove this item?")
	if (answer){
		window.document.frmreport.fgstatus.value = "D"
		window.document.frmreport.action = "CartridgeInOut.asp?act=save"
        window.document.frmreport.submit();
	}	
}

    function checkdata() {
        if (window.document.frmreport.txtQty.value == "") {
            alert("Please enter Qty of CartridgeList.");
            document.frmreport.txtQty.focus();
            return false
        }
		
		if (isNaN(window.document.frmreport.txtQty.value) == true) {
            alert("Invalid value for qty.");
            document.frmreport.txtQty.focus();
            return false
        }

        var dateFrom = document.frmreport.lstdayF.value + "/" + document.frmreport.lstmonthF.value + "/" + document.frmreport.lstyearF.value

        if (isdate(dateFrom) == false) {
            alert("The date (" + dateFrom + ") is invalid.");
            document.frmwh.lstdayF.focus();
            return false;
        }

        return true
    }
	
	function Add() {
		
		document.frmreport.txtQty.value = "";
		document.frmreport.txtNote.value = "";
		document.frmreport.txtInkStockID.value=-1;
		<%if cint(intInOut)=1 then%>
		document.frmreport.lstPrinter.selectedIndex=-1;
		<%end if%>
		document.frmreport.lstdayF.selectedIndex = <%=day(Date())-1%>;
		document.frmreport.lstmonthF.selectedIndex = <%=Month(Date())-1%>;
		document.frmreport.lstYearF.selectedIndex = <%=Year(Date())-2000%>;
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
              <table width="70%" border="0" cellpadding="0" cellspacing="0">
<%		If strError <> "" Then%>               
				<tr bgcolor="#E7EBF5">
				  <td class="red" colspan="2">&nbsp;<b><%=strError%></b></td>
				</tr>
<%		End If%>				
                <tr align="center"> 
                  <td class="blue" height="10" align="left" width="23%"> &nbsp;&nbsp;<a href="CartridgeList.asp" onMouseOver="self.status='';return true">Cartridge List</a></td>
                  <td class="blue" height="30" align="right" width="77%"></td>
                </tr>
				
                <tr align="center"> 
                  <td class="title" height="50" align="center" colspan="2">Cartridge <%if cint(intInOut)=0 then%>In<%else%>Out<%end if%></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="100%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="height:&quot;79%&quot;" height="365">
                <tr> 
                  <td bgcolor="#FFFFFF" valign="top"> 
                    <table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#617DC0"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="5">
							
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" width="30%" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" width="15%"></td>
                              <td valign="middle" width="35%" class="blue">
								<span style="font-size:14px"><%=strCartridgeCode%></span>
								<input type="hidden" name="txtID" value="<%=intCartridgeID%>"></td>
                              <td valign="top" width="20%" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
                            
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal">Date  <%if cint(intInOut)=0 then%>In<%else%>Out<%end if%></td>
                              <td valign="middle" class="blue"><%
														Response.Write arrlstFrom(1)
														Response.Write arrlstFrom(0)
														Response.Write arrlstFrom(2)%></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr> 
                            <tr bgcolor="#FFFFFF"> 
                              <td valign="top" class="blue">&nbsp;</td>
                              <td valign="middle" class="blue-normal" >Qty *</td>
                              <td valign="middle"  class="blue">
								<input type="text" name="txtQty" maxlength="20" class="blue-normal" style="width:95%;" value="<%=intQty%>"></td>
                              <td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
<%if cint(intInOut)=1 then%>
							<tr bgcolor="#FFFFFF"> 
								<td valign="top" class="blue">&nbsp;</td>
								<td valign="middle" class="blue-normal" >Printer </td>
								<td valign="middle"  class="blue">
									<select name='lstPrinter' style="width:95%;"  class='blue-normal'>
										<option value='-1' <%if cint(intPrinterID)=-1 then%>selected<%end if%>>&nbsp; </option>					
										<%if rsPrinter.RecordCount>0 then
											rsPrinter.MoveFirst
											Do while not rsPrinter.EOF%>
											
												<option value='<%=rsPrinter("PrinterID")%>' <%if cint(intPrinterID)=cint(rsPrinter("PrinterID")) then%>selected<%end if%>><%=rsPrinter("PrinterName")%></option>
												
										<%		rsPrinter.MoveNext
											loop
										end if%>
									</select></td>
								<td valign="top" class="blue-normal" align="center">&nbsp;</td>
                            </tr>
<%end if%>
                            <tr bgcolor="#FFFFFF"> 
								<td valign="top" class="blue">&nbsp;</td>
								<td valign="middle" class="blue-normal">Note</td>
								<td valign="middle" class="blue">
									<textarea name="txtNote" rows="4" cols="40" class="blue-normal"><%=strNote%></textarea> 
								</td>
								<td valign="top" class="blue-normal" align="center"><input type="hidden" name="txtInkStockID" value="<%=intInkStockID%>"></td>                                
                            </tr>                                                     
                            
                       
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
							<td class="blue" height="20" align="left">&nbsp;&nbsp;
								<a href="javascript:Add();" onMouseOver="self.status='Add'; return true;" onMouseOut="self.status=''">Add New</a></td>
							</tr>
                          </table>
<%if strLast<>"" then %>                          
						  <table width="100%" border="0" cellspacing="1" cellpadding="5">
                            <tr bgcolor="#8CA0D1"> 
                              <td class="blue" bgcolor="#8CA0D1" align="center" width="10%">No.</td>
                              <td class="blue" align="center" width="30%">Date</td>  
                              <td class="blue" align="center" width="30%">Qty</td>  
                              <td class="blue" align="center" width="30%">Note</td>
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
<input type="hidden" name="txtStt" value=<%=intInOut%>>
<input type="hidden" name="fgstatus" value="">

</form>

</body>
</html>
