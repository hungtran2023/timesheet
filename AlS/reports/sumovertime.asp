<!-- #include file = "../inc/constants.inc"-->
<!-- #include file = "../class/CEmployee.asp"-->
<!-- #include file = "../inc/createtemplate.inc"-->
<!-- #include file = "../inc/getmenu.asp"-->
<!-- #include file = "../inc/library.asp"-->
<%

'--------------------------------------------------
'
'--------------------------------------------------
Sub ResetArrayByInt(byref arr)
	for i = 0 to UBound(arr)
		arr(i)=0
	next
end sub

'--------------------------------------------------
'Get data from Timesheet
'--------------------------------------------------
Function GetOTHoursRecordset(dateStart,dateEnd)

	dim strConnect,objDatabase
	dim rs

	strConnect = Application("g_strConnect")	
	rs=null

	Set objDatabase = New clsDatabase
	If objDatabase.dbConnect(strConnect) Then
		objDatabase.cnDatabase.CursorLocation=adUseClient 
		
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objDatabase.cnDatabase
		myCmd.CommandType = adCmdStoredProc
		myCmd.CommandText = "GetDetailOfOT"
		

		Set myParam = myCmd.CreateParameter("start_date",adDate,adParamInput)
		myCmd.Parameters.Append myParam
		Set myParam = myCmd.CreateParameter("finish_date",adDate,adParamInput)
		myCmd.Parameters.Append myParam

		myCmd("start_date")=dateStart
		myCmd("finish_date")=dateEnd

		set rs=myCmd.Execute		
	end if

	set GetOTHoursRecordset = rs
end function
'--------------------------------------------------
'
'--------------------------------------------------
Function GenerateSumReportByStaff(rs)
	dim strReport,intNo,dblTotal,dblTotalAvailable
	dim dblSubtotal(5),dblGrandTotal(5)
	dim strName
	
	strReport=""

	call ResetArrayByInt(dblSubtotal)
	call ResetArrayByInt(dblGrandTotal)
	
		rs.MoveFirst
		
		intNo=1
		intI=0
				
		rs.Sort="Fullname"
  	    
  	    strName=""
		do while not rs.EOF 'and intI < rs.PageSize
 
		    if strName<> rs("IDNumber") then
		        if strName<>"" then
		            strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		                    "<td valign='top' colspan='4' class='blue' align='right'>Total (" & strName & "): </td>"
		            for jj=0 to UBound(dblSubtotal)-1
		                    strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
		            next       
		            strReport= strReport & "</tr>"
		            call ResetArrayByInt(dblSubtotal)
		        end if
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		        strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("IDNumber") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Fullname") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Jobtitle") & "</td>" 
		        strName= rs("IDNumber")
		    else
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		         strReport= strReport & "<td valign='top' class='blue-normal' colspan='3'>--</td>" 
		    end if
		    
		    dblTotal=0
    		
		    for j=0 to 3    		    
    		    dblTotal=dblTotal + CDbl(rs.Fields(6+j))  
    		    dblSubtotal(j)=dblSubtotal(j)+ CDbl(rs.Fields(6+j))
    		    dblGrandTotal(j)=dblGrandTotal(j)+ CDbl(rs.Fields(6+j))
		    next
            dblSubtotal(j)=dblSubtotal(j)+ dblTotal
    		dblGrandTotal(j)=dblGrandTotal(j)+ dblTotal
            
            strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("Project") & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue'align='right'>"& FormatNumber(dblTotal,1) & "</td>" & _
                    "</tr>" 
    			
			    intNo=intNo+1
			    rs.MoveNext
			    intI=intI+1
		loop

        strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		           "<td valign='top' colspan='4' class='blue' align='right'>Total: </td>"
        for jj=0 to UBound(dblSubtotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
        strReport= strReport & "<tr bgcolor='#FFE1E1'>" &_
       "<td valign='top' colspan='4' class='blue' align='right'>Overall Total:</td>"
        for jj=0 to UBound(dblGrandTotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblGrandTotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
	GenerateSumReportByStaff=strReport
End Function

'--------------------------------------------------
'
'--------------------------------------------------
Function GenerateSumReportByProject(rs)
	dim strReport,intNo,dblTotal,dblTotalAvailable
	dim dblSubtotal(5),dblGrandTotal(5)
	dim strName
	
	strReport=""

	call ResetArrayByInt(dblSubtotal)
	call ResetArrayByInt(dblGrandTotal)
	
		rs.MoveFirst
		
		intNo=1
		intI=0
				
		rs.Sort="Project"
  	    
  	    strName=""
		do while not rs.EOF 'and intI < rs.PageSize
 
		    if strName<> rs("Project") then
		        if strName<>"" then
		            strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		                    "<td valign='top' colspan='4' class='blue' align='right'>Total(" & strName &") </td>"
		            for jj=0 to UBound(dblSubtotal)-1
		                    strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
		            next       
		            strReport= strReport & "</tr>"
		            call ResetArrayByInt(dblSubtotal)
		        end if
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		         strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("Project") & "</td>"
		        strName= rs("Project")
		    else
		        strReport= strReport & "<tr bgcolor='#FFFFFF'>"
		         strReport= strReport & "<td valign='top' class='blue-normal'>--</td>" 
		    end if
		    
		    dblTotal=0
    		
		    for j=0 to 3    		    
    		    dblTotal=dblTotal + CDbl(rs.Fields(6+j))  
    		    dblSubtotal(j)=dblSubtotal(j)+ CDbl(rs.Fields(6+j))
    		    dblGrandTotal(j)=dblGrandTotal(j)+ CDbl(rs.Fields(6+j))
		    next
            dblSubtotal(j)=dblSubtotal(j)+ dblTotal
    		dblGrandTotal(j)=dblGrandTotal(j)+ dblTotal
    		
             strReport= strReport & "<td valign='top' class='blue-normal'>" & rs("IDNumber") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Fullname") & "</td>" & _
                  "<td valign='top' class='blue-normal'>" & rs("Jobtitle") & "</td>" & _
            
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("OTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNormal"),1) & "</td>" & _
                      "<td valign='top' class='blue-normal'align='right'>" & FormatNumber(rs("WeekendOTNight"),1) & "</td>" & _
                      "<td valign='top' class='blue'align='right'>"& FormatNumber(dblTotal,1) & "</td>" & _
                    "</tr>" 
    			
			    intNo=intNo+1
			    rs.MoveNext
			    intI=intI+1
		loop

        strReport= strReport & "<tr bgcolor='#FFF2F2'>" &_
		           "<td valign='top' colspan='4' class='blue' align='right'>Total: </td>"
        for jj=0 to UBound(dblSubtotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblSubtotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
        strReport= strReport & "<tr bgcolor='#FFE1E1'>" &_
       "<td valign='top' colspan='4' class='blue' align='right'>Overall Total:</td>"
        for jj=0 to UBound(dblGrandTotal)-1
                strReport= strReport & "<td valign='top' class='blue' align='right'>"& FormatNumber(dblGrandTotal(jj),1) &"</td>"
        next       
        strReport= strReport & "</tr>"
        
	GenerateSumReportByProject=strReport
End Function

'--------------------------------------------------
'End functions
'--------------------------------------------------

	Dim strUserName, varFullName, strTitle, strFunction, strMenu
	Dim objEmployee, objDb, gMessage
	dim intReportType

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../message.htm")
	End If
	
'-------------------------------
' Calculate pagesize
'-------------------------------

	If Not isEmpty(session("Preferences")) Then
			arrPre = session("Preferences")
			If arrPre(1, 0) > 0 Then intPageSize = arrPre(1, 0) Else intPageSize = PageSizeDefault
			Set arrPre = Nothing
	Else
		intPageSize = PageSizeDefault
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
				exit for
			end if
		next
		set getRight = nothing		
	end if	
	if fgRight = false then
		Response.Redirect("../welcome.asp")
	end if
	
'----------------------------------
' Get Full Name and Job Title
'----------------------------------

	Set objEmployee = New clsEmployee
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	objEmployee.SetFullName(session("USERID"))
	varFullName = split(objEmployee.GetFullName,";")
	strTitle = "<b>" & varFullName(0) & "</b>&nbsp;" & varFullName(1)

	strtmp1 = Replace(preferences, "XX", session("strHTTP"))
	strtmp2 = Replace(logoff, "XX", session("strHTTP"))
	strFunction = "<div align='right'><a href='../welcome.asp?choose_menu=B' class='c' onMouseOver='self.status=&quot;Return Main menu&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Main Menu</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				"<a class='c' href='javascript:_print();' onMouseOver='self.status=&quot;Print report&quot;; return true;' onMouseOut='self.status=&quot;&quot;'>Print</a>&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				strtmp1 & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../images/dot.gif' width='5' height='5'>" &_
				"&nbsp;&nbsp;&nbsp" & strtmp2 & "&nbsp;&nbsp;&nbsp;</div>"
	Set objEmployee = Nothing

'--------------------------------------------------
' Preparing data
'--------------------------------------------------

    if Request.Form("optReportType")="" then
	    intReportType=1
    else
	    intReportType = cint(Request.Form("optReportType"))
    end if
    
    intUserStyle=Request.Form("lbUserStyle")
    if intUserStyle="" then lbUserStyle=0

    strType = "1"
	IF Request.Form("opttime")<>"" then strType = Request.Form("opttime")
		
	If strType = "1" Then
	
		intMonth = Month(Date)		
		If Request.Form("lbmonth") <> "" Then intMonth = CInt(Request.Form("lbmonth"))
		intYear = Year(Date)
		If Request.Form("lbyear") <> "" Then intYear	= Request.Form("lbyear")
		strTitle2	= SayMonth(intMonth) & " - " & intYear
		
		strFrom=Cdate(intMonth & "/1/" & intYear)	
		strTo=DateAdd("m",1,strFrom)-1
	Else
	
		strFrom1	= Request.Form("txtfrom")
		strTo1		= Request.Form("txtto")
		strTitle2	= "From " & strFrom1 & " To " & strTo1
	
		varFrom		= split(strFrom1,"/")
		strFrom		= CDate(varFrom(1) & "/" & varFrom(0) & "/" & varFrom(2))
		varTo		= split(strTo1,"/")
		strTo		= CDate(varTo(1) & "/" & varTo(0) & "/" & varTo(2))			
	End If

'--------------------------------------------------
' Get current page
'--------------------------------------------------

	intCurPage = trim(Request.Form("P"))
	
	If intCurPage = "" Or Request.QueryString("act") = "vra" Then
		intCurPage = 1
	End If	

'--------------------------------------------------
' Analyse query and prepare report
'--------------------------------------------------

	If Request.QueryString("act") = "" Or Request.QueryString("act") = "vra" Then
	
		set rsSumHours=	GetOTHoursRecordset(strFrom,strTo)
		
		if rsSumHours.recordcount>0 then
			
			rsSumHours.PageSize = intPageSize    ' So ban ghi tren mot trang
			intTotalPage = rsSumHours.PageCount       ' Tong so trang						
			
			If Not IsEmpty(session("varSumHours")) Then session("varSumHours") = Empty
			set session("varSumHours") = rsSumHours
		end if	
	Else
		
		set rsSumHours=session("varSumHours")
		
		If Request.QueryString("act") = "vpa1" Then
			intCurPage = Request.Form("txtpage")
		ElseIf Request.QueryString("act") = "vpa2" Then
			intCurPage = CInt(intCurPage) - 1
		ElseIf Request.QueryString("act") = "vpa3" Then	
			intCurPage = CInt(intCurPage) + 1
		End If	
	End If
'--------------------------------------------------
' Generate Report
'--------------------------------------------------	
    strLast=""
    
	if rsSumHours.recordcount>0 then
	    rsSumHours.filter=""
	    
	    if intUserStyle>0 then
	        rsSumHours.filter="UserType=" & intUserStyle
	    end if
	     
	    if intReportType=2 then
	        strLast=GenerateSumReportByProject(rsSumHours)
	    else
	        strLast=GenerateSumReportByStaff(rsSumHours)	
	    end if
	end if
	
	'Session("StrLast")=strLast
		
	Set objDatabase = Nothing
'--------------------------------------------------
' End of preparing report
'--------------------------------------------------
'--------------------------------------------------
' Read template page from file
'--------------------------------------------------

If isEmpty(session("arrInfoCompany")) Then
		strConnect = Application("g_strConnect") 
		Set objDb = New clsDatabase
		If objDb.dbConnect(strConnect) Then

			strQuery = "select a.CompanyName, isnull(Address,'') Address, isnull(City,'') City, isnull(b.CountryName,'') Country, " &_
						"isnull(Phone,'') Phone, isnull(Fax,'') Fax, isnull(c.Logo,'') Logo from ATC_Companies a " &_
						"left join ATC_Countries b On a.CountryID = b.CountryID " &_
						"left join ATC_CompanyProfile c ON a.CompanyID = c.CompanyID " &_
						"where a.CompanyID = " & session("Inhouse")

			If objDb.runQuery(strQuery) Then
				If not objDb.noRecord Then
					arrInfoCompany = objDb.rsElement.getRows
					session("arrInfoCompany") = arrInfoCompany
					objDb.closerec
				End If
			Else
				strError = objDb.strMessage
			End If
			objDb.dbDisconnect
		Else
			strError = objDb.strMessage
		End If
		Set objDb = Nothing
	End If
	
Call ReadFromTemplateAll(arrPageTemplate, "../templates/template1/", "ats_report.htm")

arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@title", strTitle)
arrPageTemplate(0) = Replace(arrPageTemplate(0),"@@function", strFunction)
if not isEmpty(session("arrInfoCompany")) then
	arrTmp = session("arrInfoCompany")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", arrTmp(0, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", arrTmp(1, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", arrTmp(2, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", arrTmp(3, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", arrTmp(4, 0))
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", arrTmp(5, 0))
	if arrTmp(6, 0)<>"" then
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "<img src='../images/" & arrTmp(6, 0) & "' border='0'>" )
	else
		arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;" )
	end if
	set arrTmp = nothing
else
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cname", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Caddress", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccity", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Ccountry", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cphone", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Cfax", "&nbsp;")
	arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@Clogo", "&nbsp;")
end if
%>	

<html>
<head>
<title>Atlas Industries Time Sheet System</title>

<link href="../jQuery/jquery-ui.css" rel="stylesheet" type="text/css"/>
<link href="../jQuery/atlasJquery.css" rel="stylesheet" type="text/css"/>
<link href="../timesheet.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../jQuery/jquery.min.js"></script>
<script type="text/javascript" src="../jQuery/jquery-ui.min.js"></script>
<script type="text/javascript" src="../library/library.js"></script>

<script type="text/javascript">
    $(function() {
        var dates = $("#txtfrom, #txtto").datepicker({
            dateFormat: "dd/mm/yy",
            minDate: "1/1/2000",
            maxDate: "31/12/<%=year(Date())%>",
            onSelect: function(selectedDate) {
                var option = this.id == "txtfrom" ? "minDate" : "maxDate",
                                        instance = $(this).data("datepicker"),
                                        date = $.datepicker.parseDate(
                                            instance.settings.dateFormat ||
                                            $.datepicker._defaults.dateFormat,
                                            selectedDate, instance.settings);
                dates.not(this).datepicker("option", option, date);
            }
        });
    });

var objWindowSumPro;

    function _print() { //v2.0
        var str1 = "<%=strfromto%>";
        str1 = escape(str1);
        var str2 = "<%=strprintdate%>";
        str2 = escape(str2);
        //var fgprint = <%=session("NumPageSumPro")%>;
        var fgprint = 1;

        if (fgprint != 0) {
            window.status = "";
            strFeatures = "top=" + (screen.height / 2 - 275) + ",left=" + (screen.width / 2 - 390) + ",width=900,height=550,toolbar=no,"
	            + "menubar=yes,location=no,directories=no,scrollbars=yes,status=yes";
            if ((objWindowSumPro) && (!objWindowSumPro.closed)) {
                objWindowSumPro.focus();

            } else {
            objWindowSumPro = window.open("p_sumovertime.asp?fromto=" + str1 + "&printdate=" + str2 + "&type=" + <%=intReportType %>, "MyNewWindow", strFeatures);
            }
            window.status = "Opened a new browser window.";
        }
        else
            alert("No data for your request.")
    }

function window_onunload() {
	if((objWindowSumPro) && (!objWindowSumPro.closed))
		objWindowSumPro.close();
}

function checkdata()
{
	if (document.frmreport.opttime[0].checked)
	{
		if (isnull(document.frmreport.txtfrom.value)==true)
		{
			alert("Please enter startdate before click here.")
			document.frmreport.txtfrom.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtfrom.value)==false)
			{			
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtfrom.focus();
				return false;
			}
		}
		
		if (isnull(document.frmreport.txtto.value)==true)
		{
			alert("Please enter enddate before click here.")
			document.frmreport.txtto.focus();
			return false;
		}
		else
		{
			if (isdate(document.frmreport.txtto.value)==false)
			{
				alert("This value is invalid. \n Please use the following format: 'dd/mm/yyyy'");
				document.frmreport.txtto.focus();
				return false;
			}
		}
		
		if (comparedate(document.frmreport.txtfrom.value,document.frmreport.txtto.value)==false)
		{
			alert("The startdate must be less than the finishdate.")
			document.frmreport.txtfrom.focus();
			return false;
		}
	}	
	return true;
}

function _submit() {
	if(checkdata()==true) {
	    document.frmreport.action = "sumovertime.asp?act=vra";
		document.frmreport.target = "_self" ;
		document.frmreport.submit();
	}
}


</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="frmreport" method="post">
    		<%
			'--------------------------------------------------
			' Write the header of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(0))
			%>
<table style="height:auto; margin:0 auto; width:100%">
  <tr> 
    <td bgcolor="#FFFFFF" valign="top"> 
    	
  <table width="780" border="0" cellspacing="0" cellpadding="0" align="center" >
    <tr> 
      <td bgcolor="#FFFFFF" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr> 
            <td width="31%"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td valign="top" width="0%" class="blue" height="33">&nbsp;</td>
                  <td width="4%" class="blue" align="right" height="33"> 
                    <input type="radio" name="opttime" value="0" <%if strType="0" then%>checked<%end if%>/>
                  </td>
                  <td class="blue-normal" width="8%" height="33">From </td>
                  <td class="blue-normal" width="40%" height="33"> 
                    <input type="text" name="txtfrom" id="txtfrom" class="blue-normal"  size="5" style="width:75" value="<%=strFrom1%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                  <td width="8%" class="blue-normal" height="33"> To </td>
                  <td class="blue-normal" height="33" width="40%"> 
                    <input type="text" name="txtto" id="txtto" class="blue-normal" size="5" style="width:75" value="<%=strto1%>" onClick="document.frmreport.opttime[0].checked=true;">
                  </td>
                </tr>
                <tr> 
                  <td valign="top" width="0%" class="blue">&nbsp;</td>
                  <td width="4%" class="blue" align="right"> 
                    <input type="radio" name="opttime" value="1" <%if strType="1" then%>checked<%end if%>>
                  </td>
                  <td class="blue-normal" width="8%">Month </td>
                  <td class="blue-normal" width="40%"> 
                        <select name="lbmonth"  style='width:70px; height: 24px;' class="blue-normal" language="javascript" onClick="document.frmreport.opttime[1].checked=true">
					    <%for iM=1 to 12%>
				            <option <%If CInt(intMonth)=iM Then%>selected<%End If%> value="<%=iM%>"><%=MonthName(iM,true)%></option>
				        <%next%>
				  </select>
                  </td>
                  <td width="8%" class="blue-normal"> Year </td>
                  <td class="blue-normal" width="40%">
				  <select name="lbyear" size="1" style='width:70px; height: 24px;'  class="blue-normal" language="javascript" onClick="document.frmreport.opttime[1].checked=true">
				<%For ii=2000 To Year(Date)%>
				    <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
				<%Next%>
				  </select>
                  </td>
                </tr>
              </table>
            </td>
            <td width="20%" valign="top"> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
                  <td height="33" class="blue-normal">&nbsp;</td>
                  <td height="33" class="blue-normal" valign="bottom"> 
                    <input type="radio" name="optReportType" value="1" <%if intReportType=1 then Response.Write "checked"%>>
                    By staff 
                    <input type="radio" name="optReportType" value="2" <%if intReportType=2 then Response.Write "checked"%>>
                    By project</td>
                </tr>			
                <tr>
                  <td height="33" class="blue-normal" valign="middle" colspan="2">
                       <select name="lbUserStyle"  style='width:90%; height: 24px;' class="blue-normal">
				            <option  value="0" <% if intUserStyle=0 then%> selected <%end if %> >View All</option>
				            <option value="1" <% if intUserStyle=1 then%> selected <%end if %> >Atlas Staff</option>
				            <option value="3" <% if intUserStyle=3 then%> selected <%end if %> >Contract Staff</option>
				        
				  </select></td>
                </tr>		
              </table>
            </td>
            <td width="39%">     &nbsp;    
                    <table width="60" border="0" cellspacing="0" cellpadding="0" height="20" name="aa">
                        <tr> 
                          <td class="blue" align="center" bgcolor="8CA0D1" onMouseOver="this.style.backgroundColor='7791D1';" onMouseOut="this.style.backgroundColor='8CA0D1';" height="20" > 
                              <a href="javascript:_submit();" class="b">Submit</a>
                          </td>
                        </tr>
                        </table> 
              </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%if gMessage<>"" then%>
          <tr bgcolor="#E7EBF5">
			<td class="red" height="20" align="left" width="100%"> &nbsp;&nbsp;<b><%=gMessage%></b></td>
		  </tr>
		  <%end if %>
          <tr> 
            <td bgcolor="8CA0D1"><img src="../IMAGES/DOT-01.GIF" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
          </tr>
        </table>
        </td>
</tr>
</table>    
</td>
</tr>
  <tr> 
    <td style="vertical-align:top; text-align: center">      
        <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
            <td>
    		<%
			'--------------------------------------------------
			' Write the title of report page
			'--------------------------------------------------
			arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@titleofreport", "Summary of Overtime")
	        arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@fromto", strTitle2)
	        arrPageTemplate(1) = Replace(arrPageTemplate(1),"@@printdate", formatdatetime(date,vbLongDate))
	        Response.Write(arrPageTemplate(1))
			%>
			</td></tr></table>
      <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td bgcolor="#617DC0"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                <tr> 
<%if intReportType=1 then%>                
                   <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">&nbsp;StaffID</td>
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Full Name</td>
                  <td class="blue" align="center" width="14%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Jobtitle</td>
                  <td class="blue" align="center" width="29%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Project Name</td>
                  
<%else %>                  
                  <td class="blue" align="center" width="29%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Project Name</td>
                  <td class="blue" align="center" width="10%" bgcolor="#E7EBF5" rowspan="2">&nbsp;StaffID</td>
                  <td class="blue" align="center" width="15%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Full Name</td>
                  <td class="blue" align="center" width="14%" bgcolor="#E7EBF5" rowspan="2">&nbsp;Jobtitle</td>
               
<%end if %>

                  
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5" colspan="2">Nomal day</td>
                  <td class="blue" align="center" width="12%" bgcolor="#E7EBF5" colspan="2">Weekend</td>
                  <td class="blue" align="center" width="8%" bgcolor="#E7EBF5" rowspan="2">Total</td>                  
                </tr>
                <tr> 
                  <td class="blue" align="center" bgcolor="#E7EBF5">Normal</td>
                  <td class="blue" align="center" bgcolor="#E7EBF5">Night</td>
                  <td class="blue" align="center" bgcolor="#E7EBF5">Normal</td>
                  <td class="blue" align="center" bgcolor="#E7EBF5">Night</td>
                </tr>     
              
               
<%Response.Write strLast%>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  
<%if session("NumPageSumPro")>0 then%>
  <table width="780" border="0" cellspacing="0" cellpadding="0" height="20" align="center">
    <tr> 
      <td align="right" bgcolor="#E7EBF5"> 
        <table width="70%" border="0" cellspacing="1" cellpadding="0" height="18">
          <tr> 
            <td align="right" valign="middle" width="37%" class="blue-normal">Page 
            </td>
            <td align="center" valign="middle" width="13%" class="blue-normal"> 
              <input type="text" name="txtpage" class="blue-normal" value="<%=session("CurPageSumPro")%>" size="2" style="width:50">
            </td>
            <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();"  onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
            </td>
            <td align="right" valign="middle" width="15%" class="blue-normal">Pages 
               <%=session("CurpageSumPro")%>/<%=session("NumpageSumPro")%>&nbsp;&nbsp;</td>
            <td valign="middle" align="right" width="28%" class="blue-normal"><a href="javascript:prev();"  onMouseOver="self.status='Previous page'; return true;" onMouseOut="self.status='';">Previous</a> 
              /<a href="javascript:next();" onMouseOver="self.status='Next page'; return true;" onMouseOut="self.status='';"> Next</a>&nbsp;&nbsp;&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
</table>
<%end if%>
			<%
			'--------------------------------------------------
			' Write the footer of HTML page
			'--------------------------------------------------
			Response.Write(arrPageTemplate(2))    
			%>
</form>
</body>
</html>