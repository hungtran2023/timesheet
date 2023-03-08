<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%

Dim strUserid, rsSrc
dim dblCountryID, intStayfor,intEntries,strLastdate
dim strFrom1,strFrom2,strFrom3,strFrom4
dim strTo1,strTo2,strTo3,strTo4,strNote,dblVisaID


function PopulateListVISA
	
	Dim strSql,strReturn,i
	Dim objDatabase, rs
	Dim strConnect, strID,  strError,strEntryType, strLastdateTemp
	
	strSql="SELECT VisaID , StaffID , a.CountryID , StayforMonth , EntryType , LastDateArrive , ActualFrom1 , ActualTo1 , " &_
	        " ActualFrom2 , ActualTo2 , ActualFrom3 , ActualTo3 , ActualFrom4 , ActualTo4 , VisaNote ,CountryName " &_
            " FROM ATC_ReplacementHistory a INNER JOIN ATC_Countries b ON a.CountryID=b.CountryID " &_
            " WHERE StaffID=" & strUserid
    
    Call GetRecordset(strSql,rs)  
    strReturn=""
    i=1
    if not rs.Eof then
	  Do Until rs.EOF
	      strID=rs("VisaID")
	      
	      if cdbl(strID)=cdbl(dblVisaID) then
                dblCountryID=rs("CountryID")
                intStayfor=rs("StayforMonth")
                intEntries=rs("EntryType")


                strLastdate=rs("LastDateArrive")
                if strLastdate<>"" then strLastdate=day(strLastdate) & "/" & month(strLastdate) & "/" & year(strLastdate) 

                strFrom1=rs("ActualFrom1")
                if strFrom1<>"" then strFrom1=day(strFrom1) & "/" & month(strFrom1) & "/" & year(strFrom1) 
                strFrom2=rs("ActualFrom2")
                if strFrom2<>"" then strFrom2=day(strFrom2) & "/" & month(strFrom2) & "/" & year(strFrom2)
                strFrom3=rs("ActualFrom3")
                if strFrom3<>"" then strFrom3=day(strFrom3) & "/" & month(strFrom3) & "/" & year(strFrom3)
                strFrom4=rs("ActualFrom4")
                if strFrom4<>"" then strFrom4=day(strFrom4) & "/" & month(strFrom4) & "/" & year(strFrom4)

                strTo1=rs("ActualTo1")
                if strTo1<>"" then strTo1=day(strTo1) & "/" & month(strTo1) & "/" & year(strTo1) 
                strTo2=rs("ActualTo2")
                if strTo2<>"" then strTo2=day(strTo2) & "/" & month(strTo2) & "/" & year(strTo2) 
                strTo3=rs("ActualTo3")
                if strTo3<>"" then strTo3=day(strTo3) & "/" & month(strTo3) & "/" & year(strTo3) 
                strTo4=rs("ActualTo4")
                if strTo4<>"" then strTo4=day(strTo4) & "/" & month(strTo4) & "/" & year(strTo4) 

                strNote=rs("VisaNote")
	      end if
	      
	      strEntryType="Single"
	      if rs("EntryType") then strEntryType="Multiple"
	      strLastdateTemp=rs("LastDateArrive")
          if strLastdateTemp<>"" then strLastdateTemp=day(strLastdateTemp) & "/" & month(strLastdateTemp) & "/" & year(strLastdateTemp) 
                
	      strReturn= strReturn & "<tr idValue='"  & strID & "'><td>" & i & ".</td>"
          strReturn= strReturn & "<td class='editrow'>" & rs("CountryName") & "</td>"
          strReturn= strReturn & "<td class='editrow'>" & strEntryType & "</td> "
          strReturn= strReturn & "<td class='editrow'>" & strLastdateTemp & "</td> "
          strReturn= strReturn & "<td class='editrow'>" & rs("StayforMonth") & "</td> "
          strReturn= strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item' data-id='" & rs("VisaID")  & "'></button></td>  </tr>"		
	    rs.MoveNext
	    i=i+1
	  Loop       
	end if         
	
	PopulateListVISA=strReturn
end function

'--------------------------------------------------
' Check session variable if it was expired or not
'--------------------------------------------------
	If checkSession(session("USERID")) = False Then
		Response.Redirect("../../message.htm")
	End If
'-----------------------------------
'Check ACCESS right
'-----------------------------------

	tmp = Request.Form("txtpreviouspage")
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
	fgRight=true
	if fgRight = false then
		Response.Redirect("../../welcome.asp")
	end if	
'--------------------------------------------------

'--------------------------------------------------
strStatus= Request.Form("txtstatus")
strUserid= Request.querystring("id")
if strUserid="" then strUserid=Request.Form("txtuserid")

strAct= Request.querystring("act")


if strAct="u" or strAct="d" then

    if strAct="u" then
        dblCountryID=request.form("lstCountry")
        intStayfor=request.form("txtStayfor")
        intEntries=Request.Form("radEntries")

        strLastdate=request.form("txtLastdate")
        if strLastdate<>"" then strLastdate=ConvertTommddyyyy(strLastdate)
        
        strFrom1=request.form("txtFrom1")
        if strFrom1<>"" then strFrom1=ConvertTommddyyyy(strFrom1)
        strFrom2=request.form("txtFrom2")
        if strFrom2<>"" then strFrom2=ConvertTommddyyyy(strFrom2)
        strFrom3=request.form("txtFrom3")
        if strFrom3<>"" then strFrom3=ConvertTommddyyyy(strFrom3)
        strFrom4=request.form("txtFrom4")
        if strFrom4<>"" then strFrom4=ConvertTommddyyyy(strFrom4)
        
        strTo1=request.form("txtTo1")
        if strTo1<>"" then strTo1=ConvertTommddyyyy(strTo1)
        strTo2=request.form("txtTo2")
        if strTo2<>"" then strTo2=ConvertTommddyyyy(strTo2)
        strTo3=request.form("txtTo3")
        if strTo3<>"" then strTo3=ConvertTommddyyyy(strTo3)
        strTo4=request.form("txtTo4")
        if strTo4<>"" then strTo4=ConvertTommddyyyy(strTo4)
        
        strNote=request.form("txtNote")
        
        dblVisaID=request.form("txtVisaID")
        
        if dblVisaID="" then dblVisaID=-1   
          
        if cint(dblVisaID)=-1 then
            strSql="INSERT INTO ATC_ReplacementHistory  (StaffID,CountryID,StayforMonth ,EntryType ,LastDateArrive ,ActualFrom1 ,ActualTo1 ,ActualFrom2 ,ActualTo2 ," & _
                    "ActualFrom3 ,ActualTo3 ,ActualFrom4 ,ActualTo4 ,VisaNote) VALUES( " & _
                    strUserid & "," & dblCountryID & "," & intStayfor & "," & intEntries & "," & IIF(strLastdate="","null","'" & strLastdate & "'") & "," & _
                    IIF(strFrom1="","null","'" & strFrom1 & "'") & "," & IIF(strTo1="","null","'" & strTo1 & "'") & "," & _
                    IIF(strFrom2="","null","'" & strFrom2 & "'") & "," & IIF(strTo2="","null","'" & strTo2 & "'") & "," & _
                    IIF(strFrom3="","null","'" & strFrom3 & "'") & "," & IIF(strTo3="","null","'" & strTo3 & "'") & "," & _
                    IIF(strFrom4="","null","'" & strFrom4 & "'") & "," & IIF(strTo4="","null","'" & strTo4 & "'") & "," & _                
                    IIF(strNote="", "null","'" & strNote & "'") & ")"
        else
            strSql="UPDATE  ATC_ReplacementHistory  " & _
                       "SET  CountryID  = " & dblCountryID & _
                          ", StayforMonth  = " & intStayfor & _
                          ", EntryType  = " & intEntries & _
                          ", LastDateArrive  =" & IIF(strLastdate="","null","'" & strLastdate & "'") & _
                          ", ActualFrom1  = " & IIF(strFrom1="","null","'" & strFrom1 & "'") & _
                          ", ActualTo1  = " & IIF(strTo1="","null","'" & strTo1 & "'") & _
                          ", ActualFrom2  = " & IIF(strFrom2="","null","'" & strFrom2 & "'") & _
                          ", ActualTo2  = " & IIF(strTo2="","null","'" & strTo2 & "'") & _
                          ", ActualFrom3  = " & IIF(strFrom3="","null","'" & strFrom3 & "'") & _
                          ", ActualTo3  = " & IIF(strTo3="","null","'" & strTo3 & "'") & _
                          ", ActualFrom4  = " & IIF(strFrom4="","null","'" & strFrom4 & "'") & _
                          ", ActualTo4  = " & IIF(strTo4="","null","'" & strTo4 & "'") & _
                          ", VisaNote  = " & IIF(strNote="","null","'" & strNote & "'") & _
                     " WHERE VisaID=" & dblVisaID
        end if
    else
        strSql="DELETE FROM ATC_ReplacementHistory WHERE VisaID=" & Request.querystring("subid")
    end if
	
	
	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase     
    strError=""
    
	If objDatabase.dbConnect(strCnn) Then              
			if not objDatabase.runActionQuery(strSql) then 
			   gMessage = objDatabase.strMessage
            end if			  
        
    end if
    dblVisaID=-1
    dblCountryID=-1
    intStayfor=""
    intEntries=false

    strLastdate=""    
    strFrom1=""
    strFrom2=""
    strFrom3=""
    strFrom4=""
    
    strTo1=""
    strTo2=""
    strTo3=""
    strTo4=""
    
    strNote=""
    
else
    dblVisaID=strAct
end if

 strListVISA=PopulateListVISA()
'--------------------------------------------------
' Initialize recordset
'--------------------------------------------------	
        if dblCountryID="" then dblCountryID=-1
		strSql="SELECT CountryID, CountryName FROM [ATC_Countries] WHERE fgActivate=1 ORDER BY CountryName"	
		Call GetRecordset(strSql,rsDepart)
	    strCountryList= PopulateDataToListWithoutSelectTag(rsDepart,"CountryID", "CountryName", cdbl(dblCountryID))	    
	    
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
				"<img src='../../images/dot.gif' width='5' height='5'>&nbsp;&nbsp;&nbsp;" &_
				help & "&nbsp;&nbsp;&nbsp;<img src='../../images/dot.gif' width='5' height='5'>" &_
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
	if strChoseMenu = "" then strChoseMenu = "AB"
	
	'current URL without name of site and query string
	strLink = Request.ServerVariables("URL") 
	strLink = Mid(strLink , Instr(2, strLink, "/") + 1, Len(strLink))
	strFullName = varFullName(0)
	If IsEmpty(Session("strHTTP")) then Call MakeHTTP
	strMenu = getMenuTMS(getRes, strURL, strChoseMenu, strLink, strFullName, "../../")

    
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

    <link href="../../css/bootstrap.min.css" rel="stylesheet" type="text/css">
    <link href="../../css/timesheet.css" rel="stylesheet" >

    <link href="../../css/atlasJquery.css" rel="stylesheet" type="text/css" />
    <link href="../../css/style.css" rel="stylesheet" type="text/css">
    <link href="../../css/datepicker.css" rel="stylesheet" type="text/css">

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
<!--Content-->			

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tbody>
    <tr> 
        <td style="padding:20px 20px 20px 15px;"> 
        <div class="navi-info"> 
                <a class="blue" href="listofemployee.asp" onMouseOver="self.status='Show the list of employees'; return true;" onMouseOut="self.status=''">Employee List:</a>
            <span>Employee Details</span>
            </div>
        </td>
    </tr>
</tbody>
</table>

<div class="container-fluid">
    <div class="row">
        <div class="col-sm-12">
            <ul class="nav nav-tabs">
                <li><a href="employeeProfile.asp?id=<%=strUserid%>">Employee Profile</a></li>
                <li><a href="atlasinformation.asp?id=<%=strUserid%>">Atlas Information</a></li>
                <li><a href="educationskill.asp?id=<%=strUserid%>">Education/Skill</a></li>
                <li class="active"><a>Replacement History</a></li>
                <li><a href="employmenthistory.asp?id=<%=strUserid%>">Employment History</a></li>
            </ul>
        </div>
    </div>
    <div class="row">
        <div class="col-sm-12">
            <div class="tab-content employee-details-form">
               <!--- start employee profile -->
                <div class="row">
                    <div class="col-md-12 col-sm-6 col-xs-12">
                        <form  id='frmReplacement' class="form-horizontal row-border" method="POST" action="replacementhistory.asp?id=<%=strUserid%>&act=u">
<%if gMessage<>"" then%>   	                
                       
                            <div class="alert alert-danger">
                                <strong>Error:</strong> <%=gMessage%>
                            </div>
<%end if%>                            
                            <div class="panel panel-default">
                                <div class="panel-body">
                                    <div class="col-md-6">
                                        <div class="form-group">
                                            <label class="col-md-12">Offshore Visa Country</label>
                                            <div class="col-md-12">
                                                <select id="lstCountry" name="lstCountry" class="form-control" >
                                                    <option value=""></option>
                                                    <%=strCountryList%>
                                                </select>                                                
                                            </div>
                                        </div>
                                        
                                        <div class="form-group">
                                            <label class="col-md-12">Entries</label>
                                            <div class="col-md-11 border-radio">                                                
                                                    <label class="col-md-4 radio-inline" >
                                                    <input type="radio" name="radEntries" id="radSingle" value="0" <% if not intEntries then%>checked<%end if%> >Single</label>
                                                    <label class="radio-inline">
                                                    <input type="radio" name="radEntries" id="radMultiple" value="1" <% if intEntries then%>checked<%end if%> >Multiple</label>                                                
                                            </div>
                                            <div id="radMessage" class="col-md-12" style="margin-bottom:0px;"></div>                        
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Actual Replacement</label>
                                        </div>
                                        <div class="form-group">                                            
                                            <div class="col-md-12">
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">From</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtFrom1" name="txtFrom1" class="form-control datepicker inp-apply-from-date" placeholder="DD/MM/YYYY" value="<%=strFrom1%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                         
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">To</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtTo1" name="txtTo1" class="form-control datepicker inp-apply-from-date"  placeholder="DD/MM/YYYY" value="<%=strTo1%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                            </div> 
                                        </div>
                                        <div class="form-group">    
                                            <div class="col-md-12">
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">From</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtFrom2" name="txtFrom2" class="form-control datepicker inp-apply-from-date" placeholder="DD/MM/YYYY" value="<%=strFrom2%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">To</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtTo2" name="txtTo2" class="form-control datepicker inp-apply-from-date"  placeholder="DD/MM/YYYY" value="<%=strTo2%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                         <div class="form-group">
                                            <div class="col-md-12">
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">From</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtFrom3" name="txtFrom3" class="form-control datepicker inp-apply-from-date" placeholder="DD/MM/YYYY" value="<%=strFrom3%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">To</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtTo3" name="txtTo3" class="form-control datepicker inp-apply-from-date"  placeholder="DD/MM/YYYY" value="<%=strTo3%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                            </div>
                                         </div>
                                         <div class="form-group">
                                        
                                            <div class="col-md-12">
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">From</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtFrom4" name="txtFrom4" class="form-control datepicker inp-apply-from-date" placeholder="DD/MM/YYYY" value="<%=strFrom4%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                                <div class="col-md-6">
                                                    <label class="col-md-2 control-label no-blod">To</label>
                                                    <div class="input-group date">
                                                        <input type="text" id="txtTo4" name="txtTo4" class="form-control datepicker inp-apply-from-date"  placeholder="DD/MM/YYYY" value="<%=strTo4%>">
                                                        <span class="input-group-addon">
                                                            <span class="ic-calendar"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-md-6">
                                        <div class="form-group">
                                            <label class="col-md-12">Stay for / until (months)</label>
                                            <div class="col-md-12">
                                                <input  type="number"  id="txtStayfor" name="txtStayfor" class="form-control" value="<%=intStayfor%>">
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Last date to Arrive</label>
                                            <div class="col-md-12">
                                                <div class="input-group date">
                                                    <input type="text"  id="txtLastdate" name="txtLastdate"  class="form-control datepicker" placeholder="MM/DD/YYYY" value="<%=strLastdate%>">
                                                    <span class="input-group-addon">
                                                        <span class="ic-calendar"></span>
                                                    </span>
                                                </div>
                                            </div>                                            
                                        </div>
                                        <div class="form-group">
                                            <label class="col-md-12">Note</label>
                                            <div class="col-md-12">
                                                <textarea id="txtNote" name="txtNote"  cols="30" rows="10" class="form-control"><%=strNote%></textarea>
                                            </div>
                                            
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <input type="hidden" name="txtVisaID" id="txtVisaID" value="<%=dblVisaID%>"/>
                            
                            <div class="col-sm-12">
                                <div class="form-group text-right">
                                    <div class="col-md-12 no-padding">
                                        <button type="submit" class="btn btn-primary" id="btnSave">Save</button>
                                        <button type="button" class="btn btn-primary <%if cdbl(dblVisaID)<=0 then%> hide<%end if%>" id="btnCancel">Cancel</button>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-12 no-padding">
                                <table class="table table-striped table-bordered table-hover margin-bottom10 table-responsive" id="tblList">
                                    <thead class="thead-inverse">
                                        <tr>
                                            <th class="col-md-1">No.</th>
                                            <th>Country</th>
                                            <th>Entries</th>
                                            <th>Last date to Arrive</th>
                                            <th>Stay for / until</th>
                                            <th class="col-action"></th>
                                            
                                        </tr>
                                    </thead>
                                    <tbody>                                        
                                        <%=strListVISA%>
                                    </tbody>
                                </table>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
<!-- Modal for displaying the messages -->
<div class="modal fade" id="confirm-delete" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
				<h4 class="modal-title" id="myModalLabel">Confirm Delete</h4>
			</div>		
			<div class="modal-body">
				<p id="modal_message"></p>
				<p>Do you want to proceed?</p>
			</div>			
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				<a class="btn btn-danger btn-ok">Delete</a>
			</div>
        </div>
    </div>
</div>      
               
<%
Response.Write(arrTmp(1))
'--------------------------------------------------
' Write the footer of HTML page
'--------------------------------------------------
Response.Write(arrPageTemplate(2))    
%>

<script type="text/javascript" src="../../js/jquery.min.js"></script>
<script type="text/javascript" src="../../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../../js/library.js"></script>
<script type="text/javascript" src="../../js/bootstrap-datepicker.js" charset="UTF-8"></script>
<script type="text/javascript" src="../../js/bootstrap-table.js"></script>
<script type="text/javascript" src="../../js/js-control.js"></script>
<script type="text/javascript" src="../../js/formValidation.min.js"></script>
<script type="text/javascript" src="../../js/framework/bootstrap.min.js"></script>

<script type="text/javascript">

$(document).ready(function() {

    $('#txtLastdate').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtLastdate');
        });


    $('#txtFrom1').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtFrom1');
        });


    $('#txtFrom2').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtFrom2');
        });
        
        
    $('#txtFrom3').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtFrom3');
        });
        
        
    $('#txtFrom4').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtFrom4');
        });
               
   
    $('#txtTo1').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtTo1');
        });


    $('#txtTo2').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtTo2');
        });
        
        
    $('#txtTo3').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtTo3');
        });
        
        
    $('#txtTo4').on('changeDate', function(e) {
            // Revalidate the date field
            $('#frmReplacement').formValidation('revalidateField', 'txtTo4');
        });
        
        
   
    $('#frmReplacement').formValidation({
            framework: 'bootstrap',
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },        
            fields: {
                lstCountry:{
                    validators: {
                        notEmpty: {
                            message: 'The Country is required.'
                        }
                    }
                },
                radEntries: {
                    err: '#radMessage',
                    validators: {
                        notEmpty: {
                            message: 'The gender is required.'
                        }
                    }
                },
                txtStayfor: {
                    validators: {
                        notEmpty: {
                            message: 'The number of months is required and must be a number'
                        },
                        between: {
                                min: 0,
                                max:300,
                                message: 'The number of months must be greater than 0'
                        }
                    }
                },
                txtLastdate:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The Last date to Arrive is not a valid'
                        }
                    }  
                 },         
                 txtFrom1:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },
                 txtFrom2:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },  
                 txtFrom3:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },  
                 txtFrom4:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },         
                 txtTo1:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },
                 txtTo2:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },  
                 
                 txtTo3:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 },  
                 txtTo4:{
                    validators: {
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The From date is not a valid'
                        }
                    }  
                 }
           }
            
      });
      
     
       $(".editrow").on('click', function(e) {
       
            var res = $(this).parent().attr("idValue");
            window.location = 'replacementhistory.asp?id=<%=strUserid%>&act='+ res  ; // redirect
       });
        
       $("#btnCancel").click(function(){
       
            $("#txtVisaID").val(-1);
            $("#lstCountry").val("");
            $("#txtStayfor").val("");
            $("#txtLastdate").val("");
            $("#txtNote").val("");
            $("#txtFrom1").val("");
            $("#txtFrom2").val("");
            $("#txtFrom3").val("");
            $("#txtFrom4").val("");
            $("#txtTo1").val("");
            $("#txtTo2").val("");
            $("#txtTo3").val("");
            $("#txtTo4").val("");
            $("#radSingle").attr('checked', 'checked');
            $(this).addClass('hide');
        });
        
        
        $('.btn-remove-item').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to remove this Visa.");
            $('#confirm-delete').modal('show');
            
             $('#confirm-delete').find('.btn-ok').attr('href', 'replacementhistory.asp?id=<%=strUserid%>&act=d&subid='+id );
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        });      
  
});

</script>

</body>
</html>

