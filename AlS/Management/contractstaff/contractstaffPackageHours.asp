<!-- #include file = "../../class/CEmployee.asp"-->
<!-- #include file = "../../inc/createtemplate.inc"-->
<!-- #include file = "../../inc/getmenu.asp"-->
<!-- #include file = "../../inc/constants.inc"-->
<!-- #include file="../../class/clsSHA-1.asp" -->
<!-- #include file = "../../inc/library.asp"-->
<%

function PopulateList
	
	Dim strSql,strReturn,i
	Dim objDatabase, rs
	Dim strConnect, strID,  strError,strEntryType, strdateTemp

    strSql="SELECT ProjectID,SubTaskName, TDate,[Hours],Note,a.AssignmentID FROM (SELECT * FROM ATC_Timesheet UNION ALL SELECT * FROM ATC_Timesheet" & year(Date())-1 & ") a " & _
				"INNER JOIN ATC_Assignments b ON a.AssignmentID=b.AssignmentID " & _
				"INNER JOIN ATC_Tasks c ON b.SubTaskID=c.SubTaskID " & _
			"WHERE a.StaffID=" & strUserid & " ORDER By Tdate DESC"

    Call GetRecordset(strSql,rs)  
    strReturn=""
    i=1
    if not rs.Eof then
	  Do Until rs.EOF
     
          strdateTemp= day(rs("TDate")) & "/" & month(rs("TDate")) & "/" & year(rs("TDate")) 
          
	      strReturn= strReturn & "<tr idValue='"  & rs("AssignmentID") & "'>"
          strReturn= strReturn & "<td class='editrow'>" & rs("ProjectID") & "</td>"
          strReturn= strReturn & "<td class='editrow'>" & rs("SubTaskName") & "</td> "
          strReturn= strReturn & "<td class='editrow idate' iDate='" & rs("TDate") & "'>" &  strdateTemp &  "</td> "
		  strReturn= strReturn & "<td class='editrow iHours'>" &  rs("Hours") &  "</td> "
		  strReturn= strReturn & "<td class='editrow iNote'>" &  rs("Note") &  "</td> "
          strReturn= strReturn & "<td class='col-sm-1 col-action text-center'><button class='btn-remove-item' data-id='" & rs("AssignmentID") & "**" & rs("TDate")  & "'></button></td>  </tr>"		
	    rs.MoveNext
	    i=i+1
	  Loop       
	end if         
	
	PopulateList=strReturn
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
strPriKey= Request.Form("txtPriKey")
'response.write 	strPriKey

'if strStatus="" then strStatus="N"
strUserid= Request.querystring("id")
if strUserid="" then strUserid=Request.Form("txtuserid")

strAPKSearch=Request.querystring("s")
strAct=Request.Querystring("act")



if strAct<>"" then
	
	strAssignment=request.form("lstTasks")
	strDate=request.form("txtDate")
	strHours=request.form("txtHours")
	strNote=request.form("txtNote")

	strTimeSheetTable="ATC_Timesheet"
	if Year(strDate)<>Year(Date()) then strTimeSheetTable=strTimeSheetTable& year(strDate)

	
	if strAct="u" then
		If strPriKey ="" then
			strSql="INSERT INTO " & strTimeSheetTable & " (TDate,StaffID,AssignmentID,EventID,[Hours],OverTime,Note) VALUES (" & _
					"'" & ConvertTommddyyyy(strDate) & "'," & strUserid & "," & strAssignment & ",1," & strHours & ",0," & _
					IIF(strNote="", "null","'" & strNote & "'") & ")"
		else
			strPriKeyarr=Split(strPriKey,"#")
			strSql="UPDATE  " & strTimeSheetTable  &_
				" SET TDate = '" & ConvertTommddyyyy(strDate) & "'" & _
				" ,AssignmentID = " & strAssignment & _
				" ,Hours = " & strHours & _      
				" ,Note = " & IIF(strNote="", "null","'" & strNote & "'") & _
			" WHERE staffID=" & strUserid & " AND AssignmentID=" & strPriKeyarr(1) & " AND Tdate='" & strPriKeyarr(0) & "'"
			
			strPriKey=""
		end if
	else
		strPriKeyarr=Split(Request.querystring("subid"),"**")
		'response.write year(strPriKeyarr(1))
		if year(strPriKeyarr(1))<>Year(Date()) then 
			strSql="DELETE FROM ATC_Timesheet" & year(strPriKeyarr(1))& _
				" WHERE staffID=" & strUserid & " AND AssignmentID=" & strPriKeyarr(0) & " AND Tdate='" & strPriKeyarr(1) & "'"
		else
			strSql="DELETE FROM ATC_Timesheet" & _
				" WHERE staffID=" & strUserid & " AND AssignmentID=" & strPriKeyarr(0) & " AND Tdate='" & strPriKeyarr(1) & "'"
		end if
		strPriKey=""		
	end if

	strCnn = Application("g_strConnect")	
	Set objDatabase = New clsDatabase     

	strError=""

	If objDatabase.dbConnect(strCnn) Then              
		if not objDatabase.runActionQuery(strSql) then 
		   gMessage = objDatabase.strMessage
		end if			  
	end if
	
end if       

'SELECT * FROM ATC_Timesheet WHERE 

'--------------------------------------------------
' Initialize recordset
'--------------------------------------------------	

		strSql="SELECT AssignmentID, ProjectID + ' - ' + SubTaskName as Subtask_ FROM ATC_Assignments a INNER JOIN ATC_Tasks b ON a.SubTaskID=b.SubTaskID " & _
				"WHERE a.fgDelete=0 AND StaffID=" & strUserid & " AND ProjectID like '%" & strAPKSearch & "%'"	
		Call GetRecordset(strSql,rs)
	    strAssignment= PopulateDataToListWithoutSelectTag(rs,"AssignmentID", "Subtask_",-1)
	    
		strSql="SELECT Fullname FROM HR_TPStaff WHERE TPUserID=" & strUserid 
		Call GetRecordset(strSql,rsTPUser)
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
	if strChoseMenu = "" then strChoseMenu = "AG"
	
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
            <a class="blue" href="listofcontractstaff.asp" >Contact Staff List</a>
            <div class="title" style="padding:10px; text-align:center;">Package Hours for TP</div>
			<div class="blue" style="padding:0px; text-align:center;">(<%=rsTPUser("Fullname")%>)</div>
			
        </td>
    </tr>
</tbody>
</table>

<div class="container-fluid">
<!-- Tab functions -->
   
    <div class="row">
        <div class="col-sm-12">
            <div class="tab-content employee-details-form">
                <div class="row">
                    <div class="col-md-12 col-sm-6 col-xs-12">
                        <form class="form-horizontal row-border" id="contactForm" method="POST" action="contractstaffPackageHours.asp?act=u&id=<%=strUserid%>">
<%if gMessage<>"" then%>                        
                            <div id="messages" class="alert alert-danger">
                                <strong>Error:</strong> <%=gMessage%>.
                            </div>
<%end if%>
            
    						<div class="panel panel-default">
				                <div class="panel-body">
				                    <div class="col-sm-6">
				                        <div class="form-group has-error">
				                            <label class="col-md-12" >APK </label>
				                            <div class="col-md-8">
				                                <input type="text" id="txtAPK" name="txtAPK" class="form-control" value="<%=strAPK%>">
				                            </div>
											<div class="col-md-4">
                                                <button type="button" id="btnSearch"  class="btn btn-primary btn-full-width">Search</button>
                                            </div>
				                        </div>
				                        <div class="form-group">
				                            <label class="col-md-12">Task</label>
				                            <div class="col-md-12">
												<select id="lstTasks" name="lstTasks" class="form-control">
				                                        <option value=""></option>
				                                        <%=strAssignment%>  
				                                </select>
				                                <input type="hidden" id="txtTask" name="txtTask" class="form-control"  value="<%=intAssignmentID%>">
				                            </div>
				                        </div>
				                    </div>
				                    <div class="col-sm-6">
				                        <div class="form-group">
				                            <label class="col-md-12">Input Date</label>
				                            <div class="col-md-12">
				                                <div class="input-group date">
				                                    <input type="text"  id="txtDate" name="txtDate" class="form-control datepicker" placeholder="DD/MM/YYYY"  value="">
				                                    <span class="input-group-addon">
				                                        <span class="ic-calendar"></span>
				                                    </span>
				                                </div>
				                            </div>
				                        </div>
				                       
				                        <div class="form-group">
				                            <label class="col-md-12">Hours</label>
				                            <div class="col-md-12">
				                                <input type="text"  id="txtHours" name="txtHours" class="form-control"  value="">
				                            </div>
				                        </div>
				                    </div>
									<div class="col-sm-12"> 
										<div class="form-group">									
											<label class="col-md-12">Note</label>
											<div class="col-md-12">
												<textarea  id="txtNote" name="txtNote" class="form-control" ></textarea>
											</div>
										</div>
                                        
									</div>
				                    
				                </div>
								
				            </div>
				            
				            <div class="col-sm-12">
				                <div class="form-group text-right">
				                    <button type="submit" id="btnNext" class="btn btn-primary btnNext">Save & Close</button>
				                    <button type="button" id="btnCancel" class="btn btn-default">Cancel</button>				                    
				                </div>
				            </div>
				            
				            <input type="hidden" name="txtuserid" value="<%=strUserid%>"/>
							<input type="hidden" name="txtPriKey" id="txtPriKey" value="<%=strPriKey%>"/>
							
							<div class="col-md-12">
								<table class="table table-striped table-bordered table-hover margin-bottom10 table-responsive" id="tblList">
									<thead class="thead-inverse">
										<tr>
											<th style="width: 10%">APK</th>
											<th style="width: 30%">Task Name</th>
											<th style="width: 10%">Month</th>
											<th style="width: 10%">Hours</th>
											<th style="width: 35%">Note</th>
											<th style="width: 5%" class="col-action"></th>
										</tr>
									 </thead>
									 <tbody>
										<%=PopulateList()%>                       
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
                    <p>You are about to delete this item.</p>
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
//alert ("test");
	$('#txtDate')
        .on('changeDate', function(e) {
            // Revalidate the date field
            $('#contactForm').formValidation('revalidateField', 'txtDate');
        });
		
	 $('#contactForm').formValidation({
            framework: 'bootstrap',
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },        
            fields: {
                lstTasks:{
                    validators: {
                        notEmpty: {
                            message: 'The Task name is required.'
                        }
                    }
                },                
                txtDate:{
                    validators: {
						notEmpty: {
                            message: 'The Input Date is required.'
                        },
                         date: {
                            format: 'DD/MM/YYYY',
                            message: 'The Input Date is not a valid'
                        }
                    }  
                 },
				 txtHours:{
                    validators: {
						notEmpty: {
                            message: 'The hours is required.'
                        },
                        numeric: {
                            message: 'The hours is not a number'
                        }
                    }  
                 }
            }
        });//end of validation
		
	$("#btnSearch").click( function()
	{		
			search();		
    }); 
	
	$("#btnCancel").click( function()
	{		
		window.location = 'listofcontractstaff.asp'
    }); 
    
	$(".editrow").on('click', function(e) {
		var strPriKey;
		strPriKey=$(this).parent().children("td.idate").attr("iDate") + '#' + $(this).parent().attr("idValue")
		document.getElementById("lstTasks").value =$(this).parent().attr("idValue");
		$("#txtDate").val($(this).parent().children("td.idate").text());
		$("#txtHours").val($(this).parent().children("td.iHours").text());
		document.getElementById("txtNote").value =$(this).parent().children("td.iNote").text();
		
		$("#txtPriKey").val(strPriKey);
		$("#txtStatus").val("u");
		
   });
   
    $('.btn-remove-item').on('click', function(e) {
            e.preventDefault();

            var id = $(this).data('id');            
            $("#modal_message").html("You are about to remove this item.");
            $('#confirm-delete').modal('show');            
             $('#confirm-delete').find('.btn-ok').attr('href', 'contractstaffPackageHours.asp?id=<%=strUserid%>&act=d&subid='+id );
           // //alert (id);
            //$('#myModal').data('id', id).modal('show');
        });      
   
 });
  
 function search(){
    // Revalidate the date field
    window.location = 'contractstaffPackageHours.asp?id=<%=strUserid%>&s=' + $("#txtAPK").val() 
}

</script>

</body>
</html>

