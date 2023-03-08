<%
	intMonth = Month(Date)
	intYear = Year(Date)
%>
<html>
<head>
<title>Timesheet</title>

<style type="text/css">
<!--

-->
</style>
<link rel="stylesheet" href="../../../timesheet.css" type="text/css">
<script language="javascript">
<!--
	
function printpage()
{
	window.close();
}
	
//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmtms" method="post">
<table width="252" border="0" cellspacing="0" cellpadding="0" bordercolor="#003399" height="164">
  <tr> 
    <td valign="middle"> 
      <table width="250" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr bgcolor="#C0CAE6" align="center"> 
          <td height="50" class="title">Select Condition</td>
        </tr>
        <tr bgcolor="#C0CAE6">
          <td>
             <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#C0CAE6">
               <tr>
                 <td valign="top" width="10%" class="blue">&nbsp;</td>
	             <td class="blue-normal" width="15%" align="right">Month&nbsp;</td>
	             <td class="blue-normal" width="30%"> 
		           <select name="lbmonth" size="1" class="blue-normal">
			         <option <%If CInt(intMonth)=1 Then%>selected<%End If%> value="1">January</option>
			         <option <%If CInt(intMonth)=2 Then%>selected<%End If%> value="2">February</option>
	  			     <option <%If CInt(intMonth)=3 Then%>selected<%End If%> value="3">March</option>
				     <option <%If CInt(intMonth)=4 Then%>selected<%End If%> value="4">April</option>
				     <option <%If CInt(intMonth)=5 Then%>selected<%End If%> value="5">May</option>
				     <option <%If CInt(intMonth)=6 Then%>selected<%End If%> value="6">June</option>
				     <option <%If CInt(intMonth)=7 Then%>selected<%End If%> value="7">July</option>
				     <option <%If CInt(intMonth)=8 Then%>selected<%End If%> value="8">August</option>
				     <option <%If CInt(intMonth)=9 Then%>selected<%End If%> value="9">September</option>
				     <option <%If CInt(intMonth)=10 Then%>selected<%End If%> value="10">October</option>
				     <option <%If CInt(intMonth)=11 Then%>selected<%End If%> value="11">November</option>
				     <option <%If CInt(intMonth)=12 Then%>selected<%End If%> value="12">December</option>
				   </select>
	             </td>
	             <td width="15%" class="blue-normal" align="right">Year&nbsp;</td>
	             <td class="blue-normal" width="30%"> 
			       <select name="lbyear" size="1" class="blue-normal">
<%For ii=Year(Date)-2 To Year(Date)%>
 	        	     <option <%If ii=CInt(intYear) Then%>selected<%End If%> value="<%=ii%>"><%=ii%></option>
<%Next%>
	 		       </select>
	             </td>
	           </tr>
            </table>
          </td>
        </tr>  
        <tr bgcolor="#C0CAE6">
          <td>&nbsp;</td>
        </tr>
        <tr bgcolor="#C0CAE6">
          <td>
            <table width="60" border="0" cellspacing="5" cellpadding="0" align="center" height="20" name="aa">
              <tr> 
                <td bgcolor="#8CA0D1" onMouseOver="this.style.backgroundColor='#7791D1';" onMouseOut="this.style.backgroundColor='#8CA0D1';" class="blue" height="20" align="center"> 
                  <a href="javascript:printpage();" class="b">Submit</a>
                </td> 
			  </tr>
		    </table>	
		  </td>  
        </tr>
        <tr bgcolor="#C0CAE6">
          <td class="blue-normal">&nbsp;&nbsp;* Please select month and year to generate salary sheet.</td>
        </tr>  
        <tr bgcolor="#C0CAE6">
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>
