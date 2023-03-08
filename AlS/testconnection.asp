<html>
<body>

<%
strSql="PROVIDER=SQLOLEDB;DATA SOURCE=VNHCMSUN02;DATABASE=TMS_CM;USER ID=timesheet;PASSWORD=tms;"	
'strSql="Provider=SQLNCLI;Server=VNHCMSUN01;Database=TMS_CM;Uid=timesheet; Pwd=AIS@2011;"

set conn=Server.CreateObject("ADODB.Connection")
conn.Open(strSql)

set rs = Server.CreateObject("ADODB.recordset")
rs.Open "SELECT * FROM ATC_users", conn
%>

<table border="1" width="100%">
<%do until rs.EOF%>
    <tr>
    <%for each x in rs.Fields%>
       <td><%Response.Write(x.value)%></td>
    <%next
    rs.MoveNext%>
    </tr>
<%loop
rs.close
conn.close
%>
</table>

</body>
</html> 