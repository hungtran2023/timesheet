<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->  
<%  
    Set cdoConfig = CreateObject("CDO.Configuration")  
 
    With cdoConfig.Fields  
        .Item(cdoSendUsingMethod) = cdoSendUsingPort  
response.write "SMTP:" & .Item(cdoSMTPServer) & "<br>"
        .Item(cdoSMTPServer) = "vn-hcm-mail"  
        .Update  
response.write "SMTP:" & .Item(cdoSMTPServer) & "<br>"
    End With 


 
    Set cdoMessage = CreateObject("CDO.Message")  
 
    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = "uyenchi@atlas.com" 
        .To = "uyenchi@atlas.com" 
        .Subject = "Sample CDO Message" 
        .TextBody = "This is a test for CDO.message" 
        .Send 
    End With 
 
    Set cdoMessage = Nothing  
    Set cdoConfig = Nothing  
 
Response.write "<HTML><head><title>A message has been sent.</title></head><body>A message has been sent.</body></HTML>"
 
%>



