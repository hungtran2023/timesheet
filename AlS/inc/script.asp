
<%    
    Response.ContentType = "application/json; charset=utf-8"
    
    callback = Request("callback") ' assuming callback contains the callback function name

    strSql = "SELECT a.PersonID, a.Fullname, ISNULL(b.Photo,b.Username) as Username, Joindate, a.Jobtitle, " & _
                "a.Birthday,'none' as PCName,a.EmailAddress,a.EmailAddress_Ex,a.ExtPhone " &_
            "FROM HR_Employee a LEFT JOIN ATC_Users b ON a.PersonID = b.UserID " &_
            "WHERE b.Username NOT LIKE 'GROUP%' AND (a.LeaveDate IS NULL OR a.LeaveDate>=getdate()-1) " & _
            "AND Fullname like '%" & request.QueryString("term") & "%'"
    strSql =  strSql & "ORDER BY JoinDate DESC, a.FirstName, a.Lastname"

    
    strconn=Application("g_strConnect")	

    set conTem=Server.CreateObject("ADODB.Connection")
    conTem.Open(strconn)

    Set rsElementTem = Server.CreateObject("ADODB.Recordset")
    rsElementTem.Open strSql,conTem,3,3
    strArr=""
    if not rsElementTem.EOF then
        rsElementTem.MoveFirst
                
        do while not rsElementTem.EOF
            if strArr<>"" then strArr=strArr & ","
            strArr=strArr & "{"&_
                            """value"":""" & rsElementTem("Fullname") & """, ""id"":" & rsElementTem("PersonID") & "}" 
                                      
            rsElementTem.MoveNext
        loop
    end if

    Response.Write "[" & strArr & "]"
%> 