
<%    
    Response.ContentType = "application/json; charset=utf-8"
    
    callback = Request("callback") ' assuming callback contains the callback function name

    strSql = "SELECT ProjectID FROM ATC_Projects WHERE fgActivate=1"

    'strSql="SELECT * FROM HR_Employee"
    strconn=Application("g_strConnect")	

    set conTem=Server.CreateObject("ADODB.Connection")
    conTem.Open(strconn)

    Set rsElementTem = Server.CreateObject("ADODB.Recordset")
    rsElementTem.Open strSql,conTem,3,3
    strArr=""
    if not rsElementTem.EOF then
        rsElementTem.MoveFirst
        'rsElementTem.Filter="Fullname like '*" & request.QueryString("term") & "*'"
                
        do while not rsElementTem.EOF
            if strArr<>"" then strArr=strArr & ","
            strArr=strArr & "{"&_
                            """value"":""" & rsElementTem("ProjectID") & """, ""id"":" & rsElementTem("ProjectID") & "}" 
                                      
            rsElementTem.MoveNext
        loop
    end if

    Response.Write "[" & strArr & "]"
%> 