path = "\\Auxlaboratorio\incalab\laboratorio.accdb"

Destino =  "http://localhost/restapi/v1/sugart"

Set fso = CreateObject("Scripting.FileSystemObject")

If (fso.FileExists(path)) Then

    fso.CopyFile path, "d:\backlab\", true

    ' conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data Source=C:\Backup\laboratorio.accdb; Jet OLEDB:Database Password=cafe@kid_inca13;"

    '     Set dbconn = CreateObject("ADODB.Connection")
    '     Set myCommand = CreateObject("ADODB.Command")
    '     set rs = CreateObject("ADODB.Recordset")

    '     dbconn.Open connect

    '     Query = "select cañamolida as cane, azucarproducida from datosbasicos where zafra = '20-21'"

    '     rs.Open Query, dbconn

    '     if not rs.eof Then

    '         rs.movefirst
    
    '         while not rs.eof 
                
    '             msgbox cane

    
    '             rs.movenext
    
    '         Wend
    
    '     end if
    
    '     rs.Close

    '     dbconn.Close

End If



Function Valida_ip(ip)
    dim objPing, objRetStatus, Ping
    set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
    ("select * from Win32_PingStatus where address = '" & ip & "'" )
    for each objRetStatus in objPing
        if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode <> 0 then
            Ping = False
        else
            Ping = True
        end if
    next

    Valida_ip = Ping
end Function

Public Function IIf(blnExpression, vTrueResult, vFalseResult)
    If blnExpression Then
      IIf = vTrueResult
    Else
      IIf = vFalseResult
    End If
End Function

sub posting(dato,url,autoriza)

Set objHTTP = CreateObject("Microsoft.XMLHTTP")

objHTTP.open "POST", url, False
 
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "Authorization", autoriza

objHTTP.send dato
 
' MsgBox objHTTP.responseText
 
Set objHTTP = Nothing
 

end sub