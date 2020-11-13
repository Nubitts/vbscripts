Dim Origen 
Origen = "192.168.1.226"
Dim Destino 
Destino = "192.168.1.92"


dim ChOrigen = Valida_ip(origen) 
dim ChDestino = Valida_ip(Destino)

if ChOrigen = true and ChDestino = true Then

    'Actualizacion de divisiones
    divisiones(Origen,Destino)
    

end if

sub divisiones(Origen, Destino)

    dim connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3306;Database=cred_campo;User=luis;Password=admin;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select no_div, division from divisiones order by no_div;"

    Queryd = "insert into divisions (cvediv,description,activate,dateadd) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"divisions", "cvediv", "'" & rs.Fields(0) & "'" ) = false Then

                valores = " ('" & rs.Fields(0) & "','" & rs.Fields(1) & "',1,now())"

                Ejecuta Queryd & valores, Destino

            end if 

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

function valida_reg(Destino,Tabla, Campo, Valor) 

    dim Resulta

    dim connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Destino & ";PORT=3306;Database=applications;User=root;Password=12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dbconn.Open connect

    Query = "select * from " & trim(Tabla) & " where " & Campo & " = " & Valor & ";"

    rs.Open Query, dbconn

    if rs.eof then Resulta = false else Resulta = true end if

    rs.Close

    dnconn.Close


    valida_reg = Resulta

end function

sub Ejecuta(Oracion, Destino)
    dim connect1
    dim dbconn1
    dim myCommand1

    Set dbconn1 = CreateObject("ADODB.Connection")
    Set myCommand1 = CreateObject("ADODB.Command")

    connect1 = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Destino & ";PORT=3306;Database=applications;User=root;Password=12345;option=3;"

    dbconn1.Open connect1

    Set myCommand1.ActiveConnection = dbconn1

    myCommand1.CommandText = Oracion

    myCommand1.Execute

    dbconn1.Close

end sub 

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
