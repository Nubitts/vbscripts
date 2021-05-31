Dim Origen 
Origen = "192.168.1.123"
Dim Destino 
Destino = "192.168.1.226"


ChOrigen = Valida_ip(origen) 
ChDestino = Valida_ip(Destino)

if ChOrigen = true and ChDestino = true Then

   canes Origen,Destino

end if


sub canes(Origen,Destino)
    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3306;Database=bascula;User=cristobal;Password=bascristo;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Campos = "zafra,nofecha,zona,numtra,nom_flet,horent,pesob"


    Query = "select " & Campos & " from  `b_ticket` where `status` = 'BATEY' and ZAFRA = 2021;"

    Queryd = "insert into truck_in_batey (zafra,nofecha,zona,numtra,nom_flet,horent,pesob) values"

    resulta = valida_reg("") 

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

                valores = " (" & rs.Fields(0) & ", " 
                valores = valores & rs.Fields(1) & ", "
                valores = valores & "'" & rs.Fields(2) & "', "
                valores = valores & "'" & rs.Fields(3) & "', "
                valores = valores & "'" & rs.Fields(4) & "', "
                valores = valores & "'" & rs.Fields(5) & "', "
                valores = valores & rs.Fields(6) & ")"
               
                Ejecuta Queryd & valores, Destino

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

sub fleteros(Origen,Destino)
    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3306;Database=cred_campo;User=luis;Password=admin;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select num_fle, nombre, zona,1 as zafra,seltipo from fleteros order by num_fle;"

    Queryd = "insert into forwarders (cveforw, fullname,idzone,idzaf,type,activate,dateadd) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"forwarders", "cveforw", "'" & rs.Fields(0) & "'","" ) = false Then

                valores = " ('" & rs.Fields(0) & "','" & rs.Fields(1) & "'," & rs.Fields(2) & ", " & rs.Fields(3) & ", '" & left(rs.Fields(4),1) & "', 1, now())"

                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close
end sub

sub grupos(Origen,Destino)

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3306;Database=cred_campo;User=luis;Password=admin;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select clave_grupo,nombre_grupo,zona from grupos order by clave_grupo;"

    Queryd = "insert into `groups` (cvegroup,description,idzone,activate,dateadd) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"`groups`", "cvegroup", "'" & rs.Fields(0) & "'","" ) = false Then

                valores = " ('" & iif(isnull(rs.Fields(0)) = true," ",rs.Fields(0)) & "','" & iif(isnull(rs.Fields(1)) = true," ",rs.Fields(1)) & "'," & iif(isnull(rs.Fields(2)) = true,"0",rs.Fields(2)) & ",1,now())"
               
                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close
end sub

sub zonas(Origen,Destino)

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3306;Database=cred_campo;User=luis;Password=admin;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select no_zona,zona,no_div from zonas order by no_zona;"

    Queryd = "insert into zones (cvezone,description,activate,dateadd,iddiv) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"zones", "cvezone", "'" & rs.Fields(0) & "'","" ) = false Then

                valores = " ('" & rs.Fields(0) & "','" & rs.Fields(1) & "',1,now(),"& rs.Fields(2) &")"

                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close
end sub

sub divisiones(Origen, Destino)

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3306;Database=cred_campo;User=luis;Password=admin;option=3;"

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

            if valida_reg(Destino,"divisions", "cvediv", "'" & rs.Fields(0) & "'" ,"") = false Then

                valores = " ('" & rs.Fields(0) & "','" & rs.Fields(1) & "',1,now())"

                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

function valida_reg(Destino) 

    dim Resulta
    dim connect1
    dim dbconn
    dim myCommand1

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=192.168.1.226;PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand1 = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    Oracion = "truncate table truck_in_batey"

    dbconn.Open connect

    Set myCommand1.ActiveConnection = dbconn

    myCommand1.CommandText = Oracion

    myCommand1.Execute

    dbconn.Close

    valida_reg = true

end function

sub Ejecuta(Oracion, Destino)
    dim connect1
    dim dbconn1
    dim myCommand1

    Set dbconn1 = CreateObject("ADODB.Connection")
    Set myCommand1 = CreateObject("ADODB.Command")

    connect1 = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server="& Destino & ";PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    dbconn1.Open connect1

    Set myCommand1.ActiveConnection = dbconn1

MsgBox Oracion

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

Public Function IIf(blnExpression, vTrueResult, vFalseResult)
    If blnExpression Then
      IIf = vTrueResult
    Else
      IIf = vFalseResult
    End If
End Function

function conv_f(fecha)
    dim Position, dia, mes, anual, regresa

    Position = instr(fecha,"/")

    if Position > 0 Then

        dia = left(fecha,Position-1)
    
        mes = mid(fecha,position+1,2)
    
        anual = Right(fecha,4)
    
        regresa = anual & "-" & mes & "-" & dia
    
    end if

    conv_f = regresa

end function