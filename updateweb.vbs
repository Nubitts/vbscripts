Dim Origen 
Origen = "192.168.1.226"
Dim Destino 
Destino = "mysql.us.cloudlogin.co"

ChOrigen = Valida_ip(origen) 
ChDestino = Valida_ip(Destino)

if ChOrigen = true and ChDestino = true Then

    ' divisiones Origen,Destino
    
    ' zonas Origen,Destino

    ' grupos Origen,Destino

    ' fleteros Origen,Destino

    canes Origen,Destino

    ' sugar Origen,Destino

else
    msgbox "no hay conexion reportar a Informatica"
end if

sub sugar(Origen,Destino)
    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=192.168.1.226;PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select date_rec,hour,sugar,cane_ground from sugar_tempo where date_rec = DATE_SUB(curdate(),INTERVAL 1 DAY)"

    Queryd = "insert into sugar_tempo (date_rec,hour,sugar,cane_ground) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if check_reg(Destino,"sugar_tempo", "date_rec = '" & year(rs.Fields(0)) & "/" & iif(len(trim(rs.Fields(0)))=1,"0" & Month(rs.Fields(0)),Month(rs.Fields(0)) )  & "/" & day(rs.Fields(0))  & "' and hour = '" & rs.Fields(1) & "'") = true Then

                valores = "('" & year(rs.Fields(0)) & "/" & iif(len(trim(rs.Fields(0)))=1,"0" & Month(rs.Fields(0)),Month(rs.Fields(0)) )  & "/" & day(rs.Fields(0))  & "','" & rs.Fields(1) & "'," & rs.Fields(2) & ", " & rs.Fields(3) & ")"

                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close   
end sub

sub canes(Origen,Destino)
    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Campos = "idhr,zona,organiza,clave,nombre,ciclo,orden,ticket,fletero,fecha,hora,neto,descto,liquido,alzadora,diaz,zafrad,nofecha,tabla,grupo,"  & _
    "pesob,pesot,peson,pesol,plantas,socas,resocas,ton_cruda,ton_quemada,ton_descuentos,ton_castigos,btkt_cruda,btkt_quemada,btkt_caña,ton_manual,"  & _
    "ton_alzadora,ton_cosechadora,libre,fecque,horque,TPOCAN,fecpen,horent"


    Query = "select " & Campos & " from canes_tempo where zafrad = (select zafra from zafraparams where actual = 1 ) and fecha = curdate() order by orden, ticket;"

    Queryd = "insert into canes_tempo (" & Campos & ") values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"canes_tempo", "", " "," zafrad = (select zafra from zafraparams where actual = 1 ) and  orden =  " & rs.Fields(6) & " and ticket = " & rs.Fields(7)) = true Then

                valores = " (" & rs.Fields(0) & ", " 
                valores = valores & rs.Fields(1) & ", "
                valores = valores & "'" & rs.Fields(2) & "', "
                valores = valores & rs.Fields(3) & ", "
                valores = valores & "'" & rs.Fields(4) & "', "
                valores = valores & "'" & rs.Fields(5) & "', "
                valores = valores & rs.Fields(6) & ", "
                valores = valores & rs.Fields(7) & ", "
                valores = valores & rs.Fields(8) & ", "
                valores = valores & "'" & iif(instr(rs.Fields(9),"/") > 0,conv_f(rs.Fields(9)),rs.Fields(9))  & "', "
                valores = valores & "'" & rs.Fields(10) & "', "
                valores = valores & rs.Fields(11) & ", "
                valores = valores & rs.Fields(12) & ", "
                valores = valores & rs.Fields(13) & ", "
                valores = valores & rs.Fields(14) & ", "
                valores = valores & rs.Fields(15) & ", "
                valores = valores & rs.Fields(16) & ", "
                valores = valores & rs.Fields(17) & ", "
                valores = valores & rs.Fields(18) & ", "
                valores = valores & rs.Fields(19) & ", "
                valores = valores & rs.Fields(20) & ", "
                valores = valores & rs.Fields(21) & ", "
                valores = valores & rs.Fields(22) & ", "
                valores = valores & rs.Fields(23) & ", "
                valores = valores & rs.Fields(24) & ", "
                valores = valores & rs.Fields(25) & ", "
                valores = valores & rs.Fields(26) & ", "
                valores = valores & rs.Fields(27) & ", "
                valores = valores & rs.Fields(28) & ", "
                valores = valores & rs.Fields(29) & ", "
                valores = valores & rs.Fields(30) & ", "
                valores = valores & rs.Fields(31) & ", "
                valores = valores & rs.Fields(32) & ", "
                valores = valores & rs.Fields(33) & ", "
                valores = valores & rs.Fields(34) & ", "
                valores = valores & rs.Fields(35) & ", "
                valores = valores & rs.Fields(36) & ", "
                valores = valores & "'" &  rs.Fields(37) & "', "
                valores = valores & "'" &  rs.Fields(38) & "', "
                valores = valores & "'" &  rs.Fields(39) & "', "
                valores = valores & "'" &  rs.Fields(40) & "', "
                valores = valores & "'" &  rs.Fields(41) & "', "
                valores = valores & "'" &  rs.Fields(42) & "')"
                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

sub fleteros(Origen,Destino)
    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select cveforw, fullname, idzone,idzaf,type,activate,dateadd from forwarders order by cveforw;"

    Queryd = "insert into forwarders (cveforw, fullname,idzone,idzaf,type,activate,dateadd) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"forwarders", "cveforw", "'" & rs.Fields(0) & "'","" ) = true Then

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

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select cvegroup,description,idzone from groups order by cvegroup;"

    Queryd = "insert into `groups` (cvegroup,description,idzone,activate,dateadd) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"`groups`", "cvegroup", "'" & rs.Fields(0) & "'","" ) = true Then

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

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select cvezone,description,iddiv from zones order by cvezone;"

    Queryd = "insert into zones (cvezone,description,activate,dateadd,iddiv) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"zones", "cvezone", "'" & rs.Fields(0) & "'","" ) = true Then

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

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select cvediv,description,activate from divisions order by cvediv;"

    Queryd = "insert into divisions (cvediv,description,activate,dateadd) values"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            if valida_reg(Destino,"divisions", "cvediv", "'" & rs.Fields(0) & "'" ,"") = true Then

                valores = " ('" & rs.Fields(0) & "','" & rs.Fields(1) & "'," & rs.Fields(2) & ",now())"

                Ejecuta Queryd & valores, Destino

            end if 

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

function valida_reg(Destino,Tabla, Campo, Valor, expresion) 

    dim Resulta

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server="& Destino &";PORT=3306;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dbconn.Open connect

    Query = "select * from " & trim(Tabla) & " where " &  iif(Len(Trim(expresion))=0,Campo & " = " & Valor & ";",expresion )   

    rs.Open Query, dbconn

    Resulta = rs.eof

    rs.Close

    dbconn.Close


    valida_reg = Resulta

end function

sub Ejecuta(Oracion, Destino)
    dim connect1
    dim dbconn1
    dim myCommand1

    Set dbconn1 = CreateObject("ADODB.Connection")
    Set myCommand1 = CreateObject("ADODB.Command")

    connect1 = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server="& Destino &";PORT=3306;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

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

function check_reg(Destino,Tabla,expresion) 

        dim Resulta
    
        connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server="& Destino &";PORT=3306;Database=ingenioel_applications;User=ingenioel_applications;Password=ADVG12345;option=3;"
    
        Set dbconn = CreateObject("ADODB.Connection")
        Set myCommand = CreateObject("ADODB.Command")
        set rs = CreateObject("ADODB.Recordset")
    
        dbconn.Open connect
    
        Query = "select * from " & trim(Tabla) & " where " & expresion  
   
        rs.Open Query, dbconn
    
        Resulta = rs.eof
    
        rs.Close
    
        dbconn.Close
    
    
        check_reg = Resulta
    
end function    