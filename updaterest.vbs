
On Error Resume Next

Dim Origen, puerto, usuario, passw 
Origen = "192.168.1.226"

puerto = "3307"
usuario = "masteroot"
passw = "ADVG12345"
autoriza = "47ba844e3accdc9c71016c740a2111f2"

ChOrigen = Valida_ip(origen) 


if ChOrigen = true Then

  '  tprob Origen,puerto,usuario,passw,autoriza

    canes Origen,puerto,usuario,passw,autoriza

  '  sugar Origen,puerto,usuario,passw,autoriza
   
end if

sub tprob(origen, puerto, usuario, password, autoriza)

    urldestino = "http://www.ingenioelcarmen.com/restdata/v1/probv"    

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=" & puerto & ";Database=applications;User=" & usuario & ";Password=" & passw & ";option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Campos = "fecha, division, zona, probtoday, wd, nofecha"


    Query = "select " & Campos & " from vprobdeldivzn order by nofecha;"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            lRecCnt = lRecCnt + 1
            sFlds = ""
            for each fld in rs.Fields

                sFld = fld.Name & "=" & iif(instr(fld.Value,"/") = 0, toUnicode(iif(isnull(fld.Value)=true,"_",fld.Value)),conv_f(fld.Value))
                sFlds = sFlds & iif(sFlds <> "", "&", "") & sFld

            next 
            sRec = sFlds 

            posting sRec,urldestino,autoriza

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

sub sugar(origen, puerto, usuario, password, autoriza)

    urldestino = "http://www.ingenioelcarmen.com/restdata/v1/sugar"    

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=" & puerto & ";Database=applications;User=" & usuario & ";Password=" & passw & ";option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Campos = "date_rec,hour,sugar,cane_ground, nofecha"


    Query = "select " & Campos & " from sugar_tempo order by date_rec, hour;"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            lRecCnt = lRecCnt + 1
            sFlds = ""
            for each fld in rs.Fields

                sFld = fld.Name & "=" & iif(instr(fld.Value,"/") = 0, toUnicode(iif(isnull(fld.Value)=true,"_",fld.Value)),conv_f(fld.Value))
                sFlds = sFlds & iif(sFlds <> "", "&", "") & sFld

            next 
            sRec = sFlds 

            posting sRec,urldestino,autoriza

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close

end sub

sub canes(origen, puerto, usuario, password, autoriza)

    urldestino = "http://www.ingenioelcarmen.com/restdata/v1/canes"

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=" & puerto & ";Database=applications;User=" & usuario & ";Password=" & passw & ";option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Campos = "idhr,zona,organiza,clave,case when LENGTH(nombre) = 0 then '_' else nombre end as nombre,ciclo,orden,ticket,fletero,fecha,hora,neto,descto,liquido,alzadora,diaz,zafrad,nofecha,derivada,tabla,grupo,"  & _
    "pesob,pesot,peson,pesol,plantas,socas,resocas,ton_cruda,ton_quemada,ton_descuentos,ton_castigos,btkt_cruda,btkt_quemada,btkt_caña as btkt_cana,ton_manual,"  & _
    "ton_alzadora,ton_cosechadora,libre,fecque,horque,TPOCAN,fecpen,horent,nom_grupo"


    Query = "select " & Campos & " from canes_tempo where zafrad = (select zafra from zafraparams where actual = 1 ) and nofecha = '2105%' order by fecha desc, hora desc;"

    rs.Open Query, dbconn

    if not rs.eof Then

        rs.movefirst

        while not rs.eof 

            lRecCnt = lRecCnt + 1
            sFlds = ""
            for each fld in rs.Fields

                sFld = fld.Name & "=" & iif(isnull(fld.Value)=false,iif(instr(fld.Value,"/") = 0, toUnicode(iif(isnull(fld.Value)=true,"_",fld.Value)),conv_f(fld.Value)),"_")
                sFlds = sFlds & iif(sFlds <> "", "&", "") & sFld

            next 
            sRec = sFlds 

            posting sRec,urldestino,autoriza

            rs.movenext

        Wend

    end if

    rs.Close

    dbconn.Close


end sub

sub reglog(evento,describe)

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=192.168.1.226;PORT=3307;Database=applications;User=masteroot;Password=ADVG12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    valores = "'" & evento & "', '" & describe & "','APIday','API',curdate()"

    Query = "insert into logs (event,description,tuser,tapp,daterec)  values (" & Campos & ");"


    Set myCommand1.ActiveConnection = dbconn

    myCommand1.CommandText = Query

    myCommand1.Execute

    dbconn.Close

end sub

sub posting(dato,url,autoriza)

    dim Respuesta

    Set objHTTP = CreateObject("Microsoft.XMLHTTP")

    objHTTP.open "POST", url, False
    
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.setRequestHeader "Authorization", autoriza

    objHTTP.send dato
    
    Respuesta = objHTTP.responseText
    
    Set objHTTP = Nothing

    reglog "insert",Respuesta
    

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

Function IIf(blnExpression, vTrueResult, vFalseResult)
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

function toUnicode(str)
    dim x
    dim uStr
    dim uChr
    dim uChrCode
    uStr = ""
    for x = 1 to len(str)
        uChr = mid(str,x,1)
        uChrCode = asc(uChr)
        if uChrCode = 8 then ' backspace
            uChr = "\b" 
        elseif uChrCode = 9 then ' tab
            uChr = "\t" 
        elseif uChrCode = 10 then ' line feed
            uChr = "\n" 
        elseif uChrCode = 12 then ' formfeed
            uChr = "\f" 
        elseif uChrCode = 13 then ' carriage return
            uChr = "\r" 
        elseif uChrCode = 34 then ' quote 
            uChr = "\""" 
        elseif uChrCode = 39 then ' apostrophe
            uChr = "\'" 
        elseif uChrCode = 92 then ' backslash
            uChr = "\\" 
        elseif uChrCode < 32 or uChrCode > 127 then ' non-ascii characters
            uChr = "\u" & right("0000" & CStr(uChrCode),4)
        end if
        uStr = uStr & uChr
    next
    toUnicode = uStr
end function

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