
On Error Resume Next

    Dim Campos, regEx, Sin

    Sin  = "aeiouAEIOUaeiouAEIOUnN"

    Set regEx = New RegExp

    regEx.Pattern = "áéíóúÁÉÍÓÚàèìòùÀÈÌÒÙñÑ"
    regEx.Global = True

    Campos = "zafra,CODIGO,NOMBRE_P,GRUPO,NOM_GRUPO,TABLA,CICLO,TICKET,ORDCTE,TPOCAN,NUMALZ,DESCTO,CASTIGO,PESON,PESOB,PESOT,PESOL,OMITIDO,NUMTRA," & _
    "HORSAL,HORENT,FECPES,TIPO_TICK,TODESCA,FECPEN,FECQUE,HORQUE,TIPQUE,AVISO,NUMAVI,MARCA,MATERIAL,FECHAKK,rfc_empresa,noFecha,timebatey," & _ 
    "totaldescuento,totalcastigo,hora,status,diazafra,hr_code,nom_flet,nom_alz,transportista,fletero,zona,organiza,libre,ent_id_user,ent_usuario, " & _
    "sal_id_user,sal_usuario,observa"

    Dim Origen 
    Origen = "192.168.1.123"
    Dim Destino 
    Destino = "192.168.1.226"
    Dim sZafra 
    sZafra = "2021"
    Dim dbconn, connect, myCommand, rs
    Dim Ping, Query, Query1, nReg, Valores, Contiene, Ping1

    Ping = Valida_ip(Origen)
    Ping1 = Valida_ip(Destino)

    if Ping = true and Ping1 = true Then

        connect = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & Origen & ";PORT=3306;Database=bascula;User=cristobal;Password=bascristo;"

        Set dbconn = CreateObject("ADODB.Connection")
        Set myCommand = CreateObject("ADODB.Command")
        set rs = CreateObject("ADODB.Recordset")

        dim queries(3)

        queries(0) = "UPDATE b_ticket SET  NOMBRE_P = REPLACE(NOMBRE_P, 'Ñ', 'N') WHERE NOMBRE_P LIKE '%Ñ%'"
        queries(1) = "UPDATE b_ticket SET  nom_alz = REPLACE(NOMBRE_P, 'Ñ', 'N') WHERE nom_alz LIKE '%Ñ%'"
        queries(2) = "UPDATE b_ticket SET  nom_flet = REPLACE(nom_flet, 'Ñ', 'N') WHERE nom_flet LIKE '%Ñ%'"

        dbconn.Open connect

        Set myCommand1.ActiveConnection = connect

        for iC = 0 to 10
            myCommand1.CommandText =  queries(iC)
            myCommand1.Execute
        next

        Query = "select " & Campos & " from b_ticket where status = 'OK' and zafra = '" & sZafra & "' order by ticket;"

        Query1 = "insert into b_ticket_b (" & Campos & ") values "

        rs.Open Query, dbconn

        if not rs.eof Then
            rs.moveFirst

            while not rs.eof
                
                nReg = nReg +1
    
                if nReg = 1 Then
                    Ejecuta "delete from b_ticket_b",Destino 
                end if

                ajustenom = REPLACE(rs.Fields(2),"MONTAÑAS","MONTANAS",1,,1) 
    
                Valores = "(" & _
                IIF(isnull(rs.Fields(0)) = false,"'" & rs.Fields(0) & "'","null") & "," & _
                IIF(isnull(rs.Fields(1)) = false,rs.Fields(1),"null") & ","  & _
                IIF(isnull(rs.Fields(2)) = false,"'" & ajustenom  & "'","null")  & ","    & _
                IIF(isnull(rs.Fields(3)) = false,rs.Fields(3),"null") & ","  & _
                IIF(isnull(rs.Fields(4)) = false,"'" & rs.Fields(4) & "'","null")  & ","   & _
                IIF(isnull(rs.Fields(5)) = false,rs.Fields(5),"null") & ","  & _
                IIF(isnull(rs.Fields(6)) = false,"'" & rs.Fields(6) & "'" ,"null")  & ","    & _
                IIF(isnull(rs.Fields(7)) = false,rs.Fields(7),"null") & ","  & _
                IIF(isnull(rs.Fields(8)) = false,rs.Fields(8),"null") & ","  & _
                IIF(isnull(rs.Fields(9)) = false,"'" & rs.Fields(9) & "'","null")   & ","   & _
                IIF(isnull(rs.Fields(10)) = false,rs.Fields(10),"null") & ","  & _
                IIF(isnull(rs.Fields(11)) = false,rs.Fields(11),"null") & ","  & _
                IIF(isnull(rs.Fields(12)) = false,rs.Fields(12),"null") & ","  & _
                IIF(isnull(rs.Fields(13)) = false,rs.Fields(13),"null") & ","  & _
                IIF(isnull(rs.Fields(14)) = false,rs.Fields(14),"null") & ","  & _
                IIF(isnull(rs.Fields(15)) = false,rs.Fields(15),"null") & ","  & _
                IIF(isnull(rs.Fields(16)) = false,rs.Fields(16),"null") & ","  & _
                IIF(isnull(rs.Fields(17)) = false,"'" & rs.Fields(17) & "'","null")  & ","   & _
                IIF(isnull(rs.Fields(18)) = false,rs.Fields(18),"null") & ","  & _
                IIF(isnull(rs.Fields(19)) = false,"'" & rs.Fields(19) & "'","null")   & ","   & _
                IIF(isnull(rs.Fields(20)) = false,"'" & rs.Fields(20) & "'","null")   & "," & _  
                IIF(isnull(rs.Fields(21)) = false,"'" & GetFormattedDate(rs.Fields(21)) & "'","null")    & ","  & _
                IIF(isnull(rs.Fields(22)) = false,"'" & rs.Fields(22) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(23)) = false,rs.Fields(23),"null") & ","  & _
                IIF(isnull(rs.Fields(24)) = false,"'" & GetFormattedDate(rs.Fields(24)) & "'","null")  & ","   & _
                IIF(isnull(rs.Fields(25)) = false,"'" & GetFormattedDate(rs.Fields(25)) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(26)) = false,"'" & rs.Fields(26) & "'","null")    & ","  & _
                IIF(isnull(rs.Fields(27)) = false,"'" & rs.Fields(27) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(28)) = false,"'" & rs.Fields(28) & "'","null")    & ","  & _
                IIF(isnull(rs.Fields(29)) = false,rs.Fields(29),"null") & ","  & _
                IIF(isnull(rs.Fields(30)) = false,"'" & rs.Fields(30) & "'","null")  & ","  & _
                IIF(isnull(rs.Fields(31)) = false,rs.Fields(31),"null") & ","  & _
                IIF(isnull(rs.Fields(32)) = false,"'" & GetFormattedDate(rs.Fields(32)) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(33)) = false,"'" & rs.Fields(33) & "'","null")    & ","  & _
                IIF(isnull(rs.Fields(34)) = false,rs.Fields(34),"null") & ","  & _
                IIF(isnull(rs.Fields(35)) = false,"'" & rs.Fields(35) & "'","null")    & ","  & _
                IIF(isnull(rs.Fields(36)) = false,rs.Fields(36),"null") & ","  & _
                IIF(isnull(rs.Fields(37)) = false,rs.Fields(37),"null") & ","  & _
                IIF(isnull(rs.Fields(38)) = false,rs.Fields(38),"null") & ","  & _
                IIF(isnull(rs.Fields(39)) = false,"'" & rs.Fields(39) & "'","null")   & ","   & _
                IIF(isnull(rs.Fields(40)) = false,rs.Fields(40),"null") & ","  & _
                IIF(isnull(rs.Fields(41)) = false,rs.Fields(41),"null") & ","  & _
                IIF(isnull(rs.Fields(42)) = false,"'" & rs.Fields(42) & "'","null")   & ","   & _
                IIF(isnull(rs.Fields(43)) = false,"'" & rs.Fields(43) & "'","null")  & ","  & _
                IIF(isnull(rs.Fields(44)) = false,"'" & rs.Fields(44)& "'","null") & ","  & _
                IIF(isnull(rs.Fields(45)) = false,"'" & rs.Fields(45) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(46)) = false,rs.Fields(46),"null") & ","  & _
                IIF(isnull(rs.Fields(47)) = false,"'" & rs.Fields(47) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(48)) = false,"'" & rs.Fields(48) & "'","null") & ","  & _
                IIF(isnull(rs.Fields(49)) = false,"'" & rs.Fields(49) & "'","null")  & ","  & _
                IIF(isnull(rs.Fields(50)) = false,"'" & rs.Fields(50) & "'","null")   & ","  & _
                IIF(isnull(rs.Fields(51)) = false, rs.Fields(51),"null")   & ","  & _
                IIF(isnull(rs.Fields(52)) = false,"'" & rs.Fields(52) & "'","null")  & ","    & _
                IIF(isnull(rs.Fields(53)) = false,"'" & rs.Fields(53) & "'","null")  & ")"
               
                Ejecuta Query1 & Valores,Destino
    
                rs.MoveNext
            wend
    
        end if


        dbconn.Close

        Actualiza Destino,sZafra

    End if

    if err.Number <> 0 Then
        MsgBox "Solicitar asistencia a Informatica, error de ejecucion " & err.Description & " " & err.Number &  " " & ajustenom
    end if


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

    Function IIf(bClause, sTrue, sFalse)
             On Error Resume Next

            if err.Number <> 0 Then
                WScript.Echo err.Description
            end if

            If CBool(bClause) Then
                IIf = sTrue
            Else 
                IIf = sFalse
            End If
    End Function

    sub Ejecuta(Oracion, Destino)
        dim connect1
        dim dbconn1
        dim myCommand1

        Set dbconn1 = CreateObject("ADODB.Connection")
        Set myCommand1 = CreateObject("ADODB.Command")

         connect1 = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Destino & ";PORT=3306;Database=cred_campo;User=luis;Password=admin;option=3;"

        dbconn1.Open connect1

        Set myCommand1.ActiveConnection = dbconn1

        myCommand1.CommandText = Oracion

        myCommand1.Execute

        dbconn1.Close

    end sub 
        
    Function GetFormattedDate (Date)
            strDate = CDate(Date)
            strDay = DatePart("d", strDate)
            strMonth = DatePart("m", strDate)
            strYear = DatePart("yyyy", strDate)
            If strDay < 10 Then
              strDay = "0" & strDay
            End If
            If strMonth < 10 Then
              strMonth = "0" & strMonth
            End If
            GetFormattedDate = strYear & "-" & strMonth & "-" & strDay
    End Function

    Sub Actualiza(Destino, zafra)

        dim Query
        dim Campos
        
        Campos = "idhr,zona,organiza,clave,nombre,ciclo,orden,ticket,fletero,fecha,hora,neto,descto,liquido,alzadora,diaz,zafrad,nofecha,tabla,grupo,"  & _
                 "pesob,pesot,peson,pesol,plantas,socas,resocas,ton_cruda,ton_quemada,ton_descuentos,ton_castigos,btkt_cruda,btkt_quemada,btkt_caña,ton_manual,"  & _
                 "ton_alzadora,ton_cosechadora,libre,fecque,horque,TPOCAN,fecpen,horent"

        ' Query = "delete from `caña` where zafrad = " & zafra

        ' Ejecuta Query,Destino

        Query = "insert into `caña` (" & Campos & ") select " & Campos & " from vcane where zafrad = " & zafra & " and ticket not in (select ticket from `caña` where zafrad = 2021)"

        Ejecuta Query,Destino

    end sub

