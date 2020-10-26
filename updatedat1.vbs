Option Explicit
    
    dim objPing, objRetStatus, Ping
    set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '192.168.1.92'")
    for each objRetStatus in objPing
        if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode <> 0 then
        Ping = False
            ' WScript.Echo "Status code is " & objRetStatus.StatusCode
        else
            Ping = True
            ' Wscript.Echo "Bytes = " & vbTab & objRetStatus.BufferSize
            ' Wscript.Echo "Time (ms) = " & vbTab & objRetStatus.ResponseTime
            ' Wscript.Echo "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
        end if
    next

    if ping = true Then

        Dim dbconn, connect, myCommand
        dim Campos  = "ZAFRA" & _
            "CODIGO, " & _
            "NOMBRE_P, " & _
            "GRUPO, " & _
            "NOM_GRUPO, " & _
            "TABLA, " & _
            "CICLO, " & _
            "TICKET, " & _
            "ORDCTE, " & _
            "TPOCAN, " & _
            "NUMALZ, " & _
            "DESCTO, " & _
            "CASTIGO, " & _
            "PESON, " & _
            "PESOB, " & _
            "PESOT, " & _
            "PESOL, " & _
            "OMITIDO, " & _
            "NUMTRA, " & _
            "HORSAL, " & _
            "HORENT, " & _
            "FECPES, " & _
            "TIPO_TICK, " & _
            "TODESCA, " & _
            "FECPEN, " & _
            "FECQUE, " & _
            "HORQUE, " & _
            "TIPQUE, " & _
            "AVISO, " & _
            "NUMAVI, " & _
            "MARCA, " & _
            "MATERIAL, " & _
            "FECHAKK, " & _
            "rfc_empresa, " & _
            "noFecha, " & _
            "timebatey, " & _
            "totaldescuento, " & _
            "totalcastigo, " & _
            "hora, " & _
            "status, " & _
            "diazafra, " & _
            "hr_code, " & _
            "nom_flet, " & _
            "nom_alz, " & _
            "transportista, " & _
            "fletero, " & _
            "zona, " & _
            "organiza, " & _
            "libre, " & _
            "ent_id_user, " & _
            "ent_usuario, " & _
            "sal_id_user, " & _
            "sal_usuario, " & _
            "observa" 

        connect = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;PORT=3306;Database=cred_campo;User=root;Password=12345;"

        Set dbconn = CreateObject("ADODB.Connection")
        Set myCommand = CreateObject("ADODB.Command" )

        dbconn.Open connect

        dbconn.Execute = "DELETE FROM b_ticket"

        Dim xmlDoc

        Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
        xmlDoc.Async = "False"
        xmlDoc.Load( "c:\temporal\result20201026_093408.xml" )


        dim x 
        dim objSubNode 

        for each x in xmlDoc.documentElement.childNodes
            
            dim y
            dim 

            For Each y In x.childNodes
                WScript.Echo y.nodeName               
            Next ' Element
       
           
        next


        dbconn.Close

        Wscript.Echo "positivo"
    end if