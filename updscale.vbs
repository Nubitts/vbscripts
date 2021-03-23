Dim Origen, puerto, usuario, passw 

Const ForAppending = 8
Const LOG_FOLDER = "C:\logapi\"
Const LOG_FILE = "test"
Const LOG_FILE_EXTENSION = ".log"
Const LOG_FILE_SEPARATOR = "_"


Origen = "192.168.1.123"

puerto = "3306"
usuario = "cristobal "
passw = "bascristo"
autoriza = "47ba844e3accdc9c71016c740a2111f2"

ChOrigen = Valida_ip(origen) 


if ChOrigen = true Then

    gross Origen,puerto,usuario,passw,autoriza
   
end if


sub gross(origen, puerto, usuario, password, autoriza)

    urldestino = "http://192.168.1.226/restapi/v1/grossw"    

    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=" & Origen & ";PORT=" & puerto & ";Database=bascula;User=" & usuario & ";Password=" & passw & ";option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dim valores

    dbconn.Open connect

    Query = "select zafra,nofecha,zona,numtra,nom_flet,horent,pesob FROM `b_ticket` where `status` = 'BATEY' and ZAFRA = 2021;"

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



sub posting(dato,url,autoriza)

Set objHTTP = CreateObject("Microsoft.XMLHTTP")

objHTTP.open "POST", url, False
 
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "Authorization", autoriza

msgbox dato

objHTTP.send dato
 
Respuesta = objHTTP.responseText
 
Set objHTTP = Nothing
 
 WriteLog(Respuesta)

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



sub WriteLog(message)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Get log file with current date 
	logFile = LOG_FOLDER & LOG_FILE & LOG_FILE_SEPARATOR & GetCurrentDate() & LOG_FILE_EXTENSION

	'Open log file in appending mode
	Set objLogger = objFSO.OpenTextFile(logFile,ForAppending,True)
	'Prefix timestamp
	message = FormatDateTime(Now(),vbLongDate) & " " & FormatDateTime(Now(),vbLongTime) & " >>> " & message
	'Write log message
	objLogger.WriteLine(message)
	'Close file
	objLogger.Close()
End Sub

Function GetCurrentDate()
	timeStamp = Now()
 	d = PrefixZero(Day(timestamp))
    	m = PrefixZero(Month(timestamp))
    	y = Year(timestamp)
    	GetCurrentDate=  d & m &  y
End Function

'Prefix zero if day or month is in single digit
Function PrefixZero(num)
	If(Len(num)=1) Then
		PrefixZero="0"&num
	Else
		PrefixZero=num
	End If
End Function
