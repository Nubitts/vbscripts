
Dim objIP

Set objIP = CreateObject( "SScripting.IPNetwork" )

If objIP.Ping( "" ) = 0 Then
    PingSSR = True
Else
    PingSSR = False
End If

Set objIP = Nothing