Destino = "localhost"
puerto = "3306"

Oracion = "delete from canes_tempo where MONTH(CURRENT_DATE())= month(fecha) and zafrad = 2021;"

connect1 = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server="& Destino & ";PORT=" & puerto & ";Database=applications;User=root;Password=12345;option=3;"


Set dbconn1 = CreateObject("ADODB.Connection")
Set myCommand1 = CreateObject("ADODB.Command")

dbconn1.Open connect1

Set myCommand1.ActiveConnection = dbconn1

myCommand1.CommandText = Oracion

myCommand1.Execute

dbconn1.Close