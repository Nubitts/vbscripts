
    connect = "Driver={MySQL ODBC 8.0 ANSI Driver};charset=UTF8;Server=localhost;PORT=3306;Database=applications;User=root;Password=12345;option=3;"

    Set dbconn = CreateObject("ADODB.Connection")
    Set myCommand = CreateObject("ADODB.Command")
    set rs = CreateObject("ADODB.Recordset")

    dbconn1.Open connect1

    Set myCommand1.ActiveConnection = dbconn1

    oracion = "drop table if exists temporal; " & _
              "create TEMPORARY table temporal select * from canes_tempo where zafrad = (select zafra from zafraparams where actual= 1)	and " & _
              "((DATE_FORMAT(fecha, '%Y-%m-%d') = CURDATE() and LEFT(hora,2) BETWEEN 6 and 23) or (datediff( DATE_FORMAT(fecha, '%Y-%m-%d') , CURDATE() ) = 1  and LEFT(hora,2) BETWEEN 1 and 5)) " & _
              "group by left(hora,2); " & _
              "insert into canes_week " & _
              "select min(fecha) as fecha,sum(plantas) as plantas, sum(socas) as socas, sum(resocas) as resocas, sum(ton_cruda) as cruda, sum(ton_quemada) as quemada, sum(neto) as entrada, sum(ton_descuentos) as mat_ex_tot, " & _
              "case when sum(ton_alzadora) = 0 and sum(ton_cosechadora) = 0 then sum(ton_descuentos) else 0 end as mat_ex_man, " & _
              "case when sum(ton_alzadora) > 0 and sum(ton_cosechadora) = 0 then sum(ton_descuentos) else 0 end as mat_ex_mec, " & _
              "case when sum(ton_alzadora) = 0 and sum(ton_cosechadora) > 0 then sum(ton_descuentos) else 0 end as mat_ex_pic, " & _
              "sum(ton_castigos) as castigos, sum(liquido) as liquida, " & _
              "case when sum(ton_alzadora) = 0 and sum(ton_cosechadora) = 0 then sum(neto) else 0 end as alce_manual, " & _
              "case when sum(ton_alzadora) > 0 and sum(ton_cosechadora) = 0 then sum(neto) else 0 end as alce_meca, "  & _
              "case when sum(ton_alzadora) = 0 and sum(ton_cosechadora) > 0 then sum(neto) else 0 end as picada, " & _
              "count(distinct ticket) as tickets, count(if(ton_cruda >0,1,null)) as t_cruda, count(if(tpocan = 'Q',1,null)) as t_quema from temporal;"

    