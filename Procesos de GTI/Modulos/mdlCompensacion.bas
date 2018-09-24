Attribute VB_Name = "mdlCompensacion"
Option Explicit

'EAM- si no encuentra un parte retorna -1, sino retorna el codigo del parte (pcompnro)
Public Function BuscarParte(ByVal Fecha As Date)
 Dim objRs As New ADODB.Recordset
 
    'EAM- Primero verifica el alcance Individual
    StrSql = "SELECT gti_parte_compensacionhs.pcompnro FROM gti_parte_compensacionhs " & _
            " INNER JOIN gti_parte_alcpartecomphs ON gti_parte_compensacionhs.pcompnro = gti_parte_alcpartecomphs.pcompnro " & _
            " WHERE gti_parte_compensacionhs.pcompalcnivel=3 AND gti_parte_compensacionhs.pcompfecdesde <= " & ConvFecha(Fecha) & " AND gti_parte_compensacionhs.pcompfechasta >= " & ConvFecha(Fecha)
    OpenRecordset StrSql, objRs
    
    'EAM- Verifica el alcance por Estructura
    If objRs.EOF Then
        StrSql = "SELECT gti_parte_compensacionhs.pcompnro FROM gti_parte_compensacionhs " & _
                " INNER JOIN gti_parte_alcpartecomphs ON gti_parte_compensacionhs.pcompnro = gti_parte_alcpartecomphs.pcompnro " & _
                " INNER JOIN his_estructura ON gti_parte_alcpartecomphs.pcompalcorigen = his_estructura.estrnro " & _
                " INNER JOIN alcance_testr ON his_estructura.tenro = alcance_testr.tenro " & _
                " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro " & _
                " WHERE gti_parte_compensacionhs.pcompalcnivel=2 AND gti_parte_compensacionhs.pcompfecdesde <= " & ConvFecha(Fecha) & _
                " AND gti_parte_compensacionhs.pcompfechasta >= " & ConvFecha(Fecha) & " AND alcance_testr.tanro= 31"
        OpenRecordset StrSql, objRs
        
        'EAM- Verifica el alcance Global
        If objRs.EOF Then
            StrSql = "SELECT gti_parte_compensacionhs.pcompnro FROM gti_parte_compensacionhs " & _
                    " INNER JOIN gti_parte_alcpartecomphs ON gti_parte_compensacionhs.pcompnro = gti_parte_alcpartecomphs.pcompnro " & _
                    " WHERE gti_parte_compensacionhs.pcompalcnivel=1 AND gti_parte_compensacionhs.pcompfecdesde <= " & ConvFecha(Fecha) & " AND gti_parte_compensacionhs.pcompfechasta >= " & ConvFecha(Fecha)
            OpenRecordset StrSql, objRs
        End If
    End If
    
    If objRs.EOF Then
        BuscarParte = -1
    Else
        BuscarParte = objRs!pcompnro
    End If
    
End Function

Public Sub Compensar_HorasPorParte(Fecha As Date, NroTer As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Realiza la compensacion de horas del dia a travez del parte diario.
' Autor      :
' Fecha      :
' Ultima Mod.: EAM - 14/06/2012
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim th_generar As Integer

Dim rsComp As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsAdComp As New ADODB.Recordset

Dim SumaAux As Single

Dim Aux_HSAComp As Single
Dim TotHorHHMM As String

Dim pcompnro As Long
Dim thCompensable As Single
Dim cantHsCompensan As Single
Dim TotHSAComp As Single
Dim totHsGenerar As Single
Dim TipoRedondeo As Integer

    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Compensación de horas..."
    End If

    
    pcompnro = BuscarParte(Fecha)
    If (pcompnro = -1) Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro ningún parte diario de compensacion de hora en la fecha: " & Fecha
        End If
        Exit Sub
    End If
    
    Call Politica(810)
    
    
    'EAM- Busco la configuración del parte de Compensacion de hora
    StrSql = "SELECT pcompthcompensable,pcompthgenera,conhredondeo FROM gti_parte_compensacionhs " & _
            " WHERE pcompnro= " & pcompnro
    OpenRecordset StrSql, rs
           
    'EAM- Si tiene parte de compensación de horas setea la configuracion
    If Not rs.EOF Then
        th_generar = rs!pcompthgenera
        thCompensable = rs!pcompthcompensable
        TipoRedondeo = rs!conhredondeo
    End If
    
    'EAM- Obtiene la cantidad de hs del AD de las horas compensable
    StrSql = "SELECT adcanthoras FROM gti_acumdiario WHERE ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rs!pcompthcompensable
    OpenRecordset StrSql, rsAdComp
    
    'EAM- Obtiene el total de hs acumuladas
    If Not rsAdComp.EOF Then
        Aux_HSAComp = rsAdComp!adcanthoras
    End If
    

    'EAM- Busca los tipos de hora configuradas que compensan
    StrSql = "SELECT pcompdthcompensa,pcompdcantporcentaje FROM gti_parte_compensacionhs_det WHERE pcompnro= " & pcompnro
    OpenRecordset StrSql, rsComp
    
    Do While Not rsComp.EOF
        'EAM- Obtiene la cantidad de hs del AD de la hora que compensa.
        StrSql = "SELECT adcanthoras FROM gti_acumdiario WHERE ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!pcompdthcompensa
        OpenRecordset StrSql, rsAdComp
    
        If rsAdComp.EOF Then
            GoTo Continuar_rsComp
        Else
            cantHsCompensan = rsAdComp!adcanthoras
                    
            'EAM- Aplico el % de compensación a las horas que se compensan
            TotHSAComp = TotHSAComp + (Aux_HSAComp * (rsComp!pcompdcantporcentaje / 100))
                    
            If TotHSAComp > cantHsCompensan Then

                'EAM- Cacula la cantidad de hs que llega a compensar
                totHsGenerar = totHsGenerar + (cantHsCompensan / (rsComp!pcompdcantporcentaje / 100) * Aux_HSAComp)
                Aux_HSAComp = (Aux_HSAComp - totHsGenerar)
                
                'EAM- Borro la hora que compensa ya que se utilizo todas para poder compensar
                StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!pcompdthcompensa
                objConn.Execute StrSql, , adExecuteNoRecords
                                       
                TotHSAComp = 0
            Else
                If TotHSAComp < cantHsCompensan Then
                    'EAM- Calculo la cantidad de horas a generar.
                    totHsGenerar = totHsGenerar + Aux_HSAComp
                    'EAM- Resto las horas ya que las horas que compensan son mas que la que hay que compensar
                    cantHsCompensan = cantHsCompensan - TotHSAComp
                            
                    TotHorHHMM = CHoras(cantHsCompensan, 60)
                    StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & cantHsCompensan & " WHERE " & _
                                " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!pcompdthcompensa
                    objConn.Execute StrSql, , adExecuteNoRecords
                            
                    StrSql = "DELETE FROM gti_acumdiario WHERE ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & thCompensable
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Aux_HSAComp = 0
                    TotHSAComp = 0
                Else
                    totHsGenerar = totHsGenerar + Aux_HSAComp
                    StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                            " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!pcompdthcompensa
                    objConn.Execute StrSql, , adExecuteNoRecords
                            
                    StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                    " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & thCompensable
                    objConn.Execute StrSql, , adExecuteNoRecords
                            
                    Aux_HSAComp = 0
                    TotHSAComp = 0
                End If
                
                GoTo Continuar_rsComp
            End If
        End If
        
Continuar_rsComp:
        rsComp.MoveNext
    Loop


If totHsGenerar > 0 Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "hora compensada: " & th_generar & " cantidad: " & totHsGenerar
    End If

    'EAM- Revisa que el tipo de Hora a generar no tenga horas generada. Si tiene, las sumas sino las inserta
    StrSql = " Select * from gti_acumdiario where adfecha = " & ConvFecha(Fecha) & " and ternro = " & NroTer & " and thnro = " & th_generar & ""
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        'EAM- No existe el tipo de hs en el AD
        Call objFechasHoras.Convertir_A_Hora(totHsGenerar * 60, TotHorHHMM)
        Call objFechasHoras.Redondeo_Horas_Tipo(TotHorHHMM, TipoRedondeo, totHsGenerar)
        TotHorHHMM = CHoras(totHsGenerar, 60)
        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                " VALUES (" & ConvFecha(Fecha) & "," & NroTer & "," & th_generar & "," & TotHorHHMM & "," & totHsGenerar & "," & _
                CInt(False) & "," & CInt(True) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        'EAM- Sumo al tipo de hs las horas generadas
        SumaAux = rs!adcanthoras + totHsGenerar
        totHsGenerar = CHoras(SumaAux, 60)
        Call objFechasHoras.Convertir_A_Hora(totHsGenerar * 60, TotHorHHMM)
        Call objFechasHoras.Redondeo_Horas_Tipo(TotHorHHMM, TipoRedondeo, totHsGenerar)
        TotHorHHMM = CHoras(totHsGenerar, 60)
        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & SumaAux & " WHERE " & _
                " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & th_generar
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

End Sub



Public Sub Compensar_Horas(Fecha As Date, NroTer As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Realiza la compensacion de horas del dia.
' Autor      :
' Fecha      :
' Ultima Mod.: FGZ - 03/01/2007
' Descripcion: no estaba actualizando bien las horas compensadas parcialmente
' ---------------------------------------------------------------------------------------------
Dim th_generar As Integer
Dim Cant As Single

Dim rsComp As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsAdComp As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset

Dim canthoras As Single
Dim acumula As Boolean
Dim SumaAux As Single

Dim objFeriado As New Feriado
Dim objBDia As New BuscarDia

Dim Aux_HSAComp As Single
Dim TotHorHHMM As String

    acumula = True
    Usa_Conv = False

    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Compensación de horas..."
    End If


    Call Politica(810)

    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = CnTraza
    objFeriado.Feriado Fecha, Empleado.Ternro, G_traza

    Set objBDia.Conexion = objConn
    Set objBDia.ConexionTraza = CnTraza
    objBDia.Buscar_Dia Fecha, Fecha_Inicio, Nro_Turno, Empleado.Ternro, P_Asignacion, G_traza
    Call initVariablesDia(objBDia)

    Call buscar_horas_turno
    ' el 20 es porque es el codigo de la horas compensadas
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 20 " & _
    " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
    OpenRecordset StrSql, rs

    If Not rs.EOF Then
        'se setea al tipo de horas que se compensan las horas compensables
        th_generar = rs!thnro
    Else
        'Entrada en la traza
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No esta configurado el Tipo de Hora Compensada para el Turno: " & Str(Nro_Turno)
        End If
        Exit Sub
    End If

    StrSql = "SELECT * FROM gti_compnsbl WHERE " & _
    " turnro = " & Nro_Turno & " ORDER BY compsblorden ASC"
    OpenRecordset StrSql, rsComp

    Do While Not rsComp.EOF
        StrSql = "SELECT * FROM gti_acumdiario WHERE " & _
        " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!thnro
        OpenRecordset StrSql, rsAdComp

        If rsAdComp.EOF Then
            GoTo Continuar_rsComp

        Else
            'FGZ - 03/01/2007 Este es buen lugar
            Aux_HSAComp = rsAdComp!adcanthoras

            StrSql = "SELECT * FROM gti_acompsar WHERE " & _
            " compsblnro = " & rsComp!compsblnro & " ORDER BY acomporden ASC"
            OpenRecordset StrSql, rs
            
            Do While Not rs.EOF

                StrSql = "SELECT * FROM gti_acumdiario WHERE " & _
                " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rs!thnro
                OpenRecordset StrSql, rsAD

                If rsAD.EOF Then
                    GoTo continuar_rs
                Else
                    'FGZ - 03/01/2006
                    'If rsAdComp!adcanthoras > rsAD!adcanthoras Then
                    If Aux_HSAComp > rsAD!adcanthoras Then
                        Cant = Cant + (rsAD!adcanthoras * (rs!acompptje / 100))
                        'canthoras = (rsAdComp!adcanthoras - rsAD!adcanthoras) * (rs!acompptje / 100)
                        canthoras = (Aux_HSAComp - rsAD!adcanthoras) * (rs!acompptje / 100)

                        TotHorHHMM = CHoras(canthoras, 60)
                        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & canthoras & " WHERE " & _
                        " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!thnro
                        objConn.Execute StrSql, , adExecuteNoRecords

                        StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                        " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rs!thnro
                        objConn.Execute StrSql, , adExecuteNoRecords

                        Aux_HSAComp = Aux_HSAComp - rsAD!adcanthoras
                    Else
                        'If rsAdComp!adcanthoras < rsAD!adcanthoras Then
                        If Aux_HSAComp < rsAD!adcanthoras Then
                            'cant = cant + (rsAdComp!adcanthoras * (rs!acompptje / 100))
                            'canthoras = (rsAD!adcanthoras - rsAdComp!adcanthoras) * (rs!acompptje / 100)
                            Cant = Cant + (Aux_HSAComp * (rs!acompptje / 100))
                            canthoras = (rsAD!adcanthoras - Aux_HSAComp) * (rs!acompptje / 100)

                            TotHorHHMM = CHoras(canthoras, 60)
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & canthoras & " WHERE " & _
                            " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rs!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords

                            StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                            " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords

                            Aux_HSAComp = 0
                        Else
                            Cant = Cant + (rsAD!adcanthoras * (rs!acompptje / 100))
                            StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                            " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rsComp!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords

                            StrSql = "DELETE FROM gti_acumdiario WHERE " & _
                            " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & rs!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords

                            Aux_HSAComp = 0
                        End If
                        GoTo Continuar_rsComp
                    End If
                End If

continuar_rs:
                rs.MoveNext
            Loop
        End If

Continuar_rsComp:
        rsComp.MoveNext
    Loop

If Cant > 0 Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "hora compensada: " & th_generar & " cantidad: " & Cant
    End If

    ' aca tenemos que revisar primero si ese tipo de Hora ya està insertado
    ' si es asi ==> tengo que modificar el registro sumandole las horas
    ' sino lo inserto

    StrSql = " Select * from gti_acumdiario where adfecha = " & ConvFecha(Fecha) & _
        " and ternro = " & NroTer & " and thnro = " & th_generar & ""
    OpenRecordset StrSql, rsAux

    If rsAux.EOF Then
        ' Ese tipo de hora no lo tiene en el cumulado diario
        ' entonces lo inserto
        TotHorHHMM = CHoras(Cant, 60)

        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
            " VALUES (" & ConvFecha(Fecha) & "," & NroTer & "," & th_generar & "," & TotHorHHMM & "," & Cant & "," & _
            CInt(False) & "," & CInt(True) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        ' ese tipo de hora ya la tiene
        ' Entonces le actualizo el total de horas
        SumaAux = rsAux!adcanthoras + Cant
        TotHorHHMM = CHoras(SumaAux, 60)
        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & SumaAux & " WHERE " & _
            " ternro = " & NroTer & " AND adfecha = " & ConvFecha(Fecha) & " AND thnro = " & th_generar
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

End Sub

Private Sub initVariablesTurno(ByRef T As BuscarTurno)

   p_turcomp = T.Compensa_Turno
   Nro_Grupo = T.Empleado_Grupo
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo

End Sub

Private Sub initVariablesDia(ByRef D As BuscarDia)
   
   Dia_Libre = D.Dia_Libre
   Nro_Dia = D.Numero_Dia
   Nro_Subturno = D.SubTurno_Numero
   Orden_Dia = D.Orden_Dia
   Trabaja = D.Trabaja

End Sub

Private Sub buscar_horas_turno()

Dim objRs As New ADODB.Recordset

    StrSql = " SELECT diacanthoras FROM gti_dias WHERE (dianro = " & Nro_Dia & ")"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        'No Tiene ningún registro de HC para el día
        Exit Sub
    Else
        tdias_oblig = objRs!diacanthoras
    End If

End Sub

