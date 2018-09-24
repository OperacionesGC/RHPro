Attribute VB_Name = "mdlGTI_PRC01_01"
Option Explicit

Global G_traza As Boolean
Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Dim objBTurno As New BuscarTurno


Global diatipo As Byte
Global ok As Boolean
Global esFeriado As Boolean
Global hora_desde As String
Global fecha_desde As Date
Global fecha_hasta As Date
Global Hora_desde_aux As String
Global hora_hasta As String
Global Hora_Hasta_aux As String
Global No_Trabaja_just As Boolean
Global nro_jus_ent As Long
Global nro_jus_sal As Long
Global Total_horas As Single
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Horas_Oblig As Single
Global Existe_Reg As Boolean
Global Forma_embudo  As Boolean

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global Nro_Justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global Nro_Grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global Aux_Tipohora As Integer
Global aux_TipoDia As Integer
Global Sigo_Generando As Boolean
'FGZ - 07/11/2008
Global Genero_Sin_Control_Presencia  As Boolean

Global Hora_Tol As String
Global Fecha_Tol As Date
Global hora_toldto As String
Global fecha_toldto As Date

Global Usa_Conv  As Boolean

Global tol As String

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single
Global objAD_Dup As New ADODB.Recordset
Global Justif_Completa As Boolean



Public Sub PRC_01(G_traza As Boolean, P_NroTer As Long, Fecha As Date, ByRef Cant As Long, ByRef arrTempHsAD() As THsAD)
'+-------------------------------------------------------------------+
'| REGISTRACIONES -> ACUMULADO DOIARIO                               |
'|                                                                   |
'| Procedimiento: PRC_01                                             |
'| Descripci¢n: Proceso de Calculo del Acumulado Diario              |
'| Parámetros: Debug       = Si genera traza o no                    |
'|             P_Nroter    = Codigo del Emplado                      |
'|             P_fecha     = Fecha a procesar                        |
'|             Cant        = Cantidad de Registros calculados        |
'| Autor: Loustau, Ezequiel Pablo                                    |
'| Creado: 04/11/02                                                  |
'+-------------------------------------------------------------------+
Dim linea As String
Dim Acum As Single
Dim Aux As Single
Dim horas_dest As Single
Dim tipo_hora_dest As Integer
Dim acumula As Boolean
Dim sumahoras As Long
Dim objRsH_C As New ADODB.Recordset
Dim AC_Cantidad As Single
Dim AC_HS As String
Dim THAnterior As Long
Dim tiempoConversion As Long


If depurar Then
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Acumulando tipos de horas.."
End If
Nivel_Tab_Log = Nivel_Tab_Log + 1
    
Inicio_Prc01:
'OrigEAM
    ' creo la tabla temporal para conversión de horas
    If UsaConversionHoras Then
        CreateTempTable (TTempWFAd)
    End If
    
    p_fecha = Fecha
    
    P_Asignacion = False
    Acum = 0
    
    ' FGZ - 22/10/2003
    ' Mauricio esto no corre mas, se tiene que procesar igual pero sin tocar las horas manuales
    
'    'Si el acumulado diario fue cargado a mano entonces no proceso
'    StrSql = "SELECT * FROM gti_acumdiario " & _
'             " WHERE (ternro = " & P_NroTer & ") AND " & _
'             " (adfecha = " & ConvFecha(p_fecha) & ") AND " & _
'             " (admanual = " & CInt(True) & ")"
'    OpenRecordset StrSql, objRs
'    If Not objRs.EOF Then
'        Flog.writeline "El acumulado diario fue cargado manualmente ==> Salgo sin Procesar"
'        Exit Sub
'    End If
    
    'Reproceso. Borro las entradas en al acumulado para la fecha y empleado
    StrSql = " DELETE FROM gti_acumdiario " & _
             " WHERE (ternro = " & P_NroTer & ") AND (adfecha = " & ConvFecha(p_fecha) & ") AND " & _
             " (admanual = " & CInt(False) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    OpenRecordset "SELECT empleado.ternro, empleado.empleg FROM Empleado WHERE Ternro = " & P_NroTer, objRs
    If objRs.EOF Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El Empleado no existe"
        Exit Sub
    Else
        Empleado.Ternro = P_NroTer
        Empleado.Legajo = objRs!empleg
    End If

    ' Busco el turno del empleado
    Set objBTurno.Conexion = objConn
    Set objBTurno.ConexionTraza = CnTraza
    objBTurno.Buscar_Turno p_fecha, Empleado.Ternro, G_traza
    Call initVariablesTurno(objBTurno)
  
    
    Call Politica(800) ' Obtencion de la fecha de procesamiento
    
    'FGZ - 13/11/2009 ----------------------------------
'    StrSql = " SELECT thnro, SUM(horcant) as SumaHoras, hornro FROM gti_horcumplido " & _
'             " WHERE (ternro = " & P_NroTer & ") AND " & _
'             " (( hordesde = " & ConvFecha(p_fecha) & " AND " & fec_proc & " = 1) OR " & _
'             "  ( horfecrep = " & ConvFecha(p_fecha) & " AND " & fec_proc & " = 2) OR " & _
'             "  ( horhasta = " & ConvFecha(p_fecha) & " AND " & fec_proc & " = 3))" & _
'             " GROUP BY thnro, hornro"
    StrSql = " SELECT thnro, horcant as SumaHoras, horas, hornro FROM gti_horcumplido "
    StrSql = StrSql & " WHERE (ternro = " & P_NroTer & ") AND "
    StrSql = StrSql & " (( hordesde = " & ConvFecha(p_fecha) & " AND " & fec_proc & " = 1) OR "
    StrSql = StrSql & "  ( horfecrep = " & ConvFecha(p_fecha) & " AND " & fec_proc & " = 2) OR "
    StrSql = StrSql & "  ( horhasta = " & ConvFecha(p_fecha) & " AND " & fec_proc & " = 3))"
    StrSql = StrSql & " ORDER BY thnro, hornro"
    OpenRecordset StrSql, objRsH_C
    If objRsH_C.EOF Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Tiene ningún registro de HC para el día"
        End If
        Exit Sub
    Else
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inserto los tipos de horas acumulados encontrados en el HC"
        End If
    End If

    Nivel_Tab_Log = Nivel_Tab_Log + 1
    
'    If Not objRsH_C.EOF Then
'        THAnterior = objRsH_C!thnro
'        AC_Cantidad = 0
'        AC_HS = "00:00"
'    End If
    THAnterior = -1
    Do While Not objRsH_C.EOF
        If depurar Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo de Hora: " & objRsH_C!thnro & " cantidad " & objRsH_C!sumahoras
        End If
        
        If THAnterior <> objRsH_C!thnro Then
            AC_Cantidad = 0
            AC_HS = "00:00"
            THAnterior = objRsH_C!thnro
        End If
        
        AC_Cantidad = AC_Cantidad + objRsH_C!sumahoras
        Call SHoras(AC_HS, IIf(IsNull(objRsH_C!Horas), "00:00", objRsH_C!Horas), AC_HS)
        
        
        'Inserto
        StrSql = "SELECT adcanthoras FROM gti_acumdiario "
        StrSql = StrSql & " WHERE admanual = " & CInt(False) & " AND (ternro = " & P_NroTer & ") AND "
        StrSql = StrSql & " (adfecha = " & ConvFecha(p_fecha) & ") AND "
        StrSql = StrSql & " (thnro = " & objRsH_C!thnro & ")"
        OpenRecordset StrSql, objAD_Dup
        If objAD_Dup.EOF Then
            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) "
            StrSql = StrSql & " VALUES (" & ConvFecha(p_fecha) & "," & P_NroTer & "," & objRsH_C!thnro & ",'" & AC_HS & "'," & AC_Cantidad & ","
            StrSql = StrSql & CInt(False) & "," & CInt(True) & ")"
        Else
            StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & AC_Cantidad
            StrSql = StrSql & " , horas = '" & AC_HS & "'"
            StrSql = StrSql & " WHERE admanual = " & CInt(False) & " AND (ternro = " & P_NroTer & ") AND "
            StrSql = StrSql & " (adfecha = " & ConvFecha(p_fecha) & ") AND "
            StrSql = StrSql & " (thnro = " & objRsH_C!thnro & ")"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords

        'EAM- Verifica si tiene configurado la conversion de horas.
        If UsaConversionHoras Then
            If depurar Then
                Flog.writeline
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversiones antes de Autorizar..."
            End If
            
            Nivel_Tab_Log = Nivel_Tab_Log + 1
            
            
            tiempoConversion = GetTickCount
            Call AD_Antes_Autorizar(depurar, P_NroTer, Nro_Turno, Nro_Grupo, objRsH_C!hornro, Fecha_Inicio, arrTempHsAD)
            
            'EAM- Calcula el tiempo que tarda la conversion antes de autorizar
            If depurar Then
                Flog.writeline Espacios(Tabulador * 1) & "  Tiempo conversión Antes de Autorizar : " & (GetTickCount - tiempoConversion)
            End If
        Else
            'FGZ - 01/08/2014 -----------------
            Nivel_Tab_Log = Nivel_Tab_Log + 1
            'FGZ - 01/08/2014 -----------------
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Realiza Conversion de Horas para el Tipo de Hora Nro: " & objRsH_C!thnro
            End If
            
            Flog.writeline Espacios(Tabulador * 1) & "  Tiempo DIFERENCIA: " & (GetTickCount - tiempoConversion)
            Nivel_Tab_Log = Nivel_Tab_Log - 1
        End If

         objRsH_C.MoveNext
    Loop
    
    
    
    If UsaConversionHoras And (UBound(arrTempHsAD) > 0) Then
        Call CrearAD(P_NroTer, arrTempHsAD)
    End If
    
    usaCompensacionAP = False
    Call Politica(500)

    If usaCompensacionAP Then
        Select Case st_Opcion
            Case 0, 1: 'EAM- Compensación de hs configurado en el turno
                If p_turcomp Then
                    Call Compensar_Horas(p_fecha, P_NroTer)
                End If
            Case 2: 'EAM- Compensación de hs por partes diarios
                Call Compensar_HorasPorParte(p_fecha, P_NroTer)
                
        End Select
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

    


