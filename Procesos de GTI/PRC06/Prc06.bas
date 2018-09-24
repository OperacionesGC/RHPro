Attribute VB_Name = "Prc06"
Option Explicit

'Version: 1.01
'Const Version = 1.01    '
'Const FechaVersion = "11/10/2005"

'Const Version = 1.02    'Controla que el periodo de gti no este cerrado
'Const FechaVersion = "19/12/2005"

'Const Version = 1.03    'Control de tope de dias teniendo en cuenta mes de febrero
'Const FechaVersion = "28/02/2006"

'Const Version = 1.04    ' LA - Modificacion de Tope_horas_Extras(): control de Nulo para los Tipos de Horas Excedentes.
'Const FechaVersion = "16/06/2006"

'Const Version = 1.05    'FGZ - Se agregó la politica 500 de compensacionde horas en AP
'Const FechaVersion = "02/11/2006"

'Const Version = 1.06    'FGZ - Problemas en las compensaciones. no estaba actualizando bien las horas compensadas parcialmente.
'Const FechaVersion = "03/01/2007"

'Const Version = 1.07    'FGZ - Nueva Politica 585 (Tope general de Horas extras para convenio petroleros).
'Const FechaVersion = "11/01/2007"

'Const Version = 1.08    'FGZ - Politica 585: Problemas cuando la estructura no esta activa en la totalidad del proceso y no esta abierta.
'Const FechaVersion = "25/01/2007"

'Const Version = "3.00"
'Const FechaVersion = "04/06/2007"
''Modificaciones: FGZ
''      Mejoras generales de performance.

'Const Version = "3.01"
'Const FechaVersion = "02/10/2007"
''Modificaciones: FGZ
''       Se modifico la funcion Feriado para que busque en todas las estructuras asignadas en la pol. de alcance para GTI ya que buscaba solo en la primera. CAS-04896

'Const Version = "3.02"
'Const FechaVersion = "21/01/2008"
''Modificaciones: FGZ
''       Se agrego la Politica 892 - Ajustes de AP.
''       Se agrego la Politica 893 - Ajustes de AP SMT.

'Const Version = "3.03"
'Const FechaVersion = "16/04/2008"
''Modificaciones: FGZ
''   Modulo politicas: sub Cargar_DetallePoliticas.
''           Se cambió en el where el <> '' por IS NOT NULL

'Const Version = "3.04"
'Const FechaVersion = "26/06/2008"
''Modificaciones: FGZ
''   politica 893:   'FAF en SMT
''           Se Agregaron unos not NULL
''   Compensaciones:
''           Se corriegieron errores detectados en AGD

'Const Version = "3.05"
'Const FechaVersion = "21/01/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion
''   Politica 580 - Topeo de hs extras (en realidad se modifico el procedimiento que hace el topeo)


'Const Version = "3.06"
'Const FechaVersion = "25/03/2009"
'Modificaciones: FGZ
'   Politica con alcance por estructura e indiciduales. No las estaba resolviendo

'Const Version = "3.07"
'Const FechaVersion = "21/10/2009"
''Modificaciones: MB
''   Error en Encriptacion de string de conexion faltaba c_seed

''--------------------------------------------------
'Const Version = "4.00"
'Const FechaVersion = "18/11/2009"
''Modificaciones: FGZ
''    Cambios Importantes
''        ALTER table gti_horcumplido add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_hishc add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_acumdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_hisad add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_achdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_his_achdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_det add horas varchar(10) null default ‘0:00’
'
''   Se agregó 1 campo en varias tablas para agregar la funcionalidad de que el resultado
''       se pueda expresar en distintas unidades (Valor decimal o en horas y minutos)
''   -----
''   OBS:
''   -----
''       El proceso de generacion de novedades de GTI permanece sin cambios, es decir,
''           el alcance de las modificaciones afectan hasta el proceso de Acumulado Parcial.
''       La cuenta corriente de horas se sigue manejando en valores decimales solamente.

'Const Version = "4.01"
'Const FechaVersion = "09/12/2009"
''Modificaciones: FGZ
''    se corrigió el procedimiento de control de versiones.

'Const Version = "4.02"
'Const FechaVersion = "17/05/2010"
''Modificaciones: FGZ
''    Politica 585 - Tope general de Horas
''   Se hicieron varias correccoiones


'----------------------------
'Const Version = "5.00"
'Const FechaVersion = "15/06/2010"
''Modificaciones: FGZ
''    Control por entradas fuera de termino.
''           Antes
''               cuando se queria procesar algo en una fecha que caia en un periodo cerrado no se procesaba.
''           Ahora
''               Se puede reprocear un periodo cerrado solo cuando se aprueba una entrada fuera de termino.
''               Para ese reprocesamiento solo se tendrá en cuenta todas las entradas fuera de termino aprobadas.


'Const Version = "5.01"
'Const FechaVersion = "07/10/2010"
''Modificaciones: FGZ
''       Se agrego la Politica 894 - Redondeo de horas en AP.

'Const Version = "5.02"
'Const FechaVersion = "03/06/2011"
''Modificaciones: FGZ
''       Se recompiló por un error detectado en un modulo general de politica aunque no asociado a este proceso.


Const Version = "5.03"
Const FechaVersion = "21/06/2011"
'Modificaciones: FGZ
'    Se agregaó el control de firmas a las novedades horarias
'       Se modifico:
'           Buscar_Turno
'           Buscar_Turno_nuevo

'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

' Variables globales necesarias para la integracion con
' el modulo de políticas

Global G_traza As Boolean
Global fec_proc As Integer

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

'Global NroProceso As Long

'Fin de globales

Global InicioTrimestre As Date
Global FinTrimestre As Date
Global Separador As String
Global Archivo As String
Global HorasSemestre As Single
Global MinHorasExtras As Single
Global THorasTrimestre As Integer
Global THorasMinimo As Integer
Global THorasSaldo As Integer
Global THorasAcumula As Integer
Global THorasPaga As Integer

Dim objBTurno As New BuscarTurno
Global objFeriado As New Feriado





Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   Nro_Grupo = T.Empleado_Grupo
   Nro_Justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
   
End Sub



Public Sub PRC06(ByVal Nro_Cab As Long, ByVal Desde As Date, ByVal Hasta As Date, ByVal Nro_tpr As Long, ByVal Nro_Pro As Long)
Dim sumahoras As Single
Dim TotHorHHMM As String

    
    StrSql = "delete from gti_det where cgtinro = " & Nro_Cab
    objConn.Execute StrSql, , adExecuteNoRecords
     
    StrSql = "SELECT gti_acumdiario.thnro, gti_acumdiario.ternro,gti_cab.cgtinro, SUM(adcanthoras) as Sumahoras FROM gti_cab "
    StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = gti_cab.ternro "
    StrSql = StrSql & " INNER JOIN gti_tpro_th ON gti_tpro_th.thnro = gti_acumdiario.thnro "
    StrSql = StrSql & " WHERE (gti_cab.cgtinro =" & Nro_Cab & ") and (gti_cab.gpanro = " & Nro_Pro & ") "
    StrSql = StrSql & " AND ( " & ConvFecha(Desde) & " <= gti_acumdiario.adfecha  AND gti_acumdiario.adfecha <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & "(gti_tpro_th.gtprocnro = " & Nro_tpr & ")"
    StrSql = StrSql & " GROUP BY gti_acumdiario.thnro, gti_acumdiario.ternro, gti_cab.cgtinro "
    StrSql = StrSql & " ORDER BY gti_acumdiario.thnro, gti_acumdiario.ternro, gti_cab.cgtinro "
    OpenRecordset StrSql, objRs
    
    If depurar Then
        Flog.writeline "SQL :" & StrSql
    End If
    If objRs.EOF Then
        If depurar Then
            Flog.writeline "No se encontraron datos"
        End If
    End If
    
    Do While Not objRs.EOF
        TotHorHHMM = CHoras(objRs!sumahoras, 60)
        
        StrSql = "INSERT INTO gti_det(cgtinro,thnro,horas,dgticant) VALUES (" & _
                 objRs!cgtinro & "," & objRs!thnro & "," & TotHorHHMM & "," & objRs!sumahoras & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If depurar Then
            Flog.writeline "Insertando: " & objRs!cgtinro & " " & objRs!thnro & " " & objRs!sumahoras & ". Del " & Desde & " al " & Hasta
        End If
        objRs.MoveNext
    Loop
    If depurar Then
        Flog.writeline "Programa Adicional :" & Now
    End If
    Call ProgAdicional(Nro_Cab, Desde, Hasta, Nro_tpr, Nro_Pro)
    
End Sub




Private Sub ProgAdicional(Nro_Cab As Long, Desde As Date, Hasta As Date, Nro_tpr As Long, Nro_Pro As Long)
Dim ObjEmp As New ADODB.Recordset
Dim NroTer As Long
Dim Legajo As Long


    'Dentro de este programa Adicional, se llaman a Politicas que son las que realzan los trabajos sobre los Acumulados Parciales
    'Politica que determina si existen topes para las extras
    StrSql = " SELECT empleado.ternro, empleado.empleg FROM empleado"
    StrSql = StrSql & " INNER JOIN gti_cab ON gti_cab.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE gti_cab.cgtinro = " & Nro_Cab
    OpenRecordset StrSql, ObjEmp
        
    If ObjEmp.EOF Then
        Exit Sub
    End If
    
    Empleado.Ternro = ObjEmp!Ternro
    Empleado.Legajo = ObjEmp!EmpLeg
    
    p_fecha = Desde

    Set objBTurno.Conexion = objConn
    Set objBTurno.ConexionTraza = CnTraza
    objBTurno.Buscar_Turno p_fecha, Empleado.Ternro, depurar
    initVariablesTurno objBTurno
    If Not tiene_turno And Not Tiene_Justif Then
        If depurar Then
            Flog.writeline "Empleado " & Empleado.Legajo & " No tiene turno y no tiene justificacion"
        End If
        Exit Sub
    End If
    
    'FGZ - 30/10/2006
    usaCompensacionAP = False
    Call Politica(500)
    If usaCompensacionAP Then
        If depurar Then
            Flog.writeline "Inicio Compensacion"
        End If
        Call Compensar_HorasXDia(Nro_Cab, Empleado.Ternro)
        If depurar Then
            Flog.writeline "Fin Compensacion"
        End If
    End If
    
    
    ' Topea las Horas Extras segun Configuracion
    usaTopesHorasExtras = False
    Call Politica(580)
    If usaTopesHorasExtras Then
        If depurar Then
            Flog.writeline "Topea las Horas Extras segun Configuracion"
        End If
        Call Tope_Horas_Extras(Nro_Cab, Empleado.Ternro, Empleado.Legajo, Nro_Turno)
    End If

    'Topea las Horas Extras segun Configuracion
    usaTopesGralHorasExtras = False
    Call Politica(585)
    If usaTopesGralHorasExtras Then
        If depurar Then
            Flog.writeline "Topea las Horas Extras segun Configuracion"
        End If
        Call Primeras_Horas_Extras(Nro_Cab, Desde, Hasta, Nro_tpr, Nro_Pro)
    End If

    ' Topea las Horas Normales segun Configuracion - Expofrut
    usaExcedentesHorasNormales = False
    Call Politica(850)
    If usaExcedentesHorasNormales Then
        If depurar Then
            Flog.writeline "Topea las Horas Normales segun Configuracion"
        End If
        Call Tope_Horas_Normales(Nro_Cab, Empleado.Ternro, Empleado.Legajo, Nro_Turno)
    End If

    usaControlDias = False
    Call Politica(650)
    If usaControlDias Then
        If depurar Then
            Flog.writeline "Control de dias"
        End If
        Call Control_Cant_Dias(Nro_Cab, Empleado.Ternro, Empleado.Legajo, Nro_Turno, Hasta)
    End If
    
    'FGZ - 07/10/2010 ---------------------------
    'Redondeo de horas
    UsaRedondeoHoras = False
    Call Politica(894)
    If UsaRedondeoHoras Then
        If depurar Then
            Flog.writeline "Redondeo"
        End If
        Call Redondeo_Horas(Nro_Cab, Empleado.Ternro, Empleado.Legajo, Nro_Turno, Hasta)
    End If
    'FGZ - 07/10/2010 ---------------------------
    
End Sub


Private Sub Tope_Horas_Extras(Nro_Cab As Long, Ternro As Long, Legajo As Long, Nro_Turno As Long)

Dim Tope As Single
Dim Horas_50 As Single
Dim Horas_100 As Single
Dim Horas_exed_50 As Single
Dim Horas_exed_100 As Single
Dim th50 As Integer
Dim th100 As Integer
Dim ThEx50 As Integer
Dim ThEx100 As Integer
Dim Limite As Single
Dim Act50 As Boolean
Dim Act100 As Boolean

Dim objRs As New ADODB.Recordset
Dim TotHorHHMM As String



   Tope = 0
   Horas_50 = 0
   Horas_100 = 0
   Horas_exed_50 = 0
   Horas_exed_100 = 0
   
'   th50 = 1
'   th100 = 2
'   ThEx50 = 36
'   ThEx100 = 35
'   limite = 30
   
   ' Prueba
   
   
'   Tipo de Hora Configurable - 16 - Horas Extras 50%
   StrSql = "SELECT thnro FROM gti_config_tur_hor WHERE conhornro = 16 "
   StrSql = StrSql & " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
   OpenRecordset StrSql, objRs
   If Not objRs.EOF Then
        th50 = objRs!thnro
   Else
        If depurar Then
            Flog.writeline Espacios(Tabulador) & "El turno no tiene configurado el tipo de 16 - Horas Extras Simples"
        End If
        Exit Sub
   End If

'   Tipo de Hora Configurable - 17 - Horas Extras 100%
   StrSql = "SELECT thnro FROM gti_config_tur_hor WHERE conhornro = 17 "
   StrSql = StrSql & " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
   OpenRecordset StrSql, objRs
   If Not objRs.EOF Then
        th100 = objRs!thnro
   Else
        If depurar Then
            Flog.writeline Espacios(Tabulador) & "El turno no tiene configurado el tipo de 17 - Horas Extras Dobles"
        End If
        Exit Sub
   End If
    
    ' FGZ - 04/08/2003
    ' Cambio estruc_actual por his_estructura
   StrSql = " SELECT thestrexcmen,thestrlimmen FROM tiphora_estr "
   StrSql = StrSql & " INNER JOIN his_estructura ON tiphora_estr.estrnro = his_estructura.estrnro AND "
   StrSql = StrSql & " his_estructura.tenro = tiphora_estr.tenro AND his_estructura.ternro = " & Empleado.Ternro
   StrSql = StrSql & " WHERE thnro = " & th50 & " AND his_estructura.htethasta IS NULL "
   OpenRecordset StrSql, objRs
   If Not objRs.EOF Then
        'LA - 16-06-2006
        'Control de valor Nulo
        'ThEx50 = objRs!thestrexcmen
        If EsNulo(objRs!thestrexcmen) Then
            ThEx50 = 0
        Else
            ThEx50 = objRs!thestrexcmen
        End If
        Limite = Limite + objRs!thestrlimmen
   Else
        'Exit Sub
   End If

    ' FGZ - 04/08/2003
    ' Cambio estruc_actual por his_estructura
   StrSql = " SELECT thestrexcmen, thestrlimmen FROM tiphora_estr "
   StrSql = StrSql & " INNER JOIN his_estructura ON tiphora_estr.estrnro = his_estructura.estrnro AND "
   StrSql = StrSql & " his_estructura.tenro = tiphora_estr.tenro AND his_estructura.ternro = " & Empleado.Ternro
   StrSql = StrSql & " WHERE thnro = " & th100 & " AND his_estructura.htethasta IS NULL "
   OpenRecordset StrSql, objRs
   If Not objRs.EOF Then
        'LA - 16-06-2006
        'Control de valor Nulo
        'ThEx100 = objRs!thestrexcmen
        If EsNulo(objRs!thestrexcmen) Then
            ThEx100 = 0
        Else
            ThEx100 = objRs!thestrexcmen
        End If
        Limite = Limite + objRs!thestrlimmen
   Else
        'Exit Sub
   End If
   
   'FGZ - 21/01/2009 - Se modificó estas condiciones porque sino hay que tener ambos topes configurados
   If (ThEx50 + ThEx100) = 0 Then
        If depurar Then
            Flog.writeline Espacios(Tabulador) & " Los tipos de Hs " & th50 & " y " & th100 & " No tienen topes configurados ==> no hay topeo"
        End If
        Exit Sub
   End If
   
   StrSql = " SELECT thnro, dgticant FROM gti_det "
   StrSql = StrSql & " WHERE (thnro IN (" & th50 & "," & th100 & ")) AND cgtinro = " & Nro_Cab
   StrSql = StrSql & " ORDER BY thnro "
   OpenRecordset StrSql, objRs
   
   Act50 = False
   Act100 = False
         
   Do While Not objRs.EOF

    If Tope <= Limite Then
        
        If Tope + objRs!dgticant <= Limite Then

         Tope = Tope + objRs!dgticant

         If objRs!thnro = th50 Then
            Horas_50 = Horas_50 + objRs!dgticant
            Act50 = True
         End If

         If objRs!thnro = th100 Then
            Horas_100 = Horas_100 + objRs!dgticant
            Act100 = True
         End If

        Else

         If objRs!thnro = th50 Then
            Horas_50 = Horas_50 + (Limite - Tope)
            Horas_exed_50 = Horas_exed_50 + (Tope + objRs!dgticant - Limite)
            Tope = Limite
            Act50 = True
         End If

         If objRs!thnro = th100 Then
            Horas_100 = Horas_100 + (Limite - Tope)
            Horas_exed_100 = Horas_exed_100 + (Tope + objRs!dgticant - Limite)
            Tope = Limite
            Act100 = True
         End If
         
        End If
        
      Else

        If objRs!thnro = th50 Then
            Horas_exed_50 = Horas_exed_50 + objRs!dgticant
            Act50 = True
        End If
    
        If objRs!thnro = th100 Then
            Horas_exed_100 = Horas_exed_100 + objRs!dgticant
            Act100 = True
        End If
        

      End If
      
      objRs.MoveNext
    Loop
    
    If Act50 Then
        If Horas_50 <> 0 Then
                TotHorHHMM = CHoras(Horas_50, 60)
                
                StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM & ",dgticant = " & Horas_50
                StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & th50
                objConn.Execute StrSql, , adExecuteNoRecords
        Else
                StrSql = "DELETE FROM gti_det WHERE cgtinro = " & Nro_Cab & " AND thnro = " & th50
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    If Act100 Then
        If Horas_100 <> 0 Then
                TotHorHHMM = CHoras(Horas_100, 60)
                
                StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM & ",dgticant = " & Horas_100
                StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & th100
                objConn.Execute StrSql, , adExecuteNoRecords
        Else
                StrSql = "DELETE FROM gti_det WHERE cgtinro = " & Nro_Cab & " AND thnro = " & th100
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
        
    If Horas_exed_50 <> 0 Then
        ' LA - 16-06-2006
        ' Control de que que se halla configurado excedente
        If ThEx50 = 0 Then
            If depurar Then
                Flog.writeline Espacios(Tabulador) & " No se configuro el Tipo de Hora de Excedente para el Tipo de Hora extra al 50%"
            End If
        Else
                
                TotHorHHMM = CHoras(Horas_exed_50, 60)
                
               StrSql = "INSERT INTO gti_det(cgtinro,thnro,horas,dgticant) VALUES (" & _
                         Nro_Cab & "," & ThEx50 & "," & TotHorHHMM & "," & Horas_exed_50 & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    If Horas_exed_100 <> 0 Then
        ' LA - 16-06-2006
        ' Control de que que se halla configurado excedente
        If ThEx100 = 0 Then
            If depurar Then
                Flog.writeline Espacios(Tabulador) & " No se configuro el Tipo de Hora de Excedente para el Tipo de Hora extra al 100%"
            End If
        Else
                TotHorHHMM = CHoras(Horas_exed_100, 60)
                
                StrSql = "INSERT INTO gti_det(cgtinro,thnro,horas,dgticant) VALUES (" & _
                         Nro_Cab & "," & ThEx100 & "," & TotHorHHMM & "," & Horas_exed_100 & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
End Sub


Private Sub Tope_Horas_Normales(Nro_Cab As Long, Ternro As Long, Legajo As Long, Nro_Turno As Long)

Dim Tope As Single
Dim Horas_Normal As Single
Dim Horas_exced_Normal As Single
Dim ThNormal As Integer
Dim ThExNormal As Integer
Dim Limite As Single

Dim objRs As New ADODB.Recordset
Dim TotHorHHMM As String


   Tope = 0
   Horas_Normal = 0
   Horas_exced_Normal = 0
   
'   ThNormal = 3
'   ThExNormal = 1
'   limite = 48
   
'   Tipo de Hora Configurable - 1 - Horas Obligatorias
    StrSql = "SELECT thnro FROM gti_config_tur_hor WHERE conhornro = 1 "
    StrSql = StrSql & " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        ThNormal = objRs!thnro
    Else
        Exit Sub
    End If

    ' FGZ - 04/08/2003
    ' Cambio estruc_actual por his_estructura
    StrSql = " SELECT thestrexcmen,thestrlimmen FROM tiphora_estr "
    StrSql = StrSql & " INNER JOIN his_estructura ON tiphora_estr.estrnro = his_estructura.estrnro AND "
    StrSql = StrSql & " his_estructura.tenro = tiphora_estr.tenro AND his_estructura.ternro = " & Ternro
    StrSql = StrSql & " WHERE thnro = " & ThNormal & " AND his_estructura.htethasta IS NULL "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        If EsNulo(objRs!thestrexcmen) Then
            ThExNormal = 0
        Else
            ThExNormal = objRs!thestrexcmen
        End If
        Limite = Limite + objRs!thestrlimmen
    Else
         Exit Sub
    End If

   
    StrSql = " SELECT thnro, dgticant FROM gti_det "
    StrSql = StrSql & " WHERE thnro = " & ThNormal & " AND cgtinro = " & Nro_Cab
    StrSql = StrSql & " ORDER BY thnro "
    OpenRecordset StrSql, objRs
    Do While Not objRs.EOF
        If Tope <= Limite Then
            If Tope + objRs!dgticant <= Limite Then
                Tope = Tope + objRs!dgticant
                If objRs!thnro = ThNormal Then Horas_Normal = Horas_Normal + objRs!dgticant
            Else
                If objRs!thnro = ThNormal Then
                    Horas_Normal = Horas_Normal + (Limite - Tope)
                    Horas_exced_Normal = Horas_exced_Normal + (Tope + objRs!dgticant - Limite)
                    Tope = Limite
                End If
            End If
        Else
            If objRs!thnro = ThNormal Then Horas_exced_Normal = Horas_exced_Normal + objRs!dgticant
        End If
        
        objRs.MoveNext
    Loop

    If Horas_Normal <> 0 Then
        TotHorHHMM = CHoras(Horas_Normal, 60)
    
        StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM & ",dgticant = " & Horas_Normal
        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & ThNormal
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    If Horas_exced_Normal <> 0 Then
        TotHorHHMM = CHoras(Horas_exced_Normal, 60)
        StrSql = "INSERT INTO gti_det(cgtinro,thnro,horas,dgticant) VALUES (" & _
                 Nro_Cab & "," & ThExNormal & "," & TotHorHHMM & "," & Horas_exced_Normal & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

End Sub


Public Sub Main()
Dim Nro_Cab As Long
Dim FDesde As Date
Dim FHasta As Date
Dim Nro_tpr As Long
Dim Nro_Pro As Long
Dim strcmdLine As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim NroProceso As Long
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Fecha As Date
Dim objrsEmpleado As New ADODB.Recordset
Dim objRs As New ADODB.Recordset
Dim Progreso As Single
Dim IncPorc As Single
Dim ListaPar

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim Cantidad As Long

Dim ArrParametros

    
    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(PathFLog & "PRC06" & "-" & NroProceso & ".log", True)

    Cantidad_de_OpenRecordset = 0
    Cantidad_Call_Politicas = 0
    
    'Activo el manejador de errores
    On Error GoTo CE
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
     
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set objFechasHoras.Conexion = objConn
    
             
    'FGZ - 25/03/2009 ---------
    
    'Cambié el query, le agregué los campos ternro y empleg
    'StrSql = " SELECT batch_proceso.IdUser, batch_proceso.bprcparam, gti_procacum.gpadesde,gti_procacum.gpahasta,gti_cab.gpanro,gti_cab.cgtinro,batch_proceso.bpronro,gti_procacum.gtprocnro FROM batch_proceso " & _
    '         " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro " & _
    '         " INNER JOIN gti_cab ON  gti_cab.gpanro = batch_procacum.gpanro " & _
    '         " INNER JOIN gti_procacum ON batch_procacum.gpanro = gti_procacum.gpanro " & _
    '         " INNER JOIN gti_per ON gti_per.pgtinro = gti_procacum.pgtinro AND gti_per.pgtiestado = -1 " & _
    '         " INNER JOIN batch_empleado ON batch_empleado.ternro = gti_cab.ternro AND batch_empleado.bpronro = batch_proceso.bpronro" & _
    '         " WHERE batch_proceso.bpronro = " & NroProceso
    
    
    'FGZ - 03/06/2010 ---------------------------------------------------------------------
    StrSql = "SELECT iduser,bprcfecdesde,bprcfechasta,bprcparam FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
        Flog.writeline Espacios(Tabulador * 2) & "Usuario: " & objRs!IdUser
        Flog.writeline Espacios(Tabulador * 2) & "Desde: " & objRs!bprcfecdesde
        FechaDesde = objRs!bprcfecdesde
        Flog.writeline Espacios(Tabulador * 2) & "Hasta: " & objRs!bprcfechasta
        FechaHasta = objRs!bprcfechasta
        'FGZ - 01/09/2006
        Flog.writeline Espacios(Tabulador * 2) & "bprcparam: " & objRs!bprcparam
        If Not EsNulo(objRs!bprcparam) Then
            If InStr(1, objRs!bprcparam, ".") <> 0 Then
                ListaPar = Split(objRs!bprcparam, ".", -1)
                depurar = IIf(IsNumeric(ListaPar(0)), CBool(ListaPar(0)), False)
                'FGZ - 18/05/2010 - se le agregó un nuevo parametro ----
                If UBound(ListaPar) > 1 Then
                    If Not EsNulo(ListaPar(2)) Then
                        ReprocesarFT = IIf(IsNumeric(ListaPar(2)), CBool(ListaPar(2)), False)
                    Else
                        ReprocesarFT = False
                    End If
                Else
                    ReprocesarFT = False
                End If
                'FGZ - 18/05/2010 - se le agregó un nuevo parametro ----
            Else
                depurar = False
                ReprocesarFT = False
            End If
        Else
            depurar = False
            ReprocesarFT = False
        End If
        Flog.writeline Espacios(Tabulador * 2) & "Log detallado: " & depurar
        Flog.writeline Espacios(Tabulador * 2) & "Reprocesar Periodo Cerrado: " & ReprocesarFT
    Else
        Exit Sub
    End If
  
  
    'Busco los procesos de AP a procesar
    StrSql = " SELECT batch_proceso.IdUser, batch_proceso.bprcparam, gti_procacum.gpadesde,gti_procacum.gpahasta,gti_cab.gpanro,gti_cab.cgtinro"
    StrSql = StrSql & " ,empleado.empleg,gti_cab.ternro,batch_proceso.bpronro,gti_procacum.gtprocnro, gti_per.pgtiestado "
    StrSql = StrSql & " FROM batch_proceso "
    StrSql = StrSql & " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro "
    StrSql = StrSql & " INNER JOIN gti_cab ON  gti_cab.gpanro = batch_procacum.gpanro "
    StrSql = StrSql & " INNER JOIN gti_procacum ON batch_procacum.gpanro = gti_procacum.gpanro "
    StrSql = StrSql & " INNER JOIN gti_per ON gti_per.pgtinro = gti_procacum.pgtinro "
    If Not ReprocesarFT Then
        StrSql = StrSql & " AND gti_per.pgtiestado = -1 "
    End If
    StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = gti_cab.ternro AND batch_empleado.bpronro = batch_proceso.bpronro"
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = gti_cab.ternro"
    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & NroProceso
    'FGZ - 25/03/2009 ---------
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        FechaDesde = objRs!gpadesde
        FechaHasta = objRs!gpahasta
        Cantidad = objRs.RecordCount
    
'        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
'        Flog.writeline Espacios(Tabulador * 1) & "Usuario: " & objRs!IdUser
'        Flog.writeline Espacios(Tabulador * 1) & "Desde: " & FechaDesde
'        Flog.writeline Espacios(Tabulador * 1) & "Hasta: " & FechaHasta
'        Flog.writeline Espacios(Tabulador * 1) & "bprcparam: " & objRs!bprcparam
'        If Not EsNulo(objRs!bprcparam) Then
'            If InStr(1, objRs!bprcparam, ".") <> 0 Then
'                ListaPar = Split(objRs!bprcparam, ".", -1)
'                depurar = IIf(IsNumeric(ListaPar(0)), CBool(ListaPar(0)), False)
'                'FGZ - 18/05/2010 - se le agregó un nuevo parametro ----
'                ReprocesarFT = IIf(IsNumeric(ListaPar(1)), CBool(ListaPar(1)), False)
'                'FGZ - 18/05/2010 - se le agregó un nuevo parametro ----
'            Else
'                depurar = False
'                ReprocesarFT = False
'            End If
'        Else
'            depurar = False
'            ReprocesarFT = False
'        End If
'        Flog.writeline Espacios(Tabulador * 2) & "Log detallado: " & depurar
'        Flog.writeline Espacios(Tabulador * 2) & "Reprocesar Periodo Cerrado: " & ReprocesarFT
    Else
        Cantidad = 1
        Flog.writeline
        Flog.writeline "No se encontraron procesos. Revisar que el periodo se encuentre activo."
        Flog.writeline
        Flog.writeline "SQL: " & StrSql
    End If
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio    :" & Now
    
    Fecha = FechaDesde
    IncPorc = (100 / Cantidad)
    Progreso = 0
    
    'FGZ - Mejoras ----------
    Call Inicializar_Globales
    'FGZ - Mejoras ----------
    
    Do While Not objRs.EOF
inicio:

        'FGZ - 25/03/2009 ------------------------
        Empleado.Ternro = objRs!Ternro
        Empleado.Legajo = objRs!EmpLeg

        Call Cargar_PoliticasEstructuras(FechaDesde)
        Call Cargar_PoliticasIndividuales
        'FGZ - 25/03/2009 ------------------------

        
        If depurar Then
            Flog.writeline "Inicio Cabecera:" & objRs!cgtinro & " " & FechaDesde & " al " & FechaHasta
            If CBool(objRs!pgtiestado) Then
                Flog.writeline "Importante. El proceso correspondiente a un periodo cerrado. Procesamiento Fuera de Término Aprobado."
            End If
        End If
        
        Call PRC06(objRs!cgtinro, objRs!gpadesde, objRs!gpahasta, objRs!gtprocnro, objRs!gpanro)
            
        
        'FGZ - 18/01/2008 - Se agrego esta politica de ajuste --------
        ProcesoAP = objRs!gpanro
        CabeceraAP = objRs!cgtinro
        
        Call Politica(892)
        Call Politica(893)
        'FGZ - 18/01/2008 - Se agrego esta politica de ajuste --------
        
        usaDesgloseAP = False
        Call Politica(590)
        If usaDesgloseAP Then
            If depurar Then
                Flog.writeline "Desgloce Acumulado Parcial :" & Now
            End If
            Call Desglose_ACParcial(objRs!cgtinro, FechaDesde, FechaHasta)
        End If
            
siguiente:
            objRs.MoveNext
            
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Loop
        
    StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
    If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
    If objConn.State = adStateOpen Then objConn.Close
    Set objConn = Nothing
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    Set objConnProgreso = Nothing
    
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Fin       :" & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
'    Flog.writeline "Cantidad de llamadas a EsFeriado    : " & Cantidad_Feriados
'    Flog.writeline "Cantidad de llamadas a BuscarTurno  : " & Cantidad_Turnos
'    Flog.writeline "Cantidad de llamadas a BuscarDia    : " & Cantidad_Dias
'    Flog.writeline
'    Flog.writeline "Cantidad de dias procesados         : " & Cantidad_Empl_Dias_Proc
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
    
    Exit Sub

CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Cabecera" & " " & Fecha
    Flog.writeline Espacios(Tabulador * 0) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub

Private Sub Compensar_HorasXDia(ByVal Nro_Cab As Long, ByVal NroTer As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera las compensacion para el rango de fechas del proceso.
' Autor      : FGZ - 30/10/2006
' Ultima Mod.: FGZ - 03/01/2006 - no estaba actualizando bien las horas compensadas parcialmente
' ---------------------------------------------------------------------------------------------
Dim oblig As Single
Dim th_generar As Integer
Dim Cant As Single
Dim fecha_desde     As Date
Dim hora_desde      As String
Dim fecha_hasta     As Date
Dim hora_hasta      As String

Dim rsComp As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsAPComp As New ADODB.Recordset
Dim rsAp As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset

Dim canthoras As Single
Dim acumula As Boolean
Dim SumaAux As Single

Dim Aux_HSAComp As Single
Dim TotHorHHMM As String

    'el 20 es porque es el codigo de la horas compensadas
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 20 "
    StrSql = StrSql & " AND turnro = " & Nro_Turno
    StrSql = StrSql & " ORDER BY conhornro ASC, turnro ASC"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        'se setea al tipo de horas que se compensan las horas compensables
        th_generar = rs!thnro
    Else
        'Entrada en la traza
        If depurar Then
            Flog.writeline Espacios(Tabulador * 3) & "No esta configurado el Tipo de Hora Compensada para el Turno: " & Str(Nro_Turno)
            GeneraTraza Empleado.Ternro, Date, "No esta configurado el Tipo de Hora Compensada para el Turno:", Str(Nro_Turno)
        End If
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM gti_compnsbl WHERE "
    StrSql = StrSql & " turnro = " & Nro_Turno
    StrSql = StrSql & " ORDER BY compsblorden ASC"
    OpenRecordset StrSql, rsComp
    Do While Not rsComp.EOF
        StrSql = " SELECT thnro, dgticant FROM gti_det "
        StrSql = StrSql & " WHERE thnro = " & rsComp!thnro
        StrSql = StrSql & " AND cgtinro = " & Nro_Cab
        OpenRecordset StrSql, rsAPComp
        If rsAPComp.EOF Then
            GoTo Continuar_rsComp
        Else
            'FGZ - 03/01/2006
            Aux_HSAComp = rsAPComp!dgticant
        
            StrSql = "SELECT * FROM gti_acompsar "
            StrSql = StrSql & " WHERE compsblnro = " & rsComp!compsblnro
            StrSql = StrSql & " ORDER BY acomporden ASC"
            OpenRecordset StrSql, rs
            Do While Not rs.EOF
                StrSql = "SELECT * FROM gti_det "
                StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
                StrSql = StrSql & " AND thnro = " & rs!thnro
                OpenRecordset StrSql, rsAp
                If rsAp.EOF Then
                    GoTo continuar_rs
                Else
                     If Aux_HSAComp > rsAp!dgticant Then
                        Cant = Cant + (rsAp!dgticant * (rs!acompptje / 100))
                        'FGZ - 20/06/2008 - se cambio esta estaba mal
                        'canthoras = (Aux_HSAComp - rsAP!dgticant) * (rs!acompptje / 100)
                        canthoras = Aux_HSAComp - (rsAp!dgticant * (rs!acompptje / 100))
                      
                        TotHorHHMM = CHoras(canthoras, 60)
                        
                        StrSql = "UPDATE gti_det SET horas =" & TotHorHHMM & ",dgticant = " & canthoras
                        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
                        StrSql = StrSql & " AND thnro = " & rsComp!thnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        StrSql = "DELETE FROM gti_det "
                        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
                        StrSql = StrSql & " AND thnro = " & rs!thnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Aux_HSAComp = Aux_HSAComp - rsAp!dgticant
                    Else
                        If Aux_HSAComp < rsAp!dgticant Then
                            'FGZ - 20/06/2008 - se cambio esto porque estaba mal
                            'Cant = Cant + (rsAPComp!dgticant * (rs!acompptje / 100))
                            'canthoras = (rsAP!dgticant - Aux_HSAComp) * (rs!acompptje / 100)
                            
                            Cant = Cant + (Aux_HSAComp * (rs!acompptje / 100))
                            canthoras = rsAp!dgticant - (Aux_HSAComp * (rs!acompptje / 100))
                            
                            TotHorHHMM = CHoras(canthoras, 60)
                            StrSql = "UPDATE gti_det SET horas =" & TotHorHHMM & ",dgticant = " & canthoras
                            StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
                            StrSql = StrSql & " AND thnro = " & rs!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            StrSql = "DELETE FROM gti_det "
                            StrSql = StrSql & " WHERE  cgtinro = " & Nro_Cab
                            StrSql = StrSql & "AND thnro = " & rsComp!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Aux_HSAComp = 0
                        Else
                            Cant = Cant + (rsAp!dgticant * (rs!acompptje / 100))
                            
                            StrSql = "DELETE FROM gti_det "
                            StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
                            StrSql = StrSql & " AND thnro = " & rsComp!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            StrSql = "DELETE FROM gti_det "
                            StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
                            StrSql = StrSql & " AND thnro = " & rs!thnro
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
    ' aca tenemos que revisar primero si ese tipo de Hora ya està insertado
    ' si es asi ==> tengo que modificar el registro sumandole las horas
    ' sino lo inserto
    
    StrSql = "SELECT * FROM gti_det "
    StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
    StrSql = StrSql & " AND thnro = " & th_generar
    OpenRecordset StrSql, rsAux
    If rsAux.EOF Then
        TotHorHHMM = CHoras(Cant, 60)
        
        'Ese tipo de hora no lo tiene ==> lo inserto
        StrSql = " INSERT INTO gti_det(cgtinro,thnro,horas,dgticant) VALUES ("
        StrSql = StrSql & Nro_Cab & "," & th_generar & "," & TotHorHHMM & "," & Cant & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        'ese tipo de hora ya la tiene ==> le actualizo el total de horas
        SumaAux = rsAux!dgticant + Cant
        TotHorHHMM = CHoras(SumaAux, 60)
        
        StrSql = "UPDATE gti_det SET horas =" & TotHorHHMM & ",dgticant = " & SumaAux
        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab
        StrSql = StrSql & " AND thnro = " & th_generar
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Cierro todo y libero
    If rsAux.State = adStateOpen Then rsAux.Close
    If rsComp.State = adStateOpen Then rsComp.Close
    If rs.State = adStateOpen Then rs.Close
    If rsAPComp.State = adStateOpen Then rsAPComp.Close
    If rsAp.State = adStateOpen Then rsAp.Close
    
    Set rsAux = Nothing
    Set rsComp = Nothing
    Set rs = Nothing
    Set rsAPComp = Nothing
    Set rsAp = Nothing
End Sub



Private Sub Control_Cant_Dias(ByVal Nro_Cab As Long, ByVal Ternro As Long, ByVal Legajo As Long, ByVal Nro_Turno As Long, ByVal Hasta As Date)
Dim Ajuste As Single
Dim Tope As Single
Dim SumaJornadas As Single
Dim objad As New ADODB.Recordset
Dim Anio_bisiesto As Boolean
Dim TotHorHHMM As String

   Tope = 30

   StrSql = " SELECT SUM(dgticant) AS SumaJornadas FROM gti_det "
   StrSql = StrSql & " WHERE (thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80,81,84,86)) AND cgtinro = " & Nro_Cab
   OpenRecordset StrSql, objRs
    If depurar Then
        Flog.writeline "Desde ControlCantDias: " & Nro_Cab & " " & objRs!SumaJornadas
    End If
   
   If objRs!SumaJornadas > Tope Then
       StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
       StrSql = StrSql & " AND ternro = " & Ternro
       StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80,81,84,86)"
       OpenRecordset StrSql, objad
       
       Do While objad.EOF
       
        Hasta = DateDiff("d", 1, Hasta)

        StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80,81,84,86)"
        OpenRecordset StrSql, objad
       
       Loop
        StrSql = "UPDATE gti_det SET dgticant = dgticant - 1"
        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objad!thnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Actualizo el campo horas
        StrSql = " SELECT dgticant FROM gti_det "
        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objad!thnro
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            TotHorHHMM = CHorasSF(objRs!dgticant, 60)
            
            StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM
            StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objad!thnro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Else ' FGZ - 01/03/2004 'Tengo que completar hasta el tope
        If (objRs!SumaJornadas < Tope) And Month(Hasta) = 2 Then
       
            Anio_bisiesto = EsBisiesto(Year(Hasta))
    
            StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
            StrSql = StrSql & " AND ternro = " & Ternro
            StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80)"
            OpenRecordset StrSql, objad
            
            Do While objad.EOF
              Hasta = DateDiff("d", 1, Hasta)
            
              StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
              StrSql = StrSql & " AND ternro = " & Ternro
              StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80)"
              OpenRecordset StrSql, objad
            Loop
            'Se recupera el tipo de hora a utilizar para completar
    
            If (objRs!SumaJornadas = 28.99) Then
                If Anio_bisiesto Then
                    Ajuste = 1.01
                Else
                    Ajuste = 2.01
                End If
            Else
                If Anio_bisiesto Then
                    Ajuste = 1#
                Else
                    Ajuste = 2#
                End If
            End If
            ' Esta asignación solo sirve para el 2004 que es bisiesto, y considera un problema
            ' de redondeo que hace que se pierda una centésima.
            ' No se puede directamente sumar lo que falta para llegar a 30, porque hay casos
            ' en donde NO deben llegar a 30 días porque han faltado. O.D.A. 02/03/2004
            
            If depurar Then
                Flog.writeline "Desde ControlCantDias sumando: " & Ajuste & " a " & Nro_Cab & " en " & objad!thnro
            End If
            
            StrSql = "UPDATE gti_det "
            StrSql = StrSql & " SET    gti_det.dgticant = gti_det.dgticant + " & Ajuste
            StrSql = StrSql & " WHERE  gti_det.cgtinro  = " & Nro_Cab
            StrSql = StrSql & " AND    gti_det.thnro    = " & objad!thnro
            objConn.Execute StrSql, , adExecuteNoRecords
            ' Se agrega al AP lo que falta para llegar al tope, en el primer tipo de hora
            ' encontrado que esté entre las consideradas
        
            'Actualizo el campo horas
            StrSql = " SELECT dgticant FROM gti_det "
            StrSql = StrSql & " WHERE  gti_det.cgtinro  = " & Nro_Cab
            StrSql = StrSql & " AND    gti_det.thnro    = " & objad!thnro
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                TotHorHHMM = CHorasSF(objRs!dgticant, 60)
                
                StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM
                StrSql = StrSql & " WHERE  gti_det.cgtinro  = " & Nro_Cab
                StrSql = StrSql & " AND    gti_det.thnro    = " & objad!thnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
       ' Este código solo es válido para los meses de Febrero.
       ' Hasta que se controle en forma automática, se anula. O.D.A. 10/03/2004
    End If
End Sub

Private Sub Control_Cant_Dias_old(ByVal Nro_Cab As Long, ByVal Ternro As Long, ByVal Legajo As Long, ByVal Nro_Turno As Long, ByVal Hasta As Date)
Dim Ajuste As Single
Dim Tope As Single
Dim SumaJornadas As Single
Dim objad As New ADODB.Recordset

   Tope = 30

   If objRs.State = adStateOpen Then objRs.Close
   
   StrSql = " SELECT SUM(dgticant) AS SumaJornadas FROM gti_det "
   StrSql = StrSql & " WHERE (thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80,81)) AND cgtinro = " & Nro_Cab
   
   OpenRecordset StrSql, objRs
   
   Flog.writeline "Desde ControlCantDias: " & Nro_Cab & " " & objRs!SumaJornadas
   
   If objRs!SumaJornadas > Tope Then
        
       StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
       StrSql = StrSql & " AND ternro = " & Ternro
       StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80)"
       OpenRecordset StrSql, objad
       
       Do While objad.EOF
       
        Hasta = DateDiff("d", 1, Hasta)

        StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80)"
        OpenRecordset StrSql, objad
       
       Loop
        
        StrSql = "UPDATE gti_det SET dgticant = dgticant - 1"
        StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objad!thnro
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Else ' FGZ - 01/03/2004
        If (objRs!SumaJornadas < Tope) Then
            'Tengo que completar hasta el tope
            
           StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
           StrSql = StrSql & " AND ternro = " & Ternro
           StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80)"
           OpenRecordset StrSql, objad
           
           Do While objad.EOF
             Hasta = DateDiff("d", 1, Hasta)
    
             StrSql = "SELECT thnro FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Hasta)
             StrSql = StrSql & " AND ternro = " & Ternro
             StrSql = StrSql & " AND thnro IN (54,58,59,60,61,62,63,64,65,66,67,68,69,70,75,79,80)"
             OpenRecordset StrSql, objad
           Loop
' Se recupera el tipo de hora a utilizar para completar

           If (objRs!SumaJornadas = 28.99) Then
             Ajuste = 1.01
           Else
             Ajuste = 1#
           End If
' Esta asignación solo sirve para el 2004 que es bisiesto, y considera un problema
' de redondeo que hace que se pierda una centésima.
' No se puede directamente sumar lo que falta para llegar a 30, porque hay casos
' en donde NO deben llegar a 30 días porque han faltado. O.D.A. 02/03/2004

           Flog.writeline "Desde ControlCantDias sumando: " & Ajuste & " a " & Nro_Cab & " en " & objad!thnro
                       
           StrSql = "UPDATE gti_det "
           StrSql = StrSql & " SET    gti_det.dgticant = gti_det.dgticant + " & Ajuste
           StrSql = StrSql & " WHERE  gti_det.cgtinro  = " & Nro_Cab
           StrSql = StrSql & " AND    gti_det.thnro    = " & objad!thnro
           objConn.Execute StrSql, , adExecuteNoRecords
' Se agrega al AP lo que falta para llegar al tope, en el primer tipo de hora
' encontrado que esté entre las consideradas

        End If
    End If
End Sub



Public Sub Desglose_ACParcial(Nro_Cab As Long, Desde As Date, Hasta As Date)
Dim StrSql As String
Dim E1 As Integer
Dim E2 As Integer
Dim E3 As Integer
Dim te1 As Integer
Dim te2 As Integer
Dim te3 As Integer
Dim canthoras As Single
Dim l_achpnro As Long
Dim auxi As String
Dim rs As New ADODB.Recordset

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Fecha As Date

Dim Ternro As Integer
Dim TotHorHHMM As String

StrSql = "delete from gti_achparc_estr where achpnro in(" & _
"select achpnro from gti_achparcial where cgtinro = " & Nro_Cab & _
")"
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "delete from gti_achparcial where cgtinro = " & Nro_Cab
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "select ternro from gti_cab where cgtinro = " & Nro_Cab
OpenRecordset StrSql, objRs
objConn.Execute StrSql, , adExecuteNoRecords

Ternro = objRs!Ternro
    
Fecha = Desde
' Recorro el desglose del acum. diario por fecha
Do While Fecha <= Hasta
    
    'Por cada desglose para la fecha
    StrSql = "SELECT thnro,achdcanthoras,achdnro FROM gti_achdiario where" & _
    " ternro = " & Ternro & " AND achdfecha = " & ConvFecha(Fecha)
    OpenRecordset StrSql, objRs
    
    Do While Not objRs.EOF
        If Not objRs.EOF Then
            StrSql = "SELECT tenro,estrnro FROM gti_achdiario_estr WHERE achdnro = " & objRs!achdnro
            OpenRecordset StrSql, rs
            E1 = rs!estrnro
            te1 = rs!Tenro
            rs.MoveNext
            E2 = rs!estrnro
            te2 = rs!Tenro
            rs.MoveNext
            E3 = rs!estrnro
            te3 = rs!Tenro
            rs.Close
        End If
        
        ' Busco en la tabla de desgloce de acumulado parcial uno para el empleado
        StrSql = "SELECT achpnro,achpcanthoras FROM gti_achparcial" & _
        " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objRs!thnro & _
        " AND EXISTS (SELECT achpnro FROM gti_achparc_estr WHERE " & _
        " gti_achparc_estr.achpnro = gti_achparcial.achpnro AND" & _
        " estrnro = " & E1 & ")" & _
        " AND EXISTS (SELECT achpnro FROM gti_achparc_estr WHERE " & _
        " gti_achparc_estr.achpnro = gti_achparcial.achpnro AND" & _
        " estrnro = " & E2 & ")" & _
        " AND EXISTS (SELECT achpnro FROM gti_achparc_estr WHERE " & _
        " gti_achparc_estr.achpnro = gti_achparcial.achpnro AND" & _
        " estrnro = " & E3 & ")"
        OpenRecordset StrSql, rs
        
        'Si no existe creo el registro en el desglose del acumulado parcial
        'y uno por cada estructura en el desgloce de
        'acumulado por estructuras
        If rs.EOF Then
            TotHorHHMM = CHorasSF(objRs!achdcanthoras, 60)
        
            StrSql = "INSERT INTO gti_achparcial(horas,achpcanthoras,cgtinro,thnro)" & _
            " VALUES(" & TotHorHHMM & "," & objRs!achdcanthoras & "," & Nro_Cab & "," & _
            objRs!thnro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "select achpnro as next_id from gti_achparcial " & _
            " order by achpnro desc"
            OpenRecordset StrSql, rs
            l_achpnro = rs("next_id")
            
            StrSql = "INSERT INTO gti_achparc_estr(achpnro,tenro,estrnro)" & _
            " VALUES(" & l_achpnro & "," & te1 & "," & E1 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = "INSERT INTO gti_achparc_estr(achpnro,tenro,estrnro)" & _
            " VALUES(" & l_achpnro & "," & te2 & "," & E2 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = "INSERT INTO gti_achparc_estr(achpnro,tenro,estrnro)" & _
            " VALUES(" & l_achpnro & "," & te3 & "," & E3 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
        Else
        'Si existe le sumo
            TotHorHHMM = CHorasSF(rs!achpcanthoras + objRs!achdcanthoras, 60)
        
            StrSql = "UPDATE gti_achparcial SET horas = " & TotHorHHMM & ",achpcanthoras = " & _
            rs!achpcanthoras + objRs!achdcanthoras & _
            " WHERE cgtinro = " & Nro_Cab & _
            " AND achpnro = " & rs!achpnro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        objRs.MoveNext
    Loop
    Fecha = DateAdd("d", 1, Fecha)
Loop
End Sub



Function EsBisiesto(anio As Integer) As Boolean
If (anio Mod 4) = 0 Then
    If (((anio Mod 100) <> 0) And ((anio Mod 400) = 0)) Or _
        (((anio Mod 100) = 0) And ((anio Mod 400) = 0)) Or _
        (((anio Mod 100) <> 0) And ((anio Mod 400) <> 0)) Then
           EsBisiesto = True
       Else
           EsBisiesto = False
    End If
 Else
    EsBisiesto = False
End If

End Function


Private Sub Primeras_Horas_Extras(ByVal Nro_Cab As Long, Desde As Date, Hasta As Date, Nro_tpr As Long, ByVal Nro_Pro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Topea las primeras X horas extras.
' Autor      : FGZ
' Fecha      : 11/01/2007
' Ultima Mod.: 25/01/2007 - FGZ
' Descripcion: Problemas cuando la estructura no esta activa en la totalidad del proceso y no esta abierta
' ---------------------------------------------------------------------------------------------
Dim Horas_Acumuladas As Single
Dim I As Integer
Dim Horas(1 To 5000)  As THoraExced
Dim Opcion As Long
Dim Limite As Single
Dim ListaTH As String
Dim Ternro As Long
Dim Legajo As Long
Dim Aux_hasta As Date

Dim rs As New ADODB.Recordset
Dim rs_Exced As New ADODB.Recordset
Dim rs_gti_det As New ADODB.Recordset
Dim TotHorHHMM As String

    Ternro = Empleado.Ternro
    Legajo = Empleado.Legajo

   Limite = 0
   Horas_Acumuladas = 0
   
    Opcion = st_Opcion
    ListaTH = st_ListaTH
    Limite = st_TamañoVentana
   
    If depurar Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipos de Horas:" & ListaTH
        Flog.writeline Espacios(Tabulador * 4) & "Límite        :" & Limite
        Flog.writeline
    End If


    'Busco todos los tipos de horas a topear y seteo limites y tipos de horas excedentes
    StrSql = " SELECT distinct thnro FROM gti_det "
    StrSql = StrSql & " WHERE thnro IN (" & ListaTH & ")"
    StrSql = StrSql & " ORDER BY thnro "
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        Aux_hasta = Hasta
        StrSql = " SELECT thestrexcmen,htethasta FROM tiphora_estr "
        StrSql = StrSql & " INNER JOIN his_estructura ON tiphora_estr.estrnro = his_estructura.estrnro AND "
        StrSql = StrSql & " his_estructura.tenro = tiphora_estr.tenro AND his_estructura.ternro = " & Empleado.Ternro
        'StrSql = StrSql & " WHERE thnro = " & rs!thnro & " AND his_estructura.htethasta IS NULL "
        'FGZ - 25/01/2007 - cambié la condicion porque puede que no esté activa pero si que algun dia entre dentro de las fechas del proceso
        StrSql = StrSql & " WHERE thnro = " & rs!thnro
        StrSql = StrSql & " AND his_estructura.htetdesde <=" & ConvFecha(Desde) & " AND (his_estructura.htethasta >=" & ConvFecha(Desde) & " OR his_estructura.htethasta IS NULL) "
        OpenRecordset StrSql, rs_Exced
        If Not rs_Exced.EOF Then
             If Not EsNulo(rs_Exced!thestrexcmen) Then
                Horas(rs!thnro).Thnro_Exced = rs_Exced!thestrexcmen
                Horas(rs!thnro).Cant = 0
                Horas(rs!thnro).Cant_Exced = 0
             End If
             If Not EsNulo(rs_Exced!htethasta) Then
                Aux_hasta = rs_Exced!htethasta
             End If
        End If
        rs.MoveNext
    Loop
   
   
   
    'Busco las horas generadas y controlo las primeras "limite" horas generadas
    '   a partir de ahí las horas van como excedentes.
    StrSql = "SELECT gti_acumdiario.thnro, gti_acumdiario.adcanthoras FROM gti_cab "
    StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = gti_cab.ternro "
    StrSql = StrSql & " INNER JOIN gti_tpro_th ON gti_tpro_th.thnro = gti_acumdiario.thnro "
    StrSql = StrSql & " WHERE (gti_cab.cgtinro =" & Nro_Cab & ") and (gti_cab.gpanro = " & Nro_Pro & ") "
    'StrSql = StrSql & " AND (" & ConvFecha(Desde) & " <= gti_acumdiario.adfecha  AND gti_acumdiario.adfecha <= " & ConvFecha(Hasta) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Desde) & " <= gti_acumdiario.adfecha  AND gti_acumdiario.adfecha <= " & ConvFecha(Aux_hasta) & ")"
    StrSql = StrSql & " AND (gti_tpro_th.gtprocnro = " & Nro_tpr & ")"
    StrSql = StrSql & " AND (gti_acumdiario.ternro = " & Ternro & ")"
    StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & ListaTH & ")"
    StrSql = StrSql & " ORDER BY gti_acumdiario.adfecha "
    OpenRecordset StrSql, rs
    Do While Not rs.EOF 'And Horas_Acumuladas < Limite
        If (Horas_Acumuladas + rs!adcanthoras) <= Limite Then
            Horas_Acumuladas = Horas_Acumuladas + rs!adcanthoras
            Horas(rs!thnro).Cant = Horas(rs!thnro).Cant + rs!adcanthoras
        Else
            If Horas_Acumuladas >= Limite Then  'son todo excedente
                Horas(rs!thnro).Cant_Exced = Horas(rs!thnro).Cant_Exced + rs!adcanthoras
            Else
                Horas(rs!thnro).Cant_Exced = Horas_Acumuladas + rs!adcanthoras - Limite
                Horas(rs!thnro).Cant = Horas(rs!thnro).Cant + (Limite - Horas_Acumuladas)
                Horas_Acumuladas = Limite
            End If
        End If
        rs.MoveNext
    Loop
    
    For I = 1 To UBound(Horas())
        'If horas(I).Cant > 0 Then
        'FGZ - 25/01/2007 - cambié la condicion del if
        If Horas(I).Cant_Exced > 0 And Horas(I).Thnro_Exced <> 0 Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * 3) & "Actualizando Tope tipo de hora " & I & " ..."
            End If

            'FGZ - 06/05/2010 ------------------------------------------------------------------------
            'Horas originales
            If Horas(I).Cant > 0 Then
                TotHorHHMM = CHoras(Horas(I).Cant, 60)
                
                StrSql = " SELECT thnro, dgticant FROM gti_det "
                StrSql = StrSql & " WHERE thnro = " & I & " AND cgtinro = " & Nro_Cab
                OpenRecordset StrSql, rs_gti_det
                If Not rs_gti_det.EOF Then
                    StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM & ",dgticant = " & Horas(I).Cant
                    StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & I
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "INSERT INTO gti_det(cgtinro,thnro,horas,dgticant)"
                    StrSql = StrSql & " VALUES (" & Nro_Cab & "," & I & "," & TotHorHHMM & "," & Horas(I).Cant & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                StrSql = " SELECT thnro, dgticant FROM gti_det "
                StrSql = StrSql & " WHERE thnro = " & I & " AND cgtinro = " & Nro_Cab
                OpenRecordset StrSql, rs_gti_det
                If Not rs_gti_det.EOF Then
                    StrSql = "DELETE FROM gti_det WHERE cgtinro = " & Nro_Cab & " AND thnro = " & I
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            
            'FGZ - 06/05/2010 ------------------------------------------------------------------------
            'Horas excedentes
            If Horas(I).Cant_Exced > 0 Then
                TotHorHHMM = CHoras(Horas(I).Cant_Exced, 60)
                
                StrSql = " SELECT thnro, dgticant FROM gti_det "
                StrSql = StrSql & " WHERE thnro = " & Horas(I).Thnro_Exced & " AND cgtinro = " & Nro_Cab
                OpenRecordset StrSql, rs_gti_det
                If Not rs_gti_det.EOF Then
                    StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM & ",dgticant = dgticant + " & Horas(I).Cant_Exced
                    StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & Horas(I).Thnro_Exced
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "INSERT INTO gti_det(cgtinro,thnro,horas,dgticant)"
                    StrSql = StrSql & " VALUES (" & Nro_Cab & "," & Horas(I).Thnro_Exced & "," & TotHorHHMM & "," & Horas(I).Cant_Exced & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            If depurar Then
                Flog.writeline Espacios(Tabulador * 3) & "Tope actualizado"
            End If
        End If
    Next I
    
    'Cierro y libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_gti_det.State = adStateOpen Then rs_gti_det.Close
    Set rs_gti_det = Nothing
    If rs_Exced.State = adStateOpen Then rs_Exced.Close
    Set rs_Exced = Nothing
End Sub

Public Sub Inicializar_Globales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que Carga los array globales.
' Autor      : FGZ
' Fecha      : 17/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    'Politicas de alcance global
    Call Cargar_PoliticasGlobales
    Call Cargar_DetallePoliticas
    
    'FGZ - 15/06/2011
    Call ParametrosGlobales
    
End Sub


Public Sub bus_EscalaAjuste(ByVal NroGrilla As Long, ByVal Cordenada As Long, ByRef HsValida As Double, ByRef HsDescuento As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca en escala los parametros para realizar ajuste.
' Autor      : FGZ
' Fecha      : 18/01/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer

Dim J As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Long
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

'Dim Aux_Dias_trab As Double

Dim ValorCoord As Single
Dim Encontro As Boolean
Dim Encontro2 As Boolean

'Dim VersionBaseAntig As Integer
'Dim Habiles As Integer
'Dim ExcluyeFeriados As Boolean
Dim rs As New ADODB.Recordset

   
    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla
    If rs_cabgrilla.EOF Then
        'La escala de Ajustes no esta configurada
        Flog.writeline "La escala de Ajustes no esta configurada. Escala " & NroGrilla
        Exit Sub
    End If
    
    'El tipo Base variable(15) se considera para que trabaje igual que antiguedad
    TipoBase = 15
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop


    For J = 1 To rs_cabgrilla!cgrdimension
        Select Case J
        Case ant:
            Parametros(J) = Cordenada
        Case Else:
            Select Case J
            Case 1:
                'Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                'Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                'Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                'Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                'Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            Parametros(J) = valor
        End Select
    Next J

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For J = 1 To rs_cabgrilla!cgrdimension
        If J <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & J & "= " & Parametros(J)
        End If
    Next J
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro And Not Encontro2
        Select Case ant
        Case 1:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    If rs_valgrilla!vgrorden = 1 Then
                        HsValida = rs_valgrilla!vgrvalor
                        Encontro = True
                    End If
                    If rs_valgrilla!vgrorden = 2 Then
                        HsDescuento = rs_valgrilla!vgrvalor
                        Encontro2 = True
                    End If
                 End If
            End If
        Case 2:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    If rs_valgrilla!vgrorden = 1 Then
                        HsValida = rs_valgrilla!vgrvalor
                        Encontro = True
                    End If
                    If rs_valgrilla!vgrorden = 2 Then
                        HsDescuento = rs_valgrilla!vgrvalor
                        Encontro2 = True
                    End If
                 End If
            End If
        Case 3:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    If rs_valgrilla!vgrorden = 1 Then
                        HsValida = rs_valgrilla!vgrvalor
                        Encontro = True
                    End If
                    If rs_valgrilla!vgrorden = 2 Then
                        HsDescuento = rs_valgrilla!vgrvalor
                        Encontro2 = True
                    End If
                 End If
            End If
        Case 4:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    If rs_valgrilla!vgrorden = 1 Then
                        HsValida = rs_valgrilla!vgrvalor
                        Encontro = True
                    End If
                    If rs_valgrilla!vgrorden = 2 Then
                        HsDescuento = rs_valgrilla!vgrvalor
                        Encontro2 = True
                    End If
                 End If
            End If
        Case 5:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    If rs_valgrilla!vgrorden = 1 Then
                        HsValida = rs_valgrilla!vgrvalor
                        Encontro = True
                    End If
                    If rs_valgrilla!vgrorden = 2 Then
                        HsDescuento = rs_valgrilla!vgrvalor
                        Encontro2 = True
                    End If
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Private Sub Redondeo_Horas(ByVal Nro_Cab As Long, ByVal Ternro As Long, ByVal Legajo As Long, ByVal Nro_Turno As Long, ByVal Hasta As Date)
'--------------------------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------------------------
Dim TotHorHHMM As String
Dim Hora_a_Red As String
Dim TotHor As Single
   
Dim rs_Redondeo As New ADODB.Recordset
Dim objRs As New ADODB.Recordset
   
    If Not EsNulo(TipoRedondeo) Then
        StrSql = "SELECT * FROM tipredondeo WHERE trdnro =" & TipoRedondeo
        OpenRecordset StrSql, rs_Redondeo
        If rs_Redondeo.EOF Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Redondeo"
            End If
            ListaTHAP = "0"
        Else
            If rs_Redondeo!trdtipo <> 1 Then
                If depurar Then
                    Flog.writeln "No es Redondeo (Fraccionamiento)"
                End If
                ListaTHAP = "0"
            End If
        End If
    Else
        ListaTHAP = "0"
    End If
 
    

   StrSql = " SELECT thnro, dgticant,horas FROM gti_det "
   StrSql = StrSql & " WHERE (thnro IN (" & ListaTHAP & ")) AND cgtinro = " & Nro_Cab
   StrSql = StrSql & " ORDER BY thnro "
   OpenRecordset StrSql, objRs
   Do While Not objRs.EOF
        'redondeo -------------
        TotHor = objRs!dgticant

        'Redondeo del total de horas
        objFechasHoras.Convertir_A_Hora objRs!dgticant * 60, Hora_a_Red

        Call Redondeo_enHorasMinutos(objRs!Horas, rs_Redondeo!trdnro, 60, TotHorHHMM)
        TotHorHHMM = "'" & TotHorHHMM & "'"

        If depurar Then
            Flog.writeline Espacios(Tabulador * 3) & " en horas y minutos (HHMM) " & Hora_a_Red
        End If
        objFechasHoras.Redondeo_Horas_Tipo Hora_a_Red, rs_Redondeo!trdnro, TotHor
        If depurar Then
            Flog.writeline Espacios(Tabulador * 3) & " --- Tipo de redondeo " & rs_Redondeo!trdnro
            Flog.writeline Espacios(Tabulador * 3) & " --- luego del redondeo " & TotHor
        End If
        'redondeo -------------
   
        If TotHor <> 0 Then
                StrSql = "UPDATE gti_det SET horas = " & TotHorHHMM & ",dgticant = " & TotHor
                StrSql = StrSql & " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objRs!thnro
                objConn.Execute StrSql, , adExecuteNoRecords
        Else
                StrSql = "DELETE FROM gti_det WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objRs!thnro
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
   
        objRs.MoveNext
    Loop
    
'Cierro y libero
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing

If rs_Redondeo.State = adStateOpen Then rs_Redondeo.Close
Set rs_Redondeo = Nothing

End Sub

