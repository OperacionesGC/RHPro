Attribute VB_Name = "MdlRepGanancias"
Option Explicit

'Const Version = 1.01 'Correccion en la busqueda de las ddjj del item 13 - cuota medica
'Const FechaVersion = "15/12/2005"

'Const Version = 1.02 'Controla si el empleado tiene alguna DDJJ de una empresa anterior
'Const FechaVersion = "22/12/2005"

'Const Version = 1.03 'Se agrego la opcion de rectifica
'Const FechaVersion = "26/12/2005"

'Const Version = 1.04 'Se agrego el manejador de errores y que solo genere para aquellos que tienen retenciones
'Const FechaVersion = "29/12/2005"

'Const Version = 1.05 'Se agrego el manejador de errores Global
'Const FechaVersion = "02/01/2006"

'Const Version = 1.06 'Habia un error en una UPDATE en : ITEM 13 Obra Social Privada
'Const FechaVersion = "03/01/2006"

'Const Version = 1.07 'Se agregó la opcion de generar solo para los que tienen retenciones
'Const FechaVersion = "06/01/2006"

'Const Version = 1.08 '
'Const FechaVersion = "09/01/2006"

'Const Version = 1.09 'considerar que puede venir una rectificacion en enero del proximo año correspondientes al año actual
'Const FechaVersion = "09/01/2006"

'Const Version = 1.1  'la rectificacion de ganancias es identificable por el modelo.
'Const FechaVersion = "25/01/2006"

'Const Version = 1.11  'la rectificacion de ganancias es identificable por el modelo.
'Const FechaVersion = "26/01/2006"

'Const Version = 1.12  'Problema cuando tiro el reporte para todos los legajos. Ojo que tambien se modificó el asp que genera el proceso
'Const FechaVersion = "15/02/2006"

'Const Version = 1.13  'La determinacion si debe rectificar sabe del modelo de liquidación va dentro del ciclo
'Const FechaVersion = "22/02/2006"

'Const Version = 1.14  'Suegan21 con mas log en Rubro 14
'Const FechaVersion = "27/02/2006"

'Const Version = 1.15  'Suegan21 - Otras Deducciones, cambié el caso 8 y 9 porque el liq mete el honorariomedico en el 9 y me sale duplicado
'Const FechaVersion = "01/03/2006"

'Const Version = 1.16  'Suegan21 - COUTAS MEDICO ASISTENCIALES
'Const FechaVersion = "02/03/2006"

'Const Version = 1.17  'Suegan21 - COUTAS MEDICO ASISTENCIALES, se agregó el Items_OLD_LIQ
'Const FechaVersion = "06/03/2006"

'Const Version = 1.18  'Suegan21 - COUTAS MEDICO ASISTENCIALES, cuando busca las ddjj las busca ordenadas por monto descendente
'Const FechaVersion = "06/03/2006"

'Const Version = 1.19  'Suegan21 - Primas seguro y Donaciones no topean
'Const FechaVersion = "23/03/2006"

'Const Version = 1.21   'Suegan21 - Primas seguro y Donaciones no topean. ABS
'Const FechaVersion = "27/03/2006"

'Const Version = 1.22   'Suegan21 - Donaciones, problemas con los diferidos y los topes y la mar en coche
                                  'Se reutiliza el campo cuit_entidad6 y 7 para mostrar el importe de la DDJJ
                                  'Esta modificacion trae asociado una mod en el asp que muestra
'Const FechaVersion = "28/03/2006"

'Const Version = 1.23  'correccion fecha generacion
'Const FechaVersion = "26/05/2006"

'Const Version = 1.24  'Unificacion de Fuentes entre BB - BA
'Const FechaVersion = "01/06/2006"

'Const Version = 1.25  'Suegan20 - La variable emprEstrnro se paso de integer a long
'Const FechaVersion = "01/06/2006"

'Const Version = 1.26  'Suegan21 - Donaciones, estaba poniendo el importe en el cuit y luego el asp mostraba el importe
''                       esta relacionado con ...
''Const Version = 1.22   'Suegan21 - Donaciones, problemas con los diferidos y los topes y la mar en coche
'                                  'Se reutiliza el campo cuit_entidad6 y 7 para mostrar el importe de la DDJJ
'                                  'Esta modificacion trae asociado una mod en el asp que muestra
'Const FechaVersion = "07/12/2006"   'FGZ

'Const Version = 1.27  'Suegan20 - 'faltaban los joins con las tablas proceso y tipoproc cuando se generaba para un solo legajo
'Const FechaVersion = "02/01/2007"   'FGZ

'Const Version = 1.28  'Suegan20 - 'faltaban los joins con las tablas proceso y tipoproc cuando se generaba para un solo legajo
'Const FechaVersion = "16/03/2007"   'FGZ
''                                   Sub suegan21. Problemas con los items 6 y 13 para el rubro 11
''
''OBS:                               1. El concepto de OS debe sumar al item 6 de ganancias.
''                                   2. No se deben cargar DDJJ para el item 6 en 0 para que salga en el reverso del F649


'Const Version = "1.29b"
'Const FechaVersion = "26/03/2007"   'FGZ
'                                   Sub suegan20. Agregado de logs y seteo de Por_Deduccion en 100 por default
'                                                 Modificacion para que no tome + de 1 proceso cuando TODOS TIENEN LA MISMA FECHA DE PAGO!!!

'Const Version = "1.30"
'Const FechaVersion = "25/02/2008"   'MB
'                                   Sub suegan21. Las primas de seguro estaba multip por cada desliq, se puso en monto_entidad3 el valor de item_tope(8)
'                                                 y en total_entidad2 el mismo valor. Tocado para Marsh

'Const Version = "1.31"
'Const FechaVersion = "30/12/2008"   'Martin Ferraro
'                                    rs_Traza_gan!msr + Items_TOPE(50)
'                                    rs_Traza_gan!otras - Items_TOPE(50)
'                                    Entra a escala_ded con (Gan_imponible + Deducciones - Items_TOPE(50))

'Const Version = "1.32"
'Const FechaVersion = "24/02/2009"   'Martin Ferraro
'                                    Se agrego el parametro lugar
'                                    Valor absoluto a Cuit_entidad8, 9 y 10

'Const Version = "1.33"
'Const FechaVersion = "27/02/2009"   'Martin Ferraro - En la parte del dorso donde busca los items en el rubro 19
'                                                     cuando case 8 insertaba en 9 entonces si habia una sola deduccion salia en el segundo reglon
'                                                     Ademas se agrego que tenga en cuenta el item 7 de sindicato

'Const Version = "1.34"
'Const FechaVersion = "19/05/2009"   'Martin Ferraro - Faltaba la validacion de fechas para desmen

'Const Version = "1.35"
'Const FechaVersion = "31/07/2009"   'Martin Ferraro - Encriptacion de string connection

'Const Version = "1.36"
'Const FechaVersion = "13/01/2010"   'Martin Ferraro - Tomaba mal un tope

'Const Version = "1.37"
'Const FechaVersion = "12/04/2010"   'Martin Ferraro - Cambio en el calculo item 13

'Const Version = "1.38"
'Const FechaVersion = "15/09/2010"   'Martin Ferraro - Se reformulo Suegan21
'                                    'Se agrego el cambio legal RG 2866
'

'Const Version = "1.39"
'Const FechaVersion = "17/05/2011"   'FGZ - Solo busca liquidaciones de ganancia el año de la fecha de corte
                                    
'Const Version = "1.40"
'Const FechaVersion = "28/06/2011"   'FGZ - Se agregan los items 22 y 24 al rubro 15
                                    
                                    
'Const Version = "1.41"
'Const FechaVersion = "19/03/2012"   'Sebastian Stremel - correccion sindicato

'Const Version = "1.42"
'Const FechaVersion = "14/09/2012"   'Dimatz Rafael - 16847 - Se corrigio para que muestre el domicilio principal

'Const Version = "1.43"
'Const FechaVersion = "02/10/2012"   'Dimatz Rafael - 16911 - Se Compilo para OSDOP

'Const Version = "1.44"
'Const FechaVersion = "15/01/2013"   'Sebastian Stremel - 16911 - se agrega item 23 - CAS-18070 - GC - Error F649

'Const Version = "1.45"
'Const FechaVersion = "13/05/2013"   'Lisandro Moro - CAS-19228 - Horwath Argentina - Bug Formulario 649. [Entrega 2]
                                    'Correccion al rubro 15 si no posee sindicato...
                                    
'Const Version = "1.46"
'Const FechaVersion = "27/06/2013"   'Mauricio Zwenger - CAS-19228 - Horwath Argentina - Bug Formulario 649. [Entrega 3]
'                                    'se agrego el item 56 al rubro 15

'Const Version = "1.47"
'Const FechaVersion = "02/10/2013"   'FGZ - CAS-21616 - NGA - DDEE en F649 luego de cambio de ganancias
                                    'se sacó el item 56 del rubro 15, solo debe informarse en el anverso del F649 en el rubro 9b.


'Const Version = "1.48"
'Const FechaVersion = "05/12/2013"   'FGZ - CAS-22730 - H&A - CAMBIO LEGAL EN F649
                                    'Se ajustó como mostraba las deducciones especiales de acuerdo a la modificación realizada en el liquidador V5.52,
                                    'para que muestre correctamente el valor del ítem 16 Deducción especial
                                    
'Const Version = "1.49"
'Const FechaVersion = "27/05/2013"   'Miriam Ruiz - CAS-25433 - NGA - Mejora en filtro de F649 (original vs rectificativa)
                                    'Se agregó a la tabla si es original o rectificativa,


'Const Version = "1.50"
'Const FechaVersion = "19/08/2014"   'Fernandez, Matias - CAS-26543 -SANTANA - Error en Rubro 11 del F649
                                    'Se agregó a la tabla si es original o rectificativa,

Const Version = "1.51"
Const FechaVersion = "08/10/2014"   'Carmen Quintero - CAS-27438 - SANTANA TEXTIL - BUG EN F649
                                    'Se agregó validacion para el caso cuando, los montos de la DDJJ son menores a lo calculado por el proceso


'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Global IdUser As String
Global Fecha As Date
Global hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String
Dim Items_TOPE(100) As Double
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_OLD_LIQ(100) As Double
Dim Descuentos_Items As Double
Dim Gan_imponible_Items As Double
Dim Ded_a23_Items As Double
Dim Deducciones_Items As Double
Global SoloConRetenciones As Boolean
'FGZ - 02/03/2006
Dim Rectifica As Boolean




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'         If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Nombre_Arch = PathFLog & "Reporte_Ganancias" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline "Inicio                   : " & Format(Now, FormatoInternoFecha)
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main

    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 45 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin:
End Sub


Public Sub Suegan20(ByVal Bpronro As Long, ByVal FechaHasta As Date, ByVal Prorratea As Boolean, ByVal Tope_Min As Single, ByVal Tope_Max As Single, ByVal Anual_Final As Boolean, ByVal Suscribe As String, ByVal Caracter As String, ByVal Fecha_Caracter As String, ByVal Fecha_Devolucion As String, ByVal Dependencia_DGI As String, ByVal Empresa As Long, ByVal Todos_Empleados As Boolean, ByVal Lugar As String, ByVal original As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Ganancias (Formulario 649)
' Autor      : FGZ
' Fecha      : 20/04/2004
' Ult. Mod   : 29/12/2004
' Fecha      :
'Descripcion : Se le agregó un control para que no proces todas las traza_gan de todos los empleados < a la fecha
'               sino solo la ultima
'Ultima Mod  :  FGZ - 02/01/2007 - faltaban los joins con las tablas proceso y tipoproc
'Ultima Mod  :  FGZ - 26/03/2007 - Agregado de logs y seteo de Por_Deduccion en 100 por default
' --------------------------------------------------------------------------------------------
Dim I As Integer

'Auxiliares
Dim v_desde As Date
Dim v_hasta As Date
Dim Saltar_empleado As Boolean
Dim Dir_calle As String
Dim Dir_cp As String
Dim Dir_num As String
Dim Dir_dpto As String
Dim Dir_piso As String
Dim Dir_localidad As String
Dim Dir_pcia As String
Dim Direccion As String
Dim cuit As String
Dim Cuil As String
Dim Retenciones As Double
Dim Monto_Letras As String
Dim Emp_Nombre As String
Dim Emp_CUIT As String
Dim Mes_Ret As Long
Dim Ano_Ret As Long
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date

Dim AcuRG2866 As Long
Dim ValorAcuRG2866 As Double
Dim DesdeRG2866 As Date
Dim HastaRG2866 As Date

Dim Aux_Empleado As String

Dim rs_Proceso As New ADODB.Recordset
Dim rs_Rep19 As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset
Dim rs_CabdomPrincipal As New ADODB.Recordset
Dim rs_Cabdom As New ADODB.Recordset
Dim rs_Detdom As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_Localidad As New ADODB.Recordset
Dim rs_Provincia As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_Cuit As New ADODB.Recordset
Dim rs_ficharet As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Empresa2 As New ADODB.Recordset
Dim rs_items_tope As New ADODB.Recordset

Dim Ultimo_Empleado As Long
Dim Aux_Fecha_A_Utilizar As Date
Dim emprEstrnro As Long
Dim Auxiliar As String
Dim Gan_imponible As Single
Dim Deducciones As Single
'Dim Rectifica As Boolean

Dim rs_escala_ded As New ADODB.Recordset
Dim Por_Deduccion As Single
Dim EsElPrimero As Boolean
Dim Anio_Anterior As Boolean
Dim rs_Traza_temp As New ADODB.Recordset


Dim YaFueGenerado As Boolean

Dim FechaDesdeAñoActual As Date

Dim Ajuste As Double


'Inicializacion

'Adrián - No buscamos los procesos. Directamente buscamos en traza_gan.

'BUSCO LA FECHA DEL PROCESO DEL ULTIMO PROCESO DE LIQUIDACION CON FECHA DE PAGO MENOR A LA FECHA HASTA DEL REPORTE
'StrSql = "SELECT * FROM proceso WHERE profecpago <= " & ConvFecha(FechaHasta)
'StrSql = StrSql & " ORDER BY profecpago DESC"
'OpenRecordset StrSql, rs_Proceso
'If rs_Proceso.EOF Then
'    Flog.writeline "No se encontró ningun proceso anterior a esa fecha :" & FechaHasta
'    Exit Sub
'End If
'rs_Proceso.MoveLast

On Error GoTo CE

'16/09/2010 - Acumulador Cambio legal Saldo RG 2866
AcuRG2866 = 0
DesdeRG2866 = "01/01/2010"
HastaRG2866 = "31/12/2010"

StrSql = "SELECT confval FROM confrep WHERE repnro = 289 AND confnrocol = 1"
OpenRecordset StrSql, rs_Empresa2
If Not rs_Empresa2.EOF Then
    AcuRG2866 = rs_Empresa2!confval
Else
    Flog.writeline "No se encontró el Acumulador Cambio legal Saldo RG 2866 en la columna 1 de la configuracion del reporte 289"
End If
rs_Empresa2.Close


' Busco la estructura de la empresa
StrSql = " SELECT * FROM empresa WHERE empnro = " & Empresa

OpenRecordset StrSql, rs_Empresa2

If Not rs_Empresa2.EOF Then
  emprEstrnro = rs_Empresa2!Estrnro
Else
  Flog.writeline "No se encontró la empresa"
  Exit Sub
End If

'Busco los registros generados en traza_gan
'StrSql = "SELECT * FROM  traza_gan "
'StrSql = StrSql & " INNER JOIN batch_empleado ON traza_gan.ternro = batch_empleado.ternro"
'StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = batch_empleado.ternro and empresa.tenro = 10 AND empresa.estrnro = " & emprEstrnro
''StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = batch_empleado.ternro and empresa.tenro = 10 "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
'StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro "
'StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
'
''Adrián - Buscamos en traza_gan con traza_gan.fecha_pago <= " & ConvFecha(FechaHasta) y no con pronro.
'StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
''StrSql = StrSql & " WHERE traza_gan.pronro = " & rs_Proceso!pronro
''FGZ - 30/12/2004
'StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
''StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
''FGZ - 09/01/2006 Modifique esto porque si el legajo tiene fecha de baja dentro del periodo ==> no sale
'StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
''StrSql = StrSql & " ((" & ConvFecha(C_Date("01/01/" & Year(FechaHasta))) & " <= htethasta) or (htethasta is null))"
'StrSql = StrSql & " AND batch_empleado.bpronro = " & Bpronro
'StrSql = StrSql & " ORDER BY batch_empleado.ternro, fecha_pago DESC, traza_gan.pronro DESC"
'OpenRecordset StrSql, rs_Traza_gan


'FGZ - 15/02/2006
'FechaDesdeAñoActual = C_Date("01/01/" & Year(FechaHasta), "dd/mm/yyyy")
FechaDesdeAñoActual = CDate("01/01/" & Year(FechaHasta))
If Todos_Empleados Then
    StrSql = "SELECT * FROM  traza_gan "
    StrSql = StrSql & " INNER JOIN empleado ON traza_gan.ternro = empleado.ternro"
    StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = traza_gan.ternro and empresa.tenro = 10 AND empresa.estrnro = " & emprEstrnro
    StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro "
    StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
    StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
    'FGZ - 17/05/2011 -------------------------------------------------------------------
    StrSql = StrSql & " AND fecha_pago >= " & ConvFecha(FechaDesdeAñoActual)
    'FGZ - 17/05/2011 -------------------------------------------------------------------
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
    StrSql = StrSql & " AND empleado.empest = -1"
    'StrSql = StrSql & " ORDER BY empleado.ternro, fecha_pago DESC, traza_gan.pronro DESC"
    StrSql = StrSql & " ORDER BY empleado.empleg, fecha_pago DESC, traza_gan.pronro DESC"
Else
    StrSql = "SELECT * FROM  traza_gan "
    StrSql = StrSql & " INNER JOIN batch_empleado ON traza_gan.ternro = batch_empleado.ternro"
    StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = batch_empleado.ternro and empresa.tenro = 10 AND empresa.estrnro = " & emprEstrnro
    StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro "
    StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
    StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
    'FGZ - 17/05/2011 -------------------------------------------------------------------
    StrSql = StrSql & " AND fecha_pago >= " & ConvFecha(FechaDesdeAñoActual)
    'FGZ - 17/05/2011 -------------------------------------------------------------------
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
    StrSql = StrSql & " AND batch_empleado.bpronro = " & Bpronro
    StrSql = StrSql & " ORDER BY batch_empleado.ternro, fecha_pago DESC, traza_gan.pronro DESC"
End If
Flog.writeline "STRSQL= " & StrSql
OpenRecordset StrSql, rs_Traza_gan


'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Traza_gan.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

If rs_Traza_gan.EOF Then
    Flog.writeline "No hay registros generados para ganancias (TRAZA_GAN) a esa fecha : " & FechaHasta
    
    'FGZ - 09/01/2006 esto es temporal y solo es para ver porque no muestra ciertos legajos
    If Todos_Empleados Then
        StrSql = "SELECT * FROM  traza_gan "
        StrSql = StrSql & " INNER JOIN empleado ON traza_gan.ternro = empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = traza_gan.ternro and empresa.tenro = 10 "
        StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro "
        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
        StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
        'FGZ - 17/05/2011 -------------------------------------------------------------------
        StrSql = StrSql & " AND fecha_pago >= " & ConvFecha(FechaDesdeAñoActual)
        'FGZ - 17/05/2011 -------------------------------------------------------------------
        StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
        StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " AND empleado.empest = -1"
        StrSql = StrSql & " ORDER BY empleado.empleg, fecha_pago DESC, traza_gan.pronro DESC"
    Else
        StrSql = "SELECT * FROM  traza_gan "
        StrSql = StrSql & " INNER JOIN batch_empleado ON traza_gan.ternro = batch_empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = batch_empleado.ternro and empresa.tenro = 10 "
        StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
        'FGZ - 02/01/2007 - faltaban los joins con las tablas proceso y tipoproc
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro "
        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
        'FGZ - 02/01/2007 - faltaban los joins con las tablas proceso y tipoproc
        StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
        'FGZ - 17/05/2011 -------------------------------------------------------------------
        StrSql = StrSql & " AND fecha_pago >= " & ConvFecha(FechaDesdeAñoActual)
        'FGZ - 17/05/2011 -------------------------------------------------------------------
        StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
        StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " AND batch_empleado.bpronro = " & Bpronro
        StrSql = StrSql & " ORDER BY batch_empleado.ternro, fecha_pago DESC, traza_gan.pronro DESC"
    End If
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    Flog.writeline "strsql= " & StrSql
    OpenRecordset StrSql, rs_Traza_gan
    If rs_Traza_gan.EOF Then
        Flog.writeline "Efectivamente, No hay registros generados"
    Else
        Flog.writeline "El problema esta en la empresa"
    End If
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    Set rs_Traza_gan = Nothing
    Exit Sub
Else
    'FGZ - 09/01/2006
    'Cuento la cantidad de legajos a evaluar
    'FGZ - 09/01/2006 esto es temporal
    If Todos_Empleados Then
        StrSql = "SELECT distinct(traza_gan.ternro) FROM  traza_gan "
        StrSql = StrSql & " INNER JOIN empleado ON traza_gan.ternro = empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = traza_gan.ternro and empresa.tenro = 10 AND empresa.estrnro = " & emprEstrnro
        StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
        StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
        StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
        StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " AND empleado.empest = -1"
    Else
        StrSql = "SELECT distinct(traza_gan.ternro) FROM  traza_gan "
        StrSql = StrSql & " INNER JOIN batch_empleado ON traza_gan.ternro = batch_empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = batch_empleado.ternro and empresa.tenro = 10 AND empresa.estrnro = " & emprEstrnro
        StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
        StrSql = StrSql & " WHERE traza_gan.fecha_pago <= " & ConvFecha(FechaHasta)
        StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
        StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " AND batch_empleado.bpronro = " & Bpronro
    End If
    If rs_Traza_temp.State = adStateOpen Then rs_Traza_temp.Close
    OpenRecordset StrSql, rs_Traza_temp
    If rs_Traza_temp.EOF Then
        Flog.writeline "Legajos a evaluar: 0"
    Else
        Flog.writeline "Legajos a evaluar: " & rs_Traza_temp.RecordCount
    End If
    If rs_Traza_temp.State = adStateOpen Then rs_Traza_temp.Close
    Set rs_Traza_temp = Nothing

'    ' FAF - 25-01-06 - La determinacion si debe rectificar sabe del modelo de liquidación
'    Rectifica = CBool(rs_Traza_gan!ajugcias)
'    Flog.writeline "Ajuste de Ganancia según Modelo de Liquidación: " & Rectifica
End If

' Comienzo la transaccion
MyBeginTrans

Flog.writeline
Flog.writeline


Ultimo_Empleado = -1
Do While Not rs_Traza_gan.EOF
    
    
    'FGZ - 29/12/2004
    'Se le agregó este control porque levantaba todas las traza_gan de todos los empleados < a la fecha y en orden descendiente
    ' por lo que el primer registro encontrado para cada legajo es el correcto pero los siguientes son viejos
    ' y si los proceso va a quedar con resultado incorrecto
    EsElPrimero = False
    If Ultimo_Empleado <> rs_Traza_gan!Ternro Then
        'FGZ -26 / 3 / 2007
        YaFueGenerado = False
        'FAF - 25-01-06 - La determinacion si debe rectificar sale del modelo de liquidación
        Rectifica = CBool(rs_Traza_gan!ajugcias)
        Flog.writeline "Ajuste de Ganancia según Modelo de Liquidación: " & Rectifica

'        If rs_Traza_gan!Retenciones = 0 Then
'            If SoloConRetenciones Then
                Aux_Fecha_A_Utilizar = rs_Traza_gan!fecha_pago
                Ultimo_Empleado = rs_Traza_gan!Ternro
                EsElPrimero = True
'            End If
'        Else
'            Aux_Fecha_A_Utilizar = rs_Traza_gan!fecha_pago
'            Ultimo_Empleado = rs_Traza_gan!ternro
'            EsElPrimero = True
'        End If
    End If
    
    'Controlo si hay que generarle el F649, si tiene ganancias
    If rs_Traza_gan!Retenciones = 0 And SoloConRetenciones Then
        Flog.writeline "### El empleado(ternro=" & rs_Traza_gan!Ternro & ") no tiene retenciones generadas, no se le generara el F649."
        EsElPrimero = False
    Else
        EsElPrimero = True
        Flog.writeline "### El empleado(ternro=" & rs_Traza_gan!Ternro & ") tiene retenciones generadas, se le generara el F649. Retencion = " & rs_Traza_gan!Retenciones
    End If
    Flog.writeline "-------------------------------"
    
    'FGZ - 26/03/2007
    'If (rs_Traza_gan!fecha_pago = Aux_Fecha_A_Utilizar) And EsElPrimero Then
    If (rs_Traza_gan!fecha_pago = Aux_Fecha_A_Utilizar) And EsElPrimero And Not YaFueGenerado Then
        YaFueGenerado = True
        'Busco el legajo
        StrSql = "SELECT * FROM empleado"
        StrSql = StrSql & " WHERE ternro =" & rs_Traza_gan!Ternro
        If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
        OpenRecordset StrSql, rs_Empleado
        If Not rs_Empleado.EOF Then
            Aux_Empleado = rs_Empleado!empleg
            Flog.writeline "Empleado: " & rs_Empleado!empleg
            Flog.writeline "Fecha de Pago encontrada: " & rs_Traza_gan!fecha_pago
        Else
            Aux_Empleado = "???"
            Flog.writeline "Datos del empleado no encontrados (tercero " & rs_Traza_gan!Ternro & ")"
        End If
        Flog.writeline ""
        
        Saltar_empleado = False
        'v_desde = ?
        ' esto es idea mia
        v_desde = CDate("01/01/" & Year(FechaHasta))
        v_hasta = FechaHasta
               
               
        Anio_Anterior = False
        StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_Traza_gan!Ternro
        StrSql = StrSql & " AND fases.altfec <=" & ConvFecha(CDate("31/12/" & Year(FechaHasta)))
        StrSql = StrSql & " ORDER BY fases.altfec "
        OpenRecordset StrSql, rs_Fases
        If Not rs_Fases.EOF Then rs_Fases.MoveLast
        If Not rs_Fases.EOF Then
            If IsNull(rs_Fases!bajfec) Then
                v_desde = CDate("01/01/" & Year(FechaHasta))
            Else
                If rs_Fases!bajfec < CDate("01/01/" & Year(FechaHasta)) Then
                
                    If Anual_Final Then
                        Saltar_empleado = True
                        Flog.writeline "Empleado no considerado. fecha de baja " & rs_Fases!bajfec & " en año anterior"
                        Flog.writeline "-------------------------------"
                        Flog.writeline
                    Else
                      Flog.writeline "Empleado con fecha de baja " & rs_Fases!bajfec & " en año anterior"
                      
                      Anio_Anterior = True
                      v_desde = CDate("01/01/" & Year(FechaHasta))
                      v_hasta = CDate("31/12/" & Year(FechaHasta)) 'rs_Fases!bajfec 'FechaHasta
                    End If
                Else
                    v_desde = CDate("01/01/" & Year(FechaHasta))
                End If
            End If
         End If
        
        If (rs_Traza_gan!nomsr + rs_Traza_gan!msr < Tope_Min) Or (rs_Traza_gan!nomsr + rs_Traza_gan!msr > Tope_Max) Then
            Saltar_empleado = True
            Flog.writeline "Empleado no considerado. retencion menor que el minimo considerado " & Tope_Min
            Flog.writeline "-------------------------------"
            Flog.writeline
        End If
        
        If Not Saltar_empleado Then
'            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_Traza_gan!ternro
'            StrSql = StrSql & " AND fases.altfec <=" & ConvFecha(CDate("31/12/" & Year(FechaHasta)))
'            StrSql = StrSql & " ORDER BY altfec "
'            OpenRecordset StrSql, rs_Fases
'            If Not rs_Fases.EOF Then rs_Fases.MoveLast
'            If Not rs_Fases.EOF Then
'                If IsNull(rs_Fases!bajfec) Then
'                    If IsNull(v_desde) Then
'                        v_desde = CDate("01/01/" & Year(FechaHasta))
'                    End If
'                Else
'                    If rs_Fases!bajfec < CDate("31/12/" & Year(FechaHasta)) Then
'                        v_desde = IIf(IsNull(v_desde), rs_Fases!altfec, v_desde)
'                    Else
'                        v_Hasta = FechaHasta
'                        v_desde = IIf(IsNull(v_desde), rs_Fases!altfec, v_desde)
'                    End If
'                End If
'            End If
                
            Anio_Anterior = False
            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_Traza_gan!Ternro
            StrSql = StrSql & " AND fases.altfec <=" & ConvFecha(CDate("31/12/" & Year(FechaHasta)))
            StrSql = StrSql & " ORDER BY fases.altfec "
            OpenRecordset StrSql, rs_Fases
            If Not rs_Fases.EOF Then rs_Fases.MoveLast
            If Not rs_Fases.EOF Then
                If IsNull(rs_Fases!bajfec) Then
                    v_desde = CDate("01/01/" & Year(FechaHasta))
                Else
                    If rs_Fases!bajfec < CDate("01/01/" & Year(FechaHasta)) Then
                    
                        If Anual_Final Then
                            Saltar_empleado = True
                            Flog.writeline "Empleado no considerado. fecha de baja " & rs_Fases!bajfec & " en año anterior"
                            Flog.writeline "-------------------------------"
                            Flog.writeline
                        Else
                          Flog.writeline "Empleado con fecha de baja " & rs_Fases!bajfec & " en año anterior"
                          
                          Anio_Anterior = True
                          v_desde = CDate("01/01/" & Year(FechaHasta))
                          v_hasta = CDate("31/12/" & Year(FechaHasta)) 'rs_Fases!bajfec 'FechaHasta
                        End If
                    Else
                        v_desde = CDate("01/01/" & Year(FechaHasta))
                    End If
                End If
             End If
                
                
            
            Mes_Ret = IIf(Prorratea, 12, Month(v_hasta))
            Ano_Ret = Year(v_hasta)
           
            fin_mes_ret = CDate("01/" & ((Mes_Ret + 1) Mod 13) + Int(Fix(Mes_Ret / 12)) & "/" & Ano_Ret + Int(Fix(Mes_Ret / 12))) - 1
            ini_anyo_ret = CDate("01/01/" & Ano_Ret)
            v_desde = IIf(IsNull(v_desde), ini_anyo_ret, v_desde)
        
            Flog.writeline
            Flog.writeline "Fase desde: " & v_desde & " hasta: " & v_hasta
            Flog.writeline
            
            'direccion
            'Cargo con valores nulos por si no encuentra
            Dir_calle = "NO DISPONIBLE"
            Dir_cp = "N.D."
            Dir_num = "N.D."
            Dir_dpto = "N.D."
            Dir_piso = "N.D."
            Dir_localidad = "NO DISPONIBLE"
            Dir_pcia = "NO DISPONIBLE"
            
            Direccion = Lugar
            
            Flog.writeline " Busco la direccion del empleado" & Aux_Empleado
            
            StrSql = "SELECT domdefault FROM cabdom "
            StrSql = StrSql & "WHERE cabdom.ternro = " & rs_Traza_gan!Ternro
            StrSql = StrSql & " AND cabdom.domdefault = -1"
            OpenRecordset StrSql, rs_CabdomPrincipal
            Flog.writeline "Verifica si tiene Domicilio Principal" & StrSql
            
            If Not rs_CabdomPrincipal.EOF Then
                StrSql = "SELECT * FROM ter_tip "
                StrSql = StrSql & " INNER JOIN cabdom ON ter_tip.tipnro = cabdom.tipnro "
                StrSql = StrSql & " WHERE ter_tip.ternro = " & rs_Traza_gan!Ternro & " AND ter_tip.tipnro = 1"
                StrSql = StrSql & " AND cabdom.ternro = " & rs_Traza_gan!Ternro
                StrSql = StrSql & " AND cabdom.domdefault = -1"
                Flog.writeline "Tiene Domicilio Principal" & StrSql
            Else
                StrSql = "SELECT * FROM ter_tip "
                StrSql = StrSql & " INNER JOIN cabdom ON ter_tip.tipnro = cabdom.tipnro "
                StrSql = StrSql & " WHERE ter_tip.ternro = " & rs_Traza_gan!Ternro & " AND ter_tip.tipnro = 1"
                StrSql = StrSql & " AND cabdom.ternro = " & rs_Traza_gan!Ternro
                Flog.writeline "No tiene Domicilio Principal" & StrSql
            End If
                              
            rs_CabdomPrincipal.Close
                              
            OpenRecordset StrSql, rs_Cabdom
            If Not rs_Cabdom.EOF Then
                StrSql = "SELECT * FROM detdom WHERE domnro =" & rs_Cabdom!domnro
                OpenRecordset StrSql, rs_Detdom
                If Not rs_Detdom.EOF Then
                    Dir_calle = rs_Detdom!calle
                    If Not IsNull(rs_Detdom!codigopostal) Then
                        Dir_cp = CStr(rs_Detdom!codigopostal)
                    End If
                    
                    If Not IsNull(rs_Detdom!nro) Then
                        Dir_num = CStr(rs_Detdom!nro)
                    End If
                    If Not IsNull(rs_Detdom!oficdepto) Then
                        Dir_dpto = rs_Detdom!oficdepto
                    End If
                    If Not IsNull(rs_Detdom!piso) Then
                        Dir_piso = CStr(rs_Detdom!piso)
                    End If
                        
                    'Direccion = Dir_calle & " " & Dir_num & " (" & Dir_cp & ")"
                    'Direccion = rs_Detdom!calle & " " & CStr(rs_Detdom!nro) & "  (" & CStr(rs_Detdom!codigopostal) & ")"
                    
                    StrSql = "SELECT * FROM localidad WHERE locnro =" & rs_Detdom!locnro
                    OpenRecordset StrSql, rs_Localidad
                    If Not rs_Localidad.EOF Then
                        Dir_localidad = rs_Localidad!locdesc
                        'Direccion = Direccion & "  " & rs_Localidad!locdesc
                    Else
                        Flog.writeline " No se encontró la localidad del empleado"
                    End If
                    
                    StrSql = "SELECT * FROM provincia WHERE provnro =" & rs_Detdom!provnro
                    OpenRecordset StrSql, rs_Provincia
                    If Not rs_Provincia.EOF Then
                        Dir_pcia = rs_Provincia!provdesc
                    Else
                        Flog.writeline " No se encontró la provincia del empleado"
                    End If
                Else
                    'Direccion = " "
                    Flog.writeline " No se encontró la direccion del empleado"
                End If
            End If

            'Adrián -  Limitamos la longitud del string.
            'Direccion = Left(Direccion, 40)
            Flog.writeline " calle " & Dir_calle & " " & Dir_num & " piso " & Dir_piso & " Dpto " & Dir_dpto
            Flog.writeline " Localidad " & Dir_localidad
            Flog.writeline " Provincia " & Dir_pcia
            
            Flog.writeline " Busco el CUIL del empleado "
            'buscar el CUIL del empleado
            StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                     " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
                     " WHERE tercero.ternro= " & rs_Traza_gan!Ternro
            OpenRecordset StrSql, rs_Cuil
            If Not rs_Cuil.EOF Then
                Cuil = Left(CStr(rs_Cuil!NroDoc), 13)
                'CUIL = Replace(CStr(Aux_CUIL), "-", "")
            Else
                Cuil = ""
                Flog.writeline " No se encontró el Cuil "
            End If
            
            'Flog.writeline " Busco el CUIT del empleado "
            'buscar el CUIT del empleado
            'StrSql = " SELECT cuit.nrodoc FROM tercero " & _
            '         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = 6) " & _
            '         " WHERE tercero.ternro= " & rs_Traza_gan!ternro
            'OpenRecordset StrSql, rs_Cuit
            'If Not rs_Cuit.EOF Then
            '    CUIT = Left(CStr(rs_Cuit!nrodoc), 13)
            '    'CUIt = Replace(CStr(CUIt), "-", "")
            'Else
            '    CUIT = ""
            '    Flog.writeline " No se encontró el Cuit "
            'End If
            
            
            'levanto los items_tope de la tabla Temporal
            
            Flog.writeline " Busco Datos en traza_gan_item_top "
            
            StrSql = "SELECT * FROM traza_gan_Item_top "
            StrSql = StrSql & " INNER JOIN item ON item.itenro = traza_gan_Item_top.itenro "
            'StrSql = StrSql & " WHERE empresa= " & Empresa
            StrSql = StrSql & " WHERE ternro =" & rs_Traza_gan!Ternro
            StrSql = StrSql & " AND pronro =" & rs_Traza_gan!pronro
            StrSql = StrSql & " ORDER BY traza_gan_Item_top.itenro"
            OpenRecordset StrSql, rs_items_tope
            
            For I = 1 To 100
                Items_TOPE(I) = 0
                Items_DDJJ(I) = 0
                Items_LIQ(I) = 0
                'FGZ - 06/03/2006 Agregué el old_liq
                Items_OLD_LIQ(I) = 0
            Next I
            
            I = 1
            
            Descuentos_Items = 0
            Gan_imponible_Items = 0
            Ded_a23_Items = 0
            Deducciones_Items = 0
            
            Flog.writeline
            Flog.writeline
            Flog.writeline "ITEMS ---- "
            Do While Not rs_items_tope.EOF
                Flog.writeline "Item" & rs_items_tope!Itenro
                If Not CBool(rs_items_tope!itesigno) Then
                    If (rs_items_tope!itetipotope = 1) Or (rs_items_tope!itetipotope = 4) Then
                        'Items_TOPE(rs_items_tope!Itenro) = rs_items_tope!Monto * Por_Deduccion / 100
                        Items_TOPE(rs_items_tope!Itenro) = IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                        
                        Ded_a23_Items = Ded_a23_Items - CDbl(IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0))
                        Flog.writeline "----suma en Ded_a23_Items " & IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                    Else
                        Items_TOPE(rs_items_tope!Itenro) = IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                        
                        Deducciones_Items = Deducciones_Items - CDbl(IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0))
                        Flog.writeline "----suma en Deducciones_Items " & IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                    End If
                    
                Else
                    Items_TOPE(rs_items_tope!Itenro) = IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                    
                    If rs_items_tope!Itenro >= 5 Then
                        Descuentos_Items = Descuentos_Items + CDbl(IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0))
                        Flog.writeline "----suma en Descuentos_Items " & IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                    End If
                    
                    Gan_imponible_Items = Gan_imponible_Items + CDbl(IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0))
                    Flog.writeline "----suma en Gan_imponible_Items " & IIf(Not EsNulo(rs_items_tope!Monto), rs_items_tope!Monto, 0)
                End If
                
                If Not IsNull(rs_items_tope!ddjj) Then
                   Items_DDJJ(rs_items_tope!Itenro) = rs_items_tope!ddjj
                End If
                
                If Not IsNull(rs_items_tope!liq) Then
                   Items_LIQ(rs_items_tope!Itenro) = rs_items_tope!liq
                End If
                    
                'FGZ - 06/03/2006 - Agregué el old_liq
                If Not IsNull(rs_items_tope!old_liq) Then
                   Items_OLD_LIQ(rs_items_tope!Itenro) = rs_items_tope!old_liq
                End If
                
                rs_items_tope.MoveNext
            Loop
            Flog.writeline
            Flog.writeline
            
            
            Flog.writeline " Busco Datos Impositivos. Retenciones/Devoluciones Historicas "
            'Datos Impositivos
            'Retenciones/Devoluciones Hist¢ricas
            Retenciones = 0
            StrSql = "SELECT * FROM ficharet " & _
                     " WHERE empleado =" & rs_Traza_gan!Ternro
            OpenRecordset StrSql, rs_ficharet
            Do While Not rs_ficharet.EOF
                If (Month(rs_ficharet!Fecha) <= Mes_Ret) And (Year(rs_ficharet!Fecha) = Ano_Ret) Then
                    If Rectifica Then
                       'Si es una rectificacion no tengo que considerar lo calculado de retenciones en el proceso actual
                       If CLng(rs_ficharet!pronro) <> CLng(rs_Traza_gan!pronro) Then
                          Retenciones = Retenciones + rs_ficharet!Importe
                       End If
                    Else
                       Retenciones = Retenciones + rs_ficharet!Importe
                    End If
'                    Retenciones = Retenciones + rs_ficharet!Importe
                End If
                rs_ficharet.MoveNext
            Loop
       
            Flog.writeline " Determinar el monto en letras "
            'Determinar el monto en letras
            Monto_Letras = EnLetras(rs_Traza_gan!saldo)
            
            Flog.writeline " Busco los datos de la Empresa "
            'Buscar los datos de la Empresa
            StrSql = "SELECT * FROM empresa "
            StrSql = StrSql & " INNER JOIN tercero ON empresa.ternro = tercero.ternro "
            StrSql = StrSql & " WHERE empresa.empnro =" & Empresa
            OpenRecordset StrSql, rs_Empresa
            Emp_Nombre = rs_Empresa!empnom
            
            Flog.writeline " Buscar el CUIT de la Empresa "
            'Buscar el CUIT de la EMPRESA
            StrSql = " SELECT cuit.nrodoc FROM tercero " & _
                     " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = 6) " & _
                     " WHERE tercero.ternro= " & rs_Empresa!Ternro
            OpenRecordset StrSql, rs_Cuit
            If Not rs_Cuit.EOF Then
                Emp_CUIT = Left(CStr(rs_Cuit!NroDoc), 13)
                'emp_CUIt = Replace(CStr(CUIt), "-", "")
            Else
                Emp_CUIT = ""
                Flog.writeline " No se encontró el CUIT de la Empresa "
            End If
            
            Flog.writeline " Busco los datos del Tercero "
            'Busco los datos del Tercero
            StrSql = "SELECT terape,terape2,ternom,ternom2 FROM tercero"
            StrSql = StrSql & " WHERE ternro =" & rs_Traza_gan!Ternro
            OpenRecordset StrSql, rs_Tercero
            If rs_Tercero.EOF Then
                Flog.writeline " No se encontraron los datos del Tercero " & rs_Traza_gan!Ternro
            End If
            
            Flog.writeline " Depuracion por empresa, fecha y empleado "
            'Depuracion por empresa, fecha y empleado
            StrSql = "DELETE FROM rep19 WHERE "
            'StrSql = StrSql & "pliqnro =" & rs_Proceso!pliqnro
            'StrSql = StrSql & " AND pronro =" & rs_Proceso!pronro
            StrSql = StrSql & " ternro =" & rs_Traza_gan!Ternro
            StrSql = StrSql & " AND fecha =" & ConvFecha(FechaHasta)
            StrSql = StrSql & " AND empresa =" & Empresa
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'FGZ - 19/04/2004
'            Gan_imponible = IIf(Not IsNull(rs_Traza_gan!ganneta), rs_Traza_gan!ganneta, 0)
'            Deducciones = IIf(Not IsNull(rs_Traza_gan!Deducciones), rs_Traza_gan!Deducciones, 0)
            Gan_imponible = Gan_imponible_Items
            Deducciones = Deducciones_Items
            
            Flog.writeline " Declaracion de Ganancias Anual:" & Anual_Final
            
            'FGZ - 26/03/2007 - Agregado de logs
            Flog.writeline
            Flog.writeline " Año de Retencion:" & Ano_Ret
            Flog.writeline " Gan_imponible_Items:  " & Gan_imponible
            Flog.writeline "    Descuentos_Items:  " & Descuentos_Items
            Flog.writeline "    Deducciones_Items: " & Deducciones_Items
            Flog.writeline "    Ded_a23_Items:     " & Ded_a23_Items
            Flog.writeline
            
            'le aplico los porcentajes a los items que suman en el art 23
            If Ano_Ret >= 2000 And Gan_imponible > 0 Then
                
                Flog.writeline " Valor a Buscar en la escala de deduccion: " & (Gan_imponible + Deducciones - Items_TOPE(50))
                
                If Not Anual_Final Then
                
                    Flog.writeline " Mes de Retencion: 12 - es Final"
                    
                    StrSql = "SELECT * FROM escala_ded " & _
                             " WHERE esd_topeinf <= " & (Gan_imponible + Deducciones - Items_TOPE(50)) & _
                             " AND esd_topesup >=" & (Gan_imponible + Deducciones - Items_TOPE(50))
                Else
                    
                    Flog.writeline " Mes de Retencion: " & Mes_Ret
                    
                    StrSql = "SELECT * FROM escala_ded " & _
                             " WHERE esd_topeinf <= " & ((Gan_imponible + Deducciones - Items_TOPE(50)) / Mes_Ret * 12) & _
                             " AND esd_topesup >=" & ((Gan_imponible + Deducciones - Items_TOPE(50)) / Mes_Ret * 12)
                
                End If
                
                OpenRecordset StrSql, rs_escala_ded
            
                If Not rs_escala_ded.EOF Then
                    Por_Deduccion = rs_escala_ded!esd_porcentaje
                Else
                    Por_Deduccion = 100
                End If
                
                Flog.writeline "Porcentaje de deduccion: " & Por_Deduccion
            Else
                Flog.writeline " No se puede calcular el porcentaje ==> 100%"
                Por_Deduccion = 100
            End If
            
            If ((FechaHasta <= HastaRG2866) And (DesdeRG2866 <= FechaHasta) And (AcuRG2866 <> 0)) Then
                
                Flog.writeline "Aplica RG 2866. Buscando Acumulador"
                StrSql = "SELECT SUM(ammonto) suma FROM acu_mes"
                StrSql = StrSql & " WHERE ternro = " & rs_Traza_gan!Ternro
                StrSql = StrSql & " AND acunro = " & AcuRG2866
                StrSql = StrSql & " AND amanio = 2010 "
                StrSql = StrSql & " AND ammes <= " & Month(FechaHasta)
                OpenRecordset StrSql, rs_Fases
                If Not rs_Fases.EOF Then
                    ValorAcuRG2866 = IIf(EsNulo(rs_Fases!suma), 0, rs_Fases!suma)
                    ValorAcuRG2866 = Abs(ValorAcuRG2866)
                    Flog.writeline "Valor de acumulador RG 2866 = " & ValorAcuRG2866
                    Retenciones = Retenciones + ValorAcuRG2866
                Else
                    ValorAcuRG2866 = 0
                    Flog.writeline "No se encontro acumulador RG 2866"
                End If
                
            End If

            
            
            Flog.writeline " Inserto en Rep19 modificado "
            'Inserto en Rep19
            StrSql = "INSERT INTO rep19 (bpronro,pliqnro,pronro,iduser,fecha,hora,empresa,ternro"
            StrSql = StrSql & ",empleg,terape,ternom,hasta,desde,promo,ano,dir_calle,dir_cp,dir_num,dir_dpto,dir_piso,dir_localidad,dir_pcia,direccion"
            StrSql = StrSql & ",ori_rect,cuil,cuit,retenciones,estimados"
            StrSql = StrSql & ",entidad1,entidad2,entidad3,entidad4,entidad5,entidad6,entidad7,entidad8,entidad9,entidad10,entidad11,entidad12,entidad13,entidad14"
            StrSql = StrSql & ",Monto_entidad1,Monto_entidad2,Monto_entidad3,Monto_entidad4,Monto_entidad5,Monto_entidad6,Monto_entidad7,Monto_entidad8,Monto_entidad9,Monto_entidad10,Monto_entidad11,Monto_entidad12,Monto_entidad13,Monto_entidad14"
            StrSql = StrSql & ",Cuit_entidad1,Cuit_entidad2,Cuit_entidad3,Cuit_entidad4,Cuit_entidad5,Cuit_entidad6,Cuit_entidad7,Cuit_entidad8,Cuit_entidad9,Cuit_entidad10,Cuit_entidad11,Cuit_entidad12,Cuit_entidad13,Cuit_entidad14"
            StrSql = StrSql & ",Total_entidad1,Total_entidad2,Total_entidad3,Total_entidad4,Total_entidad5,Total_entidad6,Total_entidad7,Total_entidad8,Total_entidad9,Total_entidad10,Total_entidad11,Total_entidad12,Total_entidad13,Total_entidad14"
            StrSql = StrSql & ",ganneta,ganimpo,msr,nomsr,nogan,conyuge,hijo,otras_cargas,car_flia,prima_seguro,sepelio,osocial,cuota_medico"
            StrSql = StrSql & ",jubilacion,sindicato,donacion,otras,dedesp,noimpo,seguro_retiro,amortizacion,viaticos,imp_deter"
            StrSql = StrSql & ",saldo,monto_letras,emp_nombre,emp_cuit"
            StrSql = StrSql & ",prorratea,anual_final"
            
            If Not (IsNull(Suscribe) Or Suscribe = "") Then
                StrSql = StrSql & ",suscribe"
            End If
            If Not (IsNull(Caracter) Or Caracter = "") Then
                StrSql = StrSql & ",caracter"
            End If
            If Not (IsNull(Fecha_Caracter) Or Fecha_Caracter = "") Then
                StrSql = StrSql & ",fecha_caracter"
            End If
            If Not (IsNull(Fecha_Devolucion) Or Fecha_Devolucion = "") Then
                StrSql = StrSql & ",fecha_devolucion"
            End If
            If Not (IsNull(Dependencia_DGI) Or Dependencia_DGI = "") Then
                StrSql = StrSql & ",dependencia_dgi"
            End If
            
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & Bpronro & ","
            StrSql = StrSql & rs_Traza_gan!pliqnro & ","
            StrSql = StrSql & rs_Traza_gan!pronro & ","
            StrSql = StrSql & "'" & IdUser & "',"
            StrSql = StrSql & ConvFecha(FechaHasta) & ","
            StrSql = StrSql & "'" & hora & "',"
            StrSql = StrSql & rs_Traza_gan!Empresa & ","
            StrSql = StrSql & rs_Traza_gan!Ternro & ","
    
            StrSql = StrSql & rs_Traza_gan!empleg & ","
            If Not rs_Tercero.EOF Then
                Auxiliar = rs_Tercero!terape
                If Not IsNull(rs_Tercero!terape2) Then
                    Auxiliar = Auxiliar & " " & rs_Tercero!terape2
                End If
                StrSql = StrSql & "'" & Format_Str(Auxiliar, 80, False, " ") & "',"
                Auxiliar = rs_Tercero!ternom
                If Not IsNull(rs_Tercero!ternom2) Then
                    Auxiliar = Auxiliar & " " & rs_Tercero!ternom2
                End If
                StrSql = StrSql & "'" & Format_Str(Auxiliar, 80, False, " ") & "',"
            Else
                StrSql = StrSql & "' ', ' ',"
            End If
            StrSql = StrSql & ConvFecha(v_hasta) & ","
            StrSql = StrSql & ConvFecha(v_desde) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!promo), rs_Traza_gan!promo, 0) & ","
            StrSql = StrSql & Ano_Ret & ","
            StrSql = StrSql & "'" & Format_Str(Dir_calle, 30, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Dir_cp, 10, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Dir_num, 8, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Dir_dpto, 8, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Dir_piso, 8, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Dir_localidad, 30, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Dir_pcia, 30, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(Direccion, 40, False, " ") & "',"
            StrSql = StrSql & "'" & original & "',"
            StrSql = StrSql & "'" & Format_Str(Cuil, 15, False, " ") & "',"
            StrSql = StrSql & "'" & Format_Str(cuit, 15, False, " ") & "',"
            StrSql = StrSql & Retenciones & ","
            StrSql = StrSql & "0,"
            
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!Entidad1), Format_Str(rs_Traza_gan!Entidad1, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!Entidad2), Format_Str(rs_Traza_gan!Entidad2, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!Entidad3), Format_Str(rs_Traza_gan!Entidad3, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad4), Format_Str(rs_Traza_gan!entidad4, 40, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad5), Format_Str(rs_Traza_gan!entidad5, 40, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad6), Format_Str(rs_Traza_gan!entidad6, 40, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad7), Format_Str(rs_Traza_gan!entidad7, 40, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad8), Format_Str(rs_Traza_gan!entidad8, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad9), Format_Str(rs_Traza_gan!entidad9, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad10), Format_Str(rs_Traza_gan!entidad10, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad11), Format_Str(rs_Traza_gan!entidad11, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad12), Format_Str(rs_Traza_gan!entidad12, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad13), Format_Str(rs_Traza_gan!entidad13, 70, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!entidad14), Format_Str(rs_Traza_gan!entidad14, 70, False, " "), " ") & "',"
            
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!monto_entidad1), rs_Traza_gan!monto_entidad1, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!monto_entidad2), rs_Traza_gan!monto_entidad2, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad3), rs_Traza_gan!Monto_entidad3, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!monto_entidad4), rs_Traza_gan!monto_entidad4, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad5), rs_Traza_gan!Monto_entidad5, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!monto_entidad6), rs_Traza_gan!monto_entidad6, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad7), rs_Traza_gan!Monto_entidad7, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!monto_entidad8), rs_Traza_gan!monto_entidad8, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad9), rs_Traza_gan!Monto_entidad9, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad10), rs_Traza_gan!Monto_entidad10, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad11), rs_Traza_gan!Monto_entidad11, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad12), rs_Traza_gan!Monto_entidad12, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad13), rs_Traza_gan!Monto_entidad13, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Monto_entidad14), rs_Traza_gan!Monto_entidad14, 0) & ","
            
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad1), Format_Str(rs_Traza_gan!cuit_entidad1, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad2), Format_Str(rs_Traza_gan!cuit_entidad2, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad3), Format_Str(rs_Traza_gan!cuit_entidad3, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad4), Format_Str(rs_Traza_gan!cuit_entidad4, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad5), Format_Str(rs_Traza_gan!cuit_entidad5, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad6), Format_Str(rs_Traza_gan!cuit_entidad6, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad7), Format_Str(rs_Traza_gan!cuit_entidad7, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad8), Format_Str(rs_Traza_gan!cuit_entidad8, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad9), Format_Str(rs_Traza_gan!cuit_entidad9, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad10), Format_Str(rs_Traza_gan!cuit_entidad10, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad11), Format_Str(rs_Traza_gan!cuit_entidad11, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad12), Format_Str(rs_Traza_gan!cuit_entidad12, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad13), Format_Str(rs_Traza_gan!cuit_entidad13, 13, False, " "), " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(rs_Traza_gan!cuit_entidad14), Format_Str(rs_Traza_gan!cuit_entidad14, 13, False, " "), " ") & "',"
            
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!monto_entidad1), rs_Traza_gan!total_entidad1, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad2), rs_Traza_gan!Total_entidad2, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad3), rs_Traza_gan!Total_entidad3, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad4), rs_Traza_gan!Total_entidad4, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!total_entidad5), rs_Traza_gan!total_entidad5, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad6), rs_Traza_gan!Total_entidad6, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad7), rs_Traza_gan!Total_entidad7, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad8), rs_Traza_gan!Total_entidad8, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad9), rs_Traza_gan!Total_entidad9, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad10), rs_Traza_gan!Total_entidad10, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad11), rs_Traza_gan!Total_entidad11, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad12), rs_Traza_gan!Total_entidad12, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad13), rs_Traza_gan!Total_entidad13, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!Total_entidad14), rs_Traza_gan!Total_entidad14, 0) & ","
            
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!ganneta), rs_Traza_gan!ganneta, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!ganimpo), rs_Traza_gan!ganimpo, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!msr), rs_Traza_gan!msr + Items_TOPE(50), 0 + Items_TOPE(50)) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!nomsr), rs_Traza_gan!nomsr, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!nogan), rs_Traza_gan!nogan, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!conyuge), (rs_Traza_gan!conyuge * Por_Deduccion / 100), 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!hijo), (rs_Traza_gan!hijo * Por_Deduccion / 100), 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!otras_cargas), (rs_Traza_gan!otras_cargas * Por_Deduccion / 100), 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!car_flia), (rs_Traza_gan!car_flia * Por_Deduccion / 100), 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!prima_seguro), rs_Traza_gan!prima_seguro, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!sepelio), rs_Traza_gan!sepelio, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!osocial), rs_Traza_gan!osocial, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!cuota_medico), rs_Traza_gan!cuota_medico, 0) & ","
            
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!jubilacion), rs_Traza_gan!jubilacion, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!sindicato), rs_Traza_gan!sindicato, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!donacion), rs_Traza_gan!donacion, 0) & ","
            'StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!otras), (rs_Traza_gan!otras * Por_Deduccion / 100), 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!otras), rs_Traza_gan!otras - Items_TOPE(50), 0 - Items_TOPE(50)) & ","
            'FGZ - 04/12/2013 -------------------------------------------------------------------------------------------------
            'StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!dedesp), (rs_Traza_gan!dedesp * Por_Deduccion / 100), 0) & ","
            Ajuste = 0
            'If (rs_Traza_gan!Ret_Mes > 9) Or (rs_Traza_gan!Ret_Mes = 9 And buliq_periodo!pliqmes = 9) Or (rs_Traza_gan!Ret_Mes = 9 And buliq_periodo!pliqmes = 8) Then
            If (Mes_Ret >= 9) Then
                'Ajuste = (6220.8 * (Por_Deduccion / 100) * (Mes_Ret - 8))
                Ajuste = Abs(Items_TOPE(16) - (rs_Traza_gan!dedesp * Por_Deduccion / 100))
            End If
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!dedesp), ((rs_Traza_gan!dedesp * Por_Deduccion / 100) + Ajuste), 0) & ","
            'FGZ - 04/12/2013 -------------------------------------------------------------------------------------------------
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!noimpo), (rs_Traza_gan!noimpo * Por_Deduccion / 100), 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!seguro_retiro), rs_Traza_gan!seguro_retiro, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!amortizacion), rs_Traza_gan!amortizacion, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!viaticos), rs_Traza_gan!viaticos, 0) & ","
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!imp_deter), rs_Traza_gan!imp_deter, 0) & ","
    
            StrSql = StrSql & IIf(Not IsNull(rs_Traza_gan!saldo), rs_Traza_gan!saldo, 0) & ","
            StrSql = StrSql & "'" & Left(Monto_Letras, 50) & "',"
            StrSql = StrSql & "'" & Format_Str(Emp_Nombre, 40, False, " ") & "',"
            StrSql = StrSql & "'" & IIf(Not IsNull(Emp_CUIT), Format_Str(Emp_CUIT, 15, False, " "), "N.D.") & "',"
            
            StrSql = StrSql & CInt(Prorratea) & ","
            StrSql = StrSql & CInt(Anual_Final)
            
            If Not (IsNull(Suscribe) Or Suscribe = "") Then
                StrSql = StrSql & ",'" & Format_Str(Suscribe, 30, False, " ") & "'"
            End If
            If Not (IsNull(Caracter) Or Caracter = "") Then
                StrSql = StrSql & ",'" & Format_Str(Caracter, 30, False, " ") & "'"
            End If
            If Not (IsNull(Fecha_Caracter) Or Fecha_Caracter = "") Then
                StrSql = StrSql & "," & ConvFecha(CDate(Fecha_Caracter))
            End If
            If Not (IsNull(Fecha_Devolucion) Or Fecha_Devolucion = "") Then
                StrSql = StrSql & "," & ConvFecha(CDate(Fecha_Devolucion))
            End If
            If Not (IsNull(Dependencia_DGI) Or Dependencia_DGI = "") Then
                StrSql = StrSql & ",'" & Format_Str(Dependencia_DGI, 30, False, " ") & "'"
            End If
            
            StrSql = StrSql & ")"
            Flog.writeline " SQL a insertar " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline " insertó con exito "
            
            Flog.writeline " llamo a Suegan21(" & Bpronro & "," & rs_Traza_gan!Ternro & "," & Mes_Ret & "," & Ano_Ret & "," & Empresa & ")"
'            'FGZ - 18/04/2005
'            Gan_imponible = IIf(Not IsNull(rs_Traza_gan!ganneta), rs_Traza_gan!ganneta, 0)
            'FGZ - agregue el parametro fecha
            Call Suegan21(Bpronro, rs_Traza_gan!Ternro, Mes_Ret, Ano_Ret, Empresa, Gan_imponible, Por_Deduccion, Aux_Fecha_A_Utilizar)
            Flog.writeline " retornó con exito de Suegan21"
            
            'RUN reportes/suegan21.p(BUFFER rep19,  tt_derecha.ttternro-der, mes-ret, ano-ret)
        
        End If
    End If
    Flog.writeline " Actualizo el progreso del Proceso "
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
    'Siguiente registro
    rs_Traza_gan.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans

If rs_Proceso.State = adStateOpen Then rs_Proceso.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
If rs_Rep19.State = adStateOpen Then rs_Rep19.Close
If rs_CabdomPrincipal.State = adStateOpen Then rs_CabdomPrincipal.Close
If rs_Cabdom.State = adStateOpen Then rs_Cabdom.Close
If rs_Detdom.State = adStateOpen Then rs_Detdom.Close
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
If rs_Cuit.State = adStateOpen Then rs_Cuit.Close
If rs_ficharet.State = adStateOpen Then rs_ficharet.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_items_tope.State = adStateOpen Then rs_items_tope.Close

Set rs_Proceso = Nothing
Set rs_Traza_gan = Nothing
Set rs_Empresa = Nothing
Set rs_Rep19 = Nothing
Set rs_CabdomPrincipal = Nothing
Set rs_Cabdom = Nothing
Set rs_Detdom = Nothing
Set rs_Fases = Nothing
Set rs_Localidad = Nothing
Set rs_Provincia = Nothing
Set rs_Cuil = Nothing
Set rs_Cuit = Nothing
Set rs_ficharet = Nothing
Set rs_Tercero = Nothing
Set rs_Empleado = Nothing
Set rs_items_tope = Nothing

Exit Sub

CE:
    HuboError = True
    Flog.writeline " Error " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub



Public Sub LevantarParamteros(ByVal Bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim Separador As String
Dim pos1 As Integer
Dim pos2 As Integer


Dim FechaHasta As Date
Dim Tope_Min As Single
Dim Tope_Max As Single
Dim Anual_Final As Boolean
Dim Suscribe As String
Dim Caracter As String
Dim Fecha_Caracter As String 'Date
Dim Fecha_Devolucion As String 'Date
Dim Dependencia_DGI As String
Dim Empresa As Long
Dim Prorratea As Boolean
Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long

Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long

Dim AgrupaTE1 As Boolean
Dim AgrupaTE2 As Boolean
Dim AgrupaTE3 As Boolean
Dim Agrupado As Boolean

Dim Todos_Empleados As Boolean

Dim Lugar As String
Dim original As Integer
Dim ArrParam
On Error GoTo CE

'Inicializacion
Agrupado = False
Tenro1 = 0
Tenro2 = 0
Tenro3 = 0
AgrupaTE1 = False
AgrupaTE2 = False
AgrupaTE3 = False
Lugar = ""

Separador = "@"
' Levanto cada parametro por separado, el separador de parametros esta definido arriba
If Not IsNull(parametros) Then

    ArrParam = Split(parametros, Separador)
    
    If UBound(ArrParam) >= 1 Then
        FechaHasta = CDate(ArrParam(0))
        Prorratea = CBool(ArrParam(1))
        Tope_Min = CSng(ArrParam(2))
        Tope_Max = CSng(ArrParam(3))
        Anual_Final = CBool(ArrParam(4))
        Suscribe = ArrParam(5)
        Caracter = ArrParam(6)
        Fecha_Caracter = ArrParam(7)
        Fecha_Devolucion = ArrParam(8)
        Dependencia_DGI = ArrParam(9)
        Empresa = ArrParam(10)
        Todos_Empleados = CBool(ArrParam(11))
        'no se usa asi que lo reutilizo para todos los empleados
        SoloConRetenciones = CBool(ArrParam(12))
        original = ArrParam(13)
        If UBound(ArrParam) = 14 Then
            Lugar = ArrParam(14)
        End If
        

    End If
End If

Flog.writeline
Flog.writeline "Parametros ===>"
Flog.writeline Espacios(Tabulador * 2) & "Fecha Hasta: " & FechaHasta
Flog.writeline Espacios(Tabulador * 2) & "Empresa: " & Empresa
Flog.writeline Espacios(Tabulador * 2) & "Prorratea: " & Prorratea
Flog.writeline Espacios(Tabulador * 2) & "Tope Min: " & Tope_Min
Flog.writeline Espacios(Tabulador * 2) & "Tope Max: " & Tope_Max
Flog.writeline Espacios(Tabulador * 2) & "Todos los Empleados: " & Todos_Empleados
Flog.writeline Espacios(Tabulador * 2) & "Con Retenciones: " & SoloConRetenciones
Flog.writeline Espacios(Tabulador * 2) & "Anual: " & Anual_Final
Flog.writeline Espacios(Tabulador * 2) & "Suscribe: " & Suscribe
Flog.writeline Espacios(Tabulador * 2) & "Caracter: " & Caracter
Flog.writeline Espacios(Tabulador * 2) & "Fecha Caracter: " & Fecha_Caracter
Flog.writeline Espacios(Tabulador * 2) & "Fecha Devolucion: " & Fecha_Devolucion
Flog.writeline Espacios(Tabulador * 2) & "Dependencia DGI: " & Dependencia_DGI
Flog.writeline Espacios(Tabulador * 2) & "Solo Con Retenciones: " & SoloConRetenciones
Flog.writeline Espacios(Tabulador * 2) & "original: " & original
Flog.writeline Espacios(Tabulador * 2) & "Lugar: " & Lugar


Flog.writeline
Flog.writeline
Call Suegan20(Bpronro, FechaHasta, Prorratea, Tope_Min, Tope_Max, Anual_Final, Suscribe, Caracter, Fecha_Caracter, Fecha_Devolucion, Dependencia_DGI, Empresa, Todos_Empleados, Lugar, original)


Exit Sub

CE:
    HuboError = True
    Flog.writeline " Error " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub


Public Sub Suegan21OLD(ByVal Bpronro As Long, ByVal Tercero As Long, ByVal Mes_Ret As Long, ByVal Ano_Ret As Long, ByVal Empresa As Long, ByVal Gan_imponible As Single, ByVal Por_Deduccion As Single, ByVal Aux_Fecha_A_Utilizar As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Completa los datos en la tabla reporte19, necesarios para la impresion de la
'               parte posterior del fromulario
' Autor      : Javie Iraztorza
' Fecha      : 09/05/2000
' Modif      : H.J.I 16/03/2001
' Traduccion : FGZ
' Fecha      : 22/04/2004
' Ult. Mod   : FGZ - 15/03/2007 Problemas con los items 6 y 13
' --------------------------------------------------------------------------------------------
Dim I          As Integer
Dim Primera    As Boolean
Dim Totalitem  As Single
Dim J          As Integer
Dim Actualiza As Boolean
Dim Itenro As Integer
Dim Solo_Total As Boolean
Dim ListaItenro As String
Dim LiqIte6 As Double
Dim DDJJ_Item6 As Double
Dim HayDDJJ_Item13 As Boolean

Dim rs_Rep19 As New ADODB.Recordset
Dim rs_OS As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Desliq As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_Items As New ADODB.Recordset
Dim rs_Ter_doc As New ADODB.Recordset
Dim rs_Traza_gan_Aux As New ADODB.Recordset
Dim rs_Traza_gan_Ite_Top_Aux As New ADODB.Recordset

Dim descObraSoc1
Dim descObraSoc2
Dim montoX
Dim cuitX
Dim entidadX
Dim posicionOtraObraSocial
Dim cantidadI
Dim Encontro As Boolean
Dim StrSql_Aux As String

On Error GoTo CE

'Inicializacion
I = 0
J = 0

StrSql = "SELECT * FROM  rep19 "
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
OpenRecordset StrSql, rs_Rep19

Flog.writeline "ITEM 1 y 2 "
Flog.writeline "CONTROLO SI EL EMPLEADO TRABAJO EN ALGUNA EMPRESA ANTERIOR"
'********************* ITEM 1 y 2 *************************************
'********************** CONTROL EMPRESAS ANTERIORES ************************************
StrSql = "SELECT sum(desmondec) AS suma, descuit, desrazsoc FROM desmen "
StrSql = StrSql & " WHERE desmen.empleado =" & Tercero
StrSql = StrSql & " AND desmen.desano =" & Ano_Ret
StrSql = StrSql & " AND desmen.itenro IN (1,2) "
StrSql = StrSql & " GROUP BY descuit, desrazsoc "

OpenRecordset StrSql, rs_Desmen
If rs_Desmen.EOF Then
    Flog.writeline "no hay items de tipo 1 o 2 "
End If

Do While Not rs_Desmen.EOF
    If Not IsNull(rs_Desmen!descuit) Then
       If Trim(rs_Desmen!descuit) <> "" Then
            I = I + 1
          
            StrSql = "UPDATE rep19 SET "
            StrSql = StrSql & "cuit_entidad1" & I & " ='" & rs_Desmen!descuit & "'"
            StrSql = StrSql & ",entidad1" & I & " ='" & Format_Str(rs_Desmen!desrazsoc, 70, False, " ") & "'"
            StrSql = StrSql & ",monto_entidad1" & I & " = " & rs_Desmen!suma
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            
            Flog.writeline "Update Sql " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
       End If
    End If

    rs_Desmen.MoveNext
Loop

rs_Desmen.Close

'Inicializacion
I = 0
J = 0

Flog.writeline " levanto los items_tope de la tabla Temporal "

Flog.writeline "ITEM 13 Obra Social Privada "
Flog.writeline "COUTAS MEDICO ASISTENCIALES "
'********************* ITEM 13 Obra Social Privada *************************************
'********************** COUTAS MEDICO ASISTENCIALES ************************************

StrSql = "UPDATE rep19 SET "
StrSql = StrSql & " total_entidad1 =0"
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "SELECT * FROM empleado "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 17"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE empleado.ternro = " & Tercero & " AND "
StrSql = StrSql & " his_estructura.tenro = 17 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(rs_Rep19!Hasta) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(rs_Rep19!Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
OpenRecordset StrSql, rs_OS

'Solo muestro la obra social si es que tiene el item 13 liquida
'FGZ - 15/03/2007 -------------Cambié esto -----------------
'If Not rs_OS.EOF And Items_LIQ(13) <> 0 Then
'por esto

'Busco Si hay DDJJ para el item 13
StrSql = "SELECT * FROM desmen "
StrSql = StrSql & "WHERE desmen.itenro = 13 "
StrSql = StrSql & " AND desmen.empleado =" & Tercero
'19/05/2009 - Martin Ferraro - Faltaba la validacion de fechas para desmen
StrSql = StrSql & " AND desmen.desfecdes <= " & ConvFecha(rs_Rep19!Desde)
StrSql = StrSql & " AND desmen.desfechas >= " & ConvFecha(rs_Rep19!Desde)

OpenRecordset StrSql, rs_Desmen
HayDDJJ_Item13 = Not rs_Desmen.EOF

'Busco Si hay DDJJ para el item 6 ==> lo resto del total sino luego en el detalle de las ddjj
StrSql = "SELECT * FROM desmen "
StrSql = StrSql & "WHERE desmen.itenro = 6 "
StrSql = StrSql & " AND desmen.empleado =" & Tercero
'19/05/2009 - Martin Ferraro - Faltaba la validacion de fechas para desmen
StrSql = StrSql & " AND desmen.desfecdes <= " & ConvFecha(rs_Rep19!Hasta)
StrSql = StrSql & " AND desmen.desfechas >= " & ConvFecha(rs_Rep19!Hasta)
OpenRecordset StrSql, rs_Desmen
If rs_Desmen.EOF Then
    DDJJ_Item6 = 0
Else
    DDJJ_Item6 = rs_Desmen!desmondec
End If
If HayDDJJ_Item13 Then
    'Como voy a mostrar la ddjj del 13 ==> acumulo el del 6
    LiqIte6 = Abs(Items_TOPE(6))
Else
    'como no hay ddjj 13 y si hay ddjj 6 ==> separo dado que voy a mostrar ddjj del 6
    LiqIte6 = Abs(Items_TOPE(6)) - Abs(DDJJ_Item6)
End If
If Items_TOPE(6) < 0 Then
    LiqIte6 = LiqIte6 * -1
End If

If Not rs_OS.EOF And LiqIte6 <> 0 Then
'FGZ - 15/03/2007 ------------------------------------------
    StrSql = " SELECT * FROM tercero "
    StrSql = StrSql & " INNER JOIN replica_estr ON replica_estr.origen = tercero.ternro "
    StrSql = StrSql & " WHERE replica_estr.estrnro= " & rs_OS!Estrnro
    OpenRecordset StrSql, rs_Tercero
    If Not rs_Tercero.EOF Then
        StrSql = " SELECT * FROM ter_doc"
        StrSql = StrSql & " WHERE ter_doc.ternro = " & rs_Tercero!Ternro
        StrSql = StrSql & " AND ter_doc.tidnro = 6 "
        OpenRecordset StrSql, rs_Ter_doc
        
        StrSql = "UPDATE rep19 SET "
        If Not rs_Ter_doc.EOF Then
            StrSql = StrSql & "cuit_entidad1 ='" & Format_Str(rs_Ter_doc!NroDoc, 13, False, " ") & "'"
            descObraSoc2 = rs_Ter_doc!NroDoc
        Else
            StrSql = StrSql & "cuit_entidad1 =' '"
            descObraSoc2 = ""
        End If
        StrSql = StrSql & ",entidad1 ='" & Format_Str(rs_Tercero!terrazsoc, 70, False, " ") & "'"
        
        'FGZ - 15/03/2007 ------------------------------
        'Antes
'        'StrSql = StrSql & ",total_entidad1 = total_entidad1 + " & rs_Rep19!osocial
'        StrSql = StrSql & ",total_entidad1 = 0 "
'        'StrSql = StrSql & ",monto_entidad1 = " & rs_Rep19!osocial
'        StrSql = StrSql & ",monto_entidad1 = 0 "
        'Ahora
        StrSql = StrSql & ",total_entidad1 = total_entidad1 + " & Abs(LiqIte6)
        StrSql = StrSql & ",monto_entidad1 = " & LiqIte6
        'FGZ - 15/03/2007 ------------------------------
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        
        descObraSoc1 = rs_Tercero!terrazsoc
        
    Else
        StrSql = "UPDATE rep19 SET "
        StrSql = StrSql & "cuit_entidad1 =' '"
        StrSql = StrSql & ",entidad1 =''"
        'StrSql = StrSql & ",total_entidad1 = total_entidad1 + " & rs_Rep19!osocial
        StrSql = StrSql & ",total_entidad1 = 0 "
        'StrSql = StrSql & ",monto_entidad1 = " & rs_Rep19!osocial
        StrSql = StrSql & ",monto_entidad1 = 0 "
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
    End If
Else
    StrSql = "UPDATE rep19 SET "
    StrSql = StrSql & "cuit_entidad1 =' '"
    StrSql = StrSql & ",entidad1 =''"
    StrSql = StrSql & ",total_entidad1 = 0 "
    StrSql = StrSql & ",monto_entidad1 = 0 "
    StrSql = StrSql & " WHERE bpronro=" & Bpronro
    StrSql = StrSql & " AND ternro=" & Tercero
End If
Flog.writeline "Update Sql " & StrSql
objConn.Execute StrSql, , adExecuteNoRecords
    
    
'FGZ - 15/03/2007 ------------------------------
'Modifiqué esto
'i = 1
'X
If LiqIte6 <> 0 Then
    I = 2
    cantidadI = 1
Else
    I = 1
    cantidadI = 0
End If

'FGZ - 15/03/2007 ------------------------------


Itenro = 13
Primera = False
'FGZ - 15/03/2007 ----- agregué esto ------------
If LiqIte6 <> 0 Then
    If HayDDJJ_Item13 Then
        ListaItenro = "13"
    Else
        ListaItenro = "6"
    End If
Else
    ListaItenro = "6,13"
End If
'FGZ - 15/03/2007 -------------------------------

Flog.writeline "Busco la declaracion jurada "
'Busco la declaracion jurada
StrSql = "SELECT * FROM desmen "
StrSql = StrSql & "WHERE desmen.itenro IN (" & ListaItenro & ")"
StrSql = StrSql & " AND desmen.empleado =" & Tercero
StrSql = StrSql & " ORDER BY desmen.itenro, desmen.desmondec"
OpenRecordset StrSql, rs_Desmen
If rs_Desmen.EOF Then
    Flog.writeline "no hay "
End If

posicionOtraObraSocial = 0

Do While Not rs_Desmen.EOF
    Flog.writeline "declaracion jurada item " & rs_Desmen!Itenro
    Actualiza = False
    If Year(rs_Desmen!desfecdes) = Ano_Ret Then
        If rs_Desmen!desmondec <> 0 Then
            Flog.writeline "declaracion jurada monto " & rs_Desmen!desmondec
            Actualiza = True
            cantidadI = cantidadI + 1
            'StrSql = "UPDATE rep19 SET "
            'StrSql = StrSql & "total_entidad1 = total_entidad1 +" & rs_Desmen!desmondec
            'StrSql = StrSql & "total_entidad1 = total_entidad1 + " & Items_TOPE(Itenro)
        Else
            Flog.writeline "declaracion jurada monto en 0"
        End If
        
        Flog.writeline "Tope " & Items_TOPE(Itenro)
        If rs_Desmen!desmondec = 0 And Items_TOPE(Itenro) <> 0 Then
           posicionOtraObraSocial = I
        End If
        
        'If (i < 2) And (rs_Desmen!desmondec <> 0) Then   'solo puedo poner dos en la DDJJ
        If (I <= 2) Then   'solo puedo poner dos en la DDJJ
            'If Not Actualiza Then
                Actualiza = True
                StrSql = "UPDATE rep19 SET "
            'End If
            
            ' Tiene DDJJ, tomo el valor de la DDJJ sino supera el tope
            If rs_Desmen!desmondec <> 0 Then
               '12/04/2010 - Se cambio la siguiente sentencia
               'montoX = IIf(Abs(rs_Desmen!desmondec) + Abs(montoX) <= Abs(Items_TOPE(Itenro)), Abs(rs_Desmen!desmondec), Abs(Items_TOPE(Itenro) - Abs(montoX)))
               montoX = Abs(Items_TOPE(Itenro))
               cuitX = rs_Desmen!descuit
               entidadX = rs_Desmen!desrazsoc
            Else  'Si tiene DDJJ en 0 y tiene liquidado tomo lo liquidado
                If Items_LIQ(Itenro) <> 0 Then
                  Flog.writeline "tiene DDJJ en 0 y tiene liquidado tomo lo liquidado"
                  '13/01/2010 - Martin Ferraro - Tomaba mal el tope
                  'montoX = Abs(Items_LIQ(Itenro))
                  montoX = Abs(Items_TOPE(Itenro))
                  'montoX = 0
                  cuitX = rs_Desmen!descuit
                  entidadX = rs_Desmen!desrazsoc
                Else
                    Flog.writeline "tiene DDJJ en 0 y no tiene liquidado tomo lo liquidado. Buscar ultima del año"
                    
                    'montoX = Abs(Items_OLD_LIQ(Itenro))
                    montoX = IIf(Abs(Items_OLD_LIQ(Itenro)) + Abs(montoX) <= Abs(Items_TOPE(Itenro)), Abs(Items_OLD_LIQ(Itenro)), Abs(Items_TOPE(Itenro) - Abs(montoX)))
                    cuitX = rs_Desmen!descuit
                    entidadX = rs_Desmen!desrazsoc
                    
'                    'FGZ - si el proceso es de ajuste ==> deberia buscar lo liquidado en el proceso normal < a esta fecha: Aux_Fecha_A_Utilizar
'                    StrSql_Aux = "SELECT pronro FROM  traza_gan "
'                    StrSql_Aux = StrSql_Aux & " WHERE traza_gan.fecha_pago < " & ConvFecha(Aux_Fecha_A_Utilizar)
'                    StrSql_Aux = StrSql_Aux & " AND traza_gan.fecha_pago >= " & ConvFecha(C_Date("01/01/" & Year(Aux_Fecha_A_Utilizar)))
'                    StrSql_Aux = StrSql_Aux & " AND traza_gan.ternro = " & Tercero
'                    StrSql_Aux = StrSql_Aux & " ORDER BY fecha_pago DESC "
'                    If rs_Traza_gan_Aux.State = adStateOpen Then rs_Traza_gan_Aux.Close
'                    Encontro = False
'                    OpenRecordset StrSql_Aux, rs_Traza_gan_Aux
'                    Do While Not rs_Traza_gan_Aux.EOF And Not Encontro
'                        'Busco el item 13
'                        StrSql_Aux = "SELECT * FROM traza_gan_Item_top "
'                        StrSql_Aux = StrSql_Aux & " WHERE ternro =" & Tercero
'                        StrSql_Aux = StrSql_Aux & " AND pronro =" & rs_Traza_gan_Aux!pronro
'                        StrSql_Aux = StrSql_Aux & " AND itenro = 13"
'                        If rs_Traza_gan_Ite_Top_Aux.State = adStateOpen Then rs_Traza_gan_Ite_Top_Aux.Close
'                        OpenRecordset StrSql_Aux, rs_Traza_gan_Ite_Top_Aux
'                        If Not rs_Traza_gan_Ite_Top_Aux.EOF Then
'                            If Not EsNulo(rs_Traza_gan_Ite_Top_Aux!liq) Then
'                                If rs_Traza_gan_Ite_Top_Aux!liq <> 0 Then
'                                    'montoX = Abs(rs_Traza_gan_Ite_Top_Aux!liq) + Abs(Items_DDJJ(Itenro))
'                                    'montoX = Abs(rs_Traza_gan_Ite_Top_Aux!Monto) '+ Abs(Items_DDJJ(Itenro))
'                                    montoX = IIf(Abs(rs_Traza_gan_Ite_Top_Aux!Monto) + Abs(montoX) <= Abs(Items_TOPE(Itenro)), Abs(rs_Traza_gan_Ite_Top_Aux!Monto), Abs(Items_TOPE(Itenro) - Abs(montoX)))
'                                    cuitX = rs_Desmen!descuit
'                                    entidadX = rs_Desmen!desrazsoc
'                                    Encontro = True
'
'                                End If
'                            End If
'                        End If
'
'                        rs_Traza_gan_Aux.MoveNext
'                    Loop
'                    If Not Encontro Then
'                        montoX = 0
'                        cuitX = ""
'                        entidadX = ""
'                        Actualiza = False
'                        i = i - 1
'                    End If
               End If
            End If
            
            Primera = False
            Select Case I
            Case 1:
                'StrSql = StrSql & " total_entidad1 = total_entidad1 + " & montoX
                StrSql = StrSql & " total_entidad1 = " & Items_TOPE(Itenro)
                StrSql = StrSql & ",monto_entidad1 = " & montoX
                StrSql = StrSql & ",cuit_entidad1 = '" & Format_Str(cuitX, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad1 = '" & Format_Str(entidadX, 70, False, " ") & "'"
            Case 2:
                'FGZ - 15/03/2007 ------------------------------
                'Cambié esto
                
'                'StrSql = StrSql & " total_entidad1 = total_entidad1 + " & montoX
'                'StrSql = StrSql & " total_entidad2 = total_entidad2 + " & montoX
'                StrSql = StrSql & "monto_entidad2 = " & montoX
                
                'por esto
                
                StrSql = StrSql & "total_entidad1 = total_entidad1 + " & montoX
                StrSql = StrSql & ",monto_entidad2 = " & montoX
                'FGZ - 15/03/2007 ------------------------------
                StrSql = StrSql & ",cuit_entidad2 = '" & Format_Str(cuitX, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad2 = '" & Format_Str(entidadX, 70, False, " ") & "'"
            Case Else: 'los demas no se pueden dar
                Flog.writeline
                Flog.writeline "No hay entidad " & I
                Flog.writeline
            End Select

            I = I + 1
        Else
         If I >= 3 Then
            'FGZ - 15/03/2007 ------------------------------
            'Cambie esta lunea a TRUE
            Actualiza = True
            'FGZ - 15/03/2007 ------------------------------
          End If
        End If
        
        If (Actualiza And (I <= 3)) And cantidadI < 3 Then
        'If (Actualiza And (i < 3)) And cantidadI < 3 Then
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If Actualiza Then
                StrSql = "UPDATE rep19 SET "
                'FGZ - 15/03/2007 ------------------------------
                'Cambié esta linea
                'StrSql = StrSql & " total_entidad1 = total_entidad1 + " & rs_Desmen!desmondec
                'X esta
                StrSql = StrSql & " total_entidad1 = total_entidad1 + " & Abs(rs_Desmen!desmondec)
                'FGZ - 15/03/2007 ------------------------------
                StrSql = StrSql & " WHERE bpronro=" & Bpronro
                StrSql = StrSql & " AND ternro=" & Tercero
                Flog.writeline " Update SQL " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    Else
        Flog.writeline " Año de retencion <> Ano_Ret " & Ano_Ret
    End If
    rs_Desmen.MoveNext
Loop

Flog.writeline " Busco las liquidaciones anteriores "
If posicionOtraObraSocial = 0 Then
   posicionOtraObraSocial = I
End If
I = I + 1

'Busco las liquidaciones anteriores
'StrSql = "SELECT * FROM desliq where desliq.itenro = " & Itenro
'StrSql = StrSql & " AND desliq.empleado =" & Tercero
'OpenRecordset StrSql, rs_Desliq
'If rs_Desliq.EOF Then
'    Flog.writeline " No hay "
'End If
'
'Do While Not rs_Desliq.EOF
'    Solo_Total = True
'    If (Month(rs_Desliq!DLFecha) <= Mes_Ret) And (Year(rs_Desliq!DLFecha) = Ano_Ret) Then
'        If Primera Then
'            Primera = False
'            StrSql = "UPDATE rep19 SET "
'            StrSql = StrSql & "total_entidad1 = 0"
'            StrSql = StrSql & " WHERE bpronro=" & Bpronro
'            StrSql = StrSql & " AND ternro=" & Tercero
'            Flog.writeline " Update SQL " & StrSql
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'
'        If posicionOtraObraSocial <> 0 And posicionOtraObraSocial <= 2 Then
'            Solo_Total = False
'            J = J + 1
'            StrSql = "UPDATE rep19 SET "
'            Select Case posicionOtraObraSocial
'            Case 1:
'                StrSql = StrSql & " monto_entidad1 = monto_entidad1 + " & rs_Desliq!dlmonto
'            Case 2:
'                StrSql = StrSql & " monto_entidad2 = monto_entidad2 + " & rs_Desliq!dlmonto
'            End Select
'        End If
'        If Solo_Total Then
'            StrSql = " UPDATE rep19 SET total_entidad1 = total_entidad1 + " & rs_Desliq!dlmonto
'        Else
'            StrSql = StrSql & " , total_entidad1 = total_entidad1 + " & rs_Desliq!dlmonto
'        End If
'        StrSql = StrSql & " WHERE bpronro=" & Bpronro
'        StrSql = StrSql & " AND ternro=" & Tercero
'        Flog.writeline " Update SQL " & StrSql
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
'    rs_Desliq.MoveNext
'Loop

   
Flog.writeline "********************** ITEM 8 Seguro de Vida ************************************"
Flog.writeline "************************ PRIMAS DE SEGURO ***************************************"
'********************** ITEM 8 Seguro de Vida ************************************
'************************ PRIMAS DE SEGURO ***************************************
I = 2
J = 2
Itenro = 8
Primera = True

'FGZ - 23/03/2006 - agregué esto para poder chequear el tope
StrSql = "SELECT * FROM  rep19 "
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
If rs_Rep19.State = adStateOpen Then rs_Rep19.Close
OpenRecordset StrSql, rs_Rep19


Flog.writeline "Busco la declaracion jurada"
'Busco la declaracion jurada
StrSql = "SELECT * FROM desmen where desmen.itenro = " & Itenro
StrSql = StrSql & " AND desmen.empleado =" & Tercero
OpenRecordset StrSql, rs_Desmen
If rs_Desmen.EOF Then
    Flog.writeline " No hay "
End If

Do While Not rs_Desmen.EOF
    Actualiza = False
    If Year(rs_Desmen!desfecdes) = Ano_Ret Then
        If rs_Desmen!desmondec <> 0 Then
            Actualiza = True
            If rs_Rep19!Total_entidad2 <= Items_TOPE(Itenro) Then
                StrSql = "UPDATE rep19 SET "
                StrSql = StrSql & "total_entidad2 = total_entidad2 +" & rs_Desmen!desmondec
            Else
                StrSql = "UPDATE rep19 SET "
                StrSql = StrSql & "total_entidad2 = total_entidad2 +" & Items_TOPE(Itenro)
            End If
'FGZ - 23/03/2006 - Saqué esto por lo de arriba
'            StrSql = "UPDATE rep19 SET "
'            StrSql = StrSql & "total_entidad2 = total_entidad2 +" & rs_Desmen!desmondec

        End If
        If (I < 3) And (rs_Desmen!desmondec <> 0) Then   'solo puedo poner dos en la DDJJ
            If Not Actualiza Then
                Actualiza = True
                StrSql = "UPDATE rep19 SET "
            End If
            I = I + 1
            
            Primera = False
            Select Case I
            Case 3:
                'StrSql = StrSql & ",monto_entidad3 = " & rs_Desmen!desmondec
                'FGZ - 23/03/2006
                StrSql = StrSql & ",monto_entidad3 = " & IIf(Abs(rs_Desmen!desmondec) <= Items_TOPE(Itenro), rs_Desmen!desmondec, Items_TOPE(Itenro))
                StrSql = StrSql & ",cuit_entidad3 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad3 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
            Case Else: 'los demas no se pueden dar
            End Select

        End If
        If Actualiza Then
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    rs_Desmen.MoveNext
Loop

I = I + 1
Flog.writeline " Busco las liquidaciones anteriores "
'Busco las liquidaciones anteriores
StrSql = "SELECT * FROM desliq where desliq.itenro = " & Itenro
StrSql = StrSql & " AND desliq.empleado =" & Tercero
OpenRecordset StrSql, rs_Desliq
If rs_Desliq.EOF Then
    Flog.writeline " No hay "
End If
            
Do While Not rs_Desliq.EOF
    Solo_Total = True
    If (Month(rs_Desliq!DLFecha) <= Mes_Ret) And (Year(rs_Desliq!DLFecha) = Ano_Ret) Then
        If Primera Then
            Primera = False
            StrSql = "UPDATE rep19 SET "
            StrSql = StrSql & "total_entidad2 = 0"
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If I <= 3 Then
            Solo_Total = False
            J = J + 1
            StrSql = "UPDATE rep19 SET "
            Select Case I
            Case 3:
                'StrSql = StrSql & " monto_entidad3 = monto_entidad3 + " & rs_Desliq!dlmonto
                'FGZ - 23/03/2006
                'StrSql = StrSql & " monto_entidad3 = monto_entidad3 + " & IIf(Abs(rs_Desliq!dlmonto) <= Items_TOPE(Itenro), rs_Desliq!dlmonto, Items_TOPE(Itenro))
                 StrSql = StrSql & " monto_entidad3 = " & Items_TOPE(Itenro)
            End Select
        End If
        If Solo_Total Then
            'StrSql = " UPDATE rep19 SET total_entidad2 = total_entidad2 + " & rs_Desliq!dlmonto
            'FGZ -23/03/2006
            If rs_Rep19!Total_entidad2 <= Items_TOPE(Itenro) Then
                'StrSql = StrSql & " UPDATE rep19 SET total_entidad2 = total_entidad2 + " & rs_Desliq!dlmonto
                StrSql = " UPDATE rep19 SET total_entidad2 = total_entidad2 + " & rs_Desliq!dlmonto
            Else
                StrSql = " UPDATE rep19 SET total_entidad2 = total_entidad2 + " & Items_TOPE(Itenro)
            End If
        Else
            'StrSql = StrSql & ", total_entidad2 = total_entidad2 + " & rs_Desliq!dlmonto
            'FGZ - 23/03/2006
            'If rs_Rep19!Total_entidad2 <= Items_TOPE(Itenro) Then
            '    StrSql = StrSql & ", total_entidad2 = total_entidad2 + " & rs_Desliq!dlmonto
            'Else
            '    StrSql = StrSql & ", total_entidad2 = total_entidad2 + " & Items_TOPE(Itenro)
            'End If
            StrSql = StrSql & ", total_entidad2 = " & Items_TOPE(Itenro)
        End If

        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        Flog.writeline " Update SQL " & StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_Desliq.MoveNext
Loop

Flog.writeline " Busco las DDJJ en 0 para tomar el nombre y el Cuit "
'Busco las DDJJ en 0 para tomar el nombre y el Cuit
If J > 2 Then
    'Busco la declaracion jurada
    StrSql = "SELECT * FROM desmen where desmen.itenro = " & Itenro
    StrSql = StrSql & " AND desmen.empleado =" & Tercero
    StrSql = StrSql & " AND desmen.desmondec = 0"
    OpenRecordset StrSql, rs_Desmen
    
    Do While Not rs_Desmen.EOF
        StrSql = "UPDATE rep19 SET "
        Select Case I
        Case 3:
            StrSql = StrSql & "cuit_entidad3 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
            StrSql = StrSql & ",entidad3 = '" & Format_Str(rs_Desmen!desrazsoc, 70, False, " ") & "'"
        End Select
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        Flog.writeline " Update SQL " & StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
            
        rs_Desmen.MoveNext
    Loop
End If

Flog.writeline "*********************** ITEM 9 Sepelio **************************************"
Flog.writeline "********************** GASTOS DE SEPELIO ************************************"
'*********************** ITEM 9 Sepelio **************************************
'********************** GASTOS DE SEPELIO ************************************
I = 3
J = 3
Primera = True
Itenro = 9

StrSql = "SELECT * FROM  rep19 "
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
If rs_Rep19.State = adStateOpen Then rs_Rep19.Close
OpenRecordset StrSql, rs_Rep19

Flog.writeline " Busco la declaracion jurada"
'Busco la declaracion jurada
StrSql = "SELECT * FROM desmen where desmen.itenro = " & Itenro
StrSql = StrSql & " AND desmen.empleado =" & Tercero
OpenRecordset StrSql, rs_Desmen
If rs_Desmen.EOF Then
    Flog.writeline " No hay "
End If

Do While Not rs_Desmen.EOF
    Actualiza = False
    If Year(rs_Desmen!desfecdes) = Ano_Ret Then
        If rs_Desmen!desmondec <> 0 Then
            Actualiza = True
            If rs_Rep19!Total_entidad3 < Items_TOPE(3) Then
                StrSql = "UPDATE rep19 SET "
                StrSql = StrSql & "total_entidad3 = total_entidad3 +" & rs_Desmen!desmondec
            Else
                StrSql = "UPDATE rep19 SET "
                StrSql = StrSql & "total_entidad3 = total_entidad3 +" & Items_TOPE(3)
            End If
        End If
        If (I < 5) And (rs_Desmen!desmondec <> 0) Then   'solo puedo poner dos en la DDJJ
            If Not Actualiza Then
                Actualiza = True
                StrSql = "UPDATE rep19 SET "
            End If
            I = I + 1
            
            Primera = False
            Select Case I
            Case 4:
                StrSql = StrSql & ",monto_entidad4 = " & rs_Desmen!desmondec
                StrSql = StrSql & ",cuit_entidad4 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad4 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
            Case 5:
                StrSql = StrSql & ",monto_entidad5 = " & rs_Desmen!desmondec
                StrSql = StrSql & ",cuit_entidad5 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad5 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
            Case Else: 'los demas no se pueden dar
            End Select
        End If
        If Actualiza Then
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    rs_Desmen.MoveNext
Loop

I = I + 1

'Revisar esto porque no me gusta

'Busco las liquidaciones anteriores
StrSql = "SELECT * FROM desliq where desliq.itenro = " & Itenro
StrSql = StrSql & " AND desliq.empleado =" & Tercero
OpenRecordset StrSql, rs_Desliq
If rs_Desliq.EOF Then
    Flog.writeline " No hay "
End If
            
Do While Not rs_Desliq.EOF
    Solo_Total = True
    If (Month(rs_Desliq!DLFecha) <= Mes_Ret) And (Year(rs_Desliq!DLFecha) = Ano_Ret) Then
        If Primera Then
            Primera = False
            StrSql = "UPDATE rep19 SET "
            StrSql = StrSql & "total_entidad3 = 0"
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If I <= 5 Then
            Solo_Total = False
            J = J + 1
            StrSql = "UPDATE rep19 SET "
            Select Case I
            Case 4:
                StrSql = StrSql & " monto_entidad4 = monto_entidad4 + " & rs_Desliq!dlmonto
            Case 5:
                StrSql = StrSql & " monto_entidad5 = monto_entidad5 + " & rs_Desliq!dlmonto
            End Select
        End If
        If Solo_Total Then
            StrSql = " UPDATE rep19 SET total_entidad3 = total_entidad3 + " & rs_Desliq!dlmonto
        Else
            StrSql = StrSql & ", total_entidad3 = total_entidad3 + " & rs_Desliq!dlmonto
        End If
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        Flog.writeline " Update SQL " & StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_Desliq.MoveNext
Loop

Flog.writeline " Busco las DDJJ en 0 para tomar el nombre y el Cuit "
'Busco las DDJJ en 0 para tomar el nombre y el Cuit
If J > 3 Then
    Flog.writeline " Busco la declaracion jurada"
    'Busco la declaracion jurada
    StrSql = "SELECT * FROM desmen where desmen.itenro = " & Itenro
    StrSql = StrSql & " AND desmen.empleado =" & Tercero
    StrSql = StrSql & " AND desmen.desmondec = 0"
    OpenRecordset StrSql, rs_Desmen
    
    Do While Not rs_Desmen.EOF
        StrSql = "UPDATE rep19 SET "
        Select Case I
        Case 4:
            StrSql = StrSql & "cuit_entidad4 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
            StrSql = StrSql & ",entidad4 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
        Case 5:
            StrSql = StrSql & "cuit_entidad5 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
            StrSql = StrSql & ",entidad5 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
        End Select
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        Flog.writeline " Update SQL " & StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
            
        rs_Desmen.MoveNext
    Loop
End If

' ???????
Flog.writeline " Guardo el tope del sepelio para los diferidos"
'Guardo el tope del sepelio para los diferidos
StrSql = "UPDATE rep19 SET "
StrSql = StrSql & "mon_conyuge = " & Items_TOPE(Itenro)
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
Flog.writeline " Update SQL " & StrSql
objConn.Execute StrSql, , adExecuteNoRecords


Flog.writeline "*********************** ITEM 15 Donaciones **************************************"
Flog.writeline "*************************** DONACIONES ******************************************"
'*********************** ITEM 15 Donaciones **************************************
'*************************** DONACIONES ******************************************
I = 5
J = 5
Primera = True
Itenro = 15

StrSql = "SELECT * FROM  rep19 "
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
If rs_Rep19.State = adStateOpen Then rs_Rep19.Close
OpenRecordset StrSql, rs_Rep19

Flog.writeline " Busco la declaracion jurada"
'Busco la declaracion jurada
StrSql = "SELECT * FROM desmen where desmen.itenro = " & Itenro
StrSql = StrSql & " AND desmen.empleado =" & Tercero
OpenRecordset StrSql, rs_Desmen
If rs_Desmen.EOF Then
    Flog.writeline " No hay"
End If
Do While Not rs_Desmen.EOF
    'La suma la hago teniendo en cuenta el tope
    'Actualiza = False
    If Year(rs_Desmen!desfecdes) = Ano_Ret Then
        If rs_Desmen!desmondec <> 0 Then
            'Actualiza = True
            If rs_Rep19!Total_entidad4 < Items_TOPE(Itenro) Then
                StrSql = "UPDATE rep19 SET "
                StrSql = StrSql & "total_entidad4 = total_entidad4 + " & rs_Desmen!desmondec
            Else
                StrSql = "UPDATE rep19 SET "
                StrSql = StrSql & "total_entidad4 = total_entidad4 + " & Items_TOPE(Itenro)
            End If
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        If (I < 7) And (rs_Desmen!desmondec <> 0) Then   'solo puedo poner dos en la DDJJ
'            If Not Actualiza Then
'                Actualiza = True
                StrSql = "UPDATE rep19 SET "
'            End If
            I = I + 1
            
            Primera = False
            Select Case I
            Case 6:
                'StrSql = StrSql & " monto_entidad6 = " & rs_Desmen!desmondec
                'FGZ - 23/03/2006
                StrSql = StrSql & " monto_entidad6 = " & IIf(Abs(rs_Desmen!desmondec) <= Items_TOPE(Itenro), rs_Desmen!desmondec, Items_TOPE(Itenro))
                'FGZ - 28/03/2006
                'StrSql = StrSql & ",cuit_entidad6 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                
                'FGZ - 07/12/2006 - restauré el cuit
                'StrSql = StrSql & ",cuit_entidad6 = '" & rs_Desmen!desmondec & "'"
                StrSql = StrSql & ",cuit_entidad6 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad6 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
            Case 7:
                'StrSql = StrSql & " monto_entidad7 = " & rs_Desmen!desmondec
                'FGZ -23/03/2006
                StrSql = StrSql & " monto_entidad7 = " & IIf(Abs(rs_Desmen!desmondec) <= Items_TOPE(Itenro), rs_Desmen!desmondec, Items_TOPE(Itenro))
                'FGZ - 28/03/2006
                'StrSql = StrSql & ",cuit_entidad7 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                'FGZ - 07/12/2006 - restauré el cuit
                'StrSql = StrSql & ",cuit_entidad7 = '" & rs_Desmen!desmondec & "'"
                StrSql = StrSql & ",cuit_entidad7 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
                StrSql = StrSql & ",entidad7 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
            Case Else: 'los demas no se pueden dar
            End Select

'        End If
'        If Actualiza Then
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    rs_Desmen.MoveNext
Loop

I = I + 1

'Revisar esto porque no me gusta

'Busco las liquidaciones anteriores
StrSql = "SELECT * FROM desliq where desliq.itenro = " & Itenro
StrSql = StrSql & " AND desliq.empleado =" & Tercero
OpenRecordset StrSql, rs_Desliq
If rs_Desliq.EOF Then
    Flog.writeline " No hay"
End If
            
Do While Not rs_Desliq.EOF
    Solo_Total = True
    If (Month(rs_Desliq!DLFecha) <= Mes_Ret) And (Year(rs_Desliq!DLFecha) = Ano_Ret) Then
        If Primera Then
            Primera = False
            StrSql = "UPDATE rep19 SET "
            StrSql = StrSql & "total_entidad4 = 0"
            StrSql = StrSql & " WHERE bpronro=" & Bpronro
            StrSql = StrSql & " AND ternro=" & Tercero
            Flog.writeline " Update SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If I <= 7 Then
            Solo_Total = False
            J = J + 1
            StrSql = "UPDATE rep19 SET "
            Select Case I
            Case 6:
                StrSql = StrSql & " monto_entidad6 = monto_entidad6 + " & rs_Desliq!dlmonto
            Case 7:
                StrSql = StrSql & " monto_entidad7 = monto_entidad7 + " & rs_Desliq!dlmonto
            End Select
        End If
        If Solo_Total Then
            StrSql = "UPDATE rep19 SET total_entidad4 = total_entidad4 + " & rs_Desliq!dlmonto
        Else
            StrSql = StrSql & ", total_entidad4 = total_entidad4 + " & rs_Desliq!dlmonto
        End If
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        Flog.writeline " Update SQL " & StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_Desliq.MoveNext
Loop

Flog.writeline " Busco las DDJJ en 0 para tomar el nombre y el Cuit"
'Busco las DDJJ en 0 para tomar el nombre y el Cuit
If J > 5 Then
    'Busco la declaracion jurada
    StrSql = "SELECT * FROM desmen where desmen.itenro = " & Itenro
    StrSql = StrSql & " AND desmen.empleado =" & Tercero
    StrSql = StrSql & " AND desmen.desmondec = 0"
    OpenRecordset StrSql, rs_Desmen
    
    Do While Not rs_Desmen.EOF
        StrSql = "UPDATE rep19 SET "
        Select Case I
        Case 6:
            StrSql = StrSql & "cuit_entidad6 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
            StrSql = StrSql & ",entidad6 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
        Case 7:
            StrSql = StrSql & "cuit_entidad7 = '" & Format_Str(rs_Desmen!descuit, 13, False, " ") & "'"
            StrSql = StrSql & ",entidad7 = '" & Format_Str(rs_Desmen!desrazsoc, 40, False, " ") & "'"
        End Select
        StrSql = StrSql & " WHERE bpronro=" & Bpronro
        StrSql = StrSql & " AND ternro=" & Tercero
        Flog.writeline " Update SQL " & StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
            
        rs_Desmen.MoveNext
    Loop
End If

' ???????
Flog.writeline " Guardo el tope del sepelio para los diferidos "
'Guardo el tope del sepelio para los diferidos
StrSql = "UPDATE rep19 SET "
StrSql = StrSql & "mon_hijo = " & Items_TOPE(Itenro)
Flog.writeline " Update SQL " & StrSql
objConn.Execute StrSql, , adExecuteNoRecords


Flog.writeline "*********************** ITEM 14 Aportes Voluntarios Jubilación *******************"
Flog.writeline "******************************* OTRAS DEDUCCIONES ********************************"
'*********************** ITEM 14 Aportes Voluntarios Jubilaci¢n *******************
'******************************* OTRAS DEDUCCIONES ********************************
I = 7
J = 7
Itenro = 14
Primera = True
       
StrSql = "UPDATE rep19 SET "
StrSql = StrSql & "total_entidad5 = 0"
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
Flog.writeline " Update SQL " & StrSql
objConn.Execute StrSql, , adExecuteNoRecords

'27/02/2009 - Martin Ferraro - Se agrego que tenga en cuenta el item 7
StrSql = "SELECT * FROM item WHERE item.itenro = 14 OR item.itenro = 7 OR item.itenro >= 17"
OpenRecordset StrSql, rs_Items

Do While Not rs_Items.EOF
    Flog.writeline " Busco la declaracion jurada Item " & rs_Items!Itenro
    'Busco la declaracion jurada
    StrSql = " SELECT desmen.*,item.itenom FROM desmen INNER JOIN item ON item.itenro = desmen.itenro "
    StrSql = StrSql & " WHERE desmen.itenro = " & rs_Items!Itenro
    StrSql = StrSql & " AND desmen.empleado =" & Tercero
    OpenRecordset StrSql, rs_Desmen
    If rs_Desmen.EOF Then
        Flog.writeline " No hay items " & rs_Items!Itenro
    End If
    
    Totalitem = 0
    
    Do While Not rs_Desmen.EOF
        'La suma la hago teniendo en cuenta el tope
        Actualiza = False
        If Year(rs_Desmen!desfecdes) = Ano_Ret Then
            If Items_TOPE(rs_Items!Itenro) <> 0 Then
                Totalitem = Totalitem + Abs(Items_TOPE(rs_Items!Itenro))
            End If
            If (I < 10) And (Items_TOPE(rs_Items!Itenro) <> 0) Then   'solo puedo poner dos en la DDJJ
                If Not Actualiza Then
                    Actualiza = True
                    StrSql = "UPDATE rep19 SET "
                End If
                I = I + 1
                
                Primera = False
                Select Case I
'                Case 8:
'                    StrSql = StrSql & " monto_entidad8 = " & Items_TOPE(rs_Items!Itenro)
'                    StrSql = StrSql & ",entidad8 = '" & Format_Str(rs_Desmen!itenom & "(" & rs_Desmen!desrazsoc & ")", 40, False, " ") & "'"
'                    'StrSql = StrSql & ",cuit_entidad8 = '" & Format_Str(rs_Desmen!Descuit, 13, False, " ") & "'"
'                    StrSql = StrSql & ",cuit_entidad8 = '" & FormatNumber(rs_Desmen!desmondec, 2) & "'"
'                Case 9:
'                    StrSql = StrSql & " monto_entidad9 = " & Items_TOPE(rs_Items!Itenro)
'                    StrSql = StrSql & ",entidad9 = '" & Format_Str(rs_Desmen!itenom & "(" & rs_Desmen!desrazsoc & ")", 40, False, " ") & "'"
'                    'StrSql = StrSql & ",cuit_entidad9 = '" & Format_Str(rs_Desmen!Descuit, 13, False, " ") & "'"
'                    StrSql = StrSql & ",cuit_entidad9 = '" & FormatNumber(rs_Desmen!desmondec, 2) & "'"


'FGZ - cambié el caso 8 y 9 porque el liq mete el honorariomedico en el 9 y me sale duplicado
'27/02/2009 - Martin Ferraro - Cuando case 8 insertaba en 9 entonces si habia una sola deduccion salia en el segundo reglon
                Case 8:
                    StrSql = StrSql & " monto_entidad8 = " & Items_TOPE(rs_Items!Itenro)
                    StrSql = StrSql & ",entidad8 = '" & Format_Str(rs_Desmen!itenom & "(" & rs_Desmen!desrazsoc & ")", 70, False, " ") & "'"
                    'StrSql = StrSql & ",cuit_entidad9 = '" & Format_Str(rs_Desmen!Descuit, 13, False, " ") & "'"
                    '25/02/2008 - Martin Ferraro - Valor Absoluto
                    StrSql = StrSql & ",cuit_entidad8 = '" & Abs(rs_Desmen!desmondec) & "'"
                Case 9:
                    StrSql = StrSql & " monto_entidad9 = " & Items_TOPE(rs_Items!Itenro)
                    StrSql = StrSql & ",entidad9 = '" & Format_Str(rs_Desmen!itenom & "(" & rs_Desmen!desrazsoc & ")", 70, False, " ") & "'"
                    'StrSql = StrSql & ",cuit_entidad8 = '" & Format_Str(rs_Desmen!Descuit, 13, False, " ") & "'"
                    '25/02/2008 - Martin Ferraro - Valor Absoluto
                    StrSql = StrSql & ",cuit_entidad9 = '" & Abs(rs_Desmen!desmondec) & "'"
                Case 10:
                    StrSql = StrSql & " monto_entidad10 = " & Items_TOPE(rs_Items!Itenro)
                    StrSql = StrSql & ",entidad10 = '" & Format_Str(rs_Desmen!itenom & "(" & rs_Desmen!desrazsoc & ")", 70, False, " ") & "'"
                    'StrSql = StrSql & ",cuit_entidad10 = '" & Format_Str(rs_Desmen!Descuit, 13, False, " ") & "'"
                    '25/02/2008 - Martin Ferraro - Valor Absoluto
                    StrSql = StrSql & ",cuit_entidad10 = '" & Abs(rs_Desmen!desmondec) & "'"
                Case Else: 'los demas no se pueden dar
                End Select
            Else
                Flog.writeline " i >= 10 O Items_TOPE(item) = 0 "
            End If
            If Actualiza Then
                StrSql = StrSql & " WHERE bpronro=" & Bpronro
                StrSql = StrSql & " AND ternro=" & Tercero
                Flog.writeline " Update SQL " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            Flog.writeline " desfecdes " & rs_Desmen!desfecdes & "<> año de Retencion " & Ano_Ret
        End If
        rs_Desmen.MoveNext
    Loop
    
    StrSql = "UPDATE rep19 SET "
    StrSql = StrSql & "total_entidad5 = total_entidad5 + " & Totalitem
    StrSql = StrSql & " WHERE bpronro=" & Bpronro
    StrSql = StrSql & " AND ternro=" & Tercero
    Flog.writeline " Update SQL " & StrSql
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rs_Items.MoveNext
Loop

Flog.writeline " Fin SueGan21 " & StrSql

'faltaria liberar todo y cerra
If rs_Rep19.State = adStateOpen Then rs_Rep19.Close
If rs_OS.State = adStateOpen Then rs_OS.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Desliq.State = adStateOpen Then rs_Desliq.Close
If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
If rs_Items.State = adStateOpen Then rs_Items.Close
If rs_Ter_doc.State = adStateOpen Then rs_Ter_doc.Close

Set rs_Rep19 = Nothing
Set rs_OS = Nothing
Set rs_Tercero = Nothing
Set rs_Desliq = Nothing
Set rs_Desmen = Nothing
Set rs_Items = Nothing
Set rs_Ter_doc = Nothing

Exit Sub

CE:
    'Resume Next
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub


Public Sub Suegan21(ByVal Bpronro As Long, ByVal Tercero As Long, ByVal Mes_Ret As Long, ByVal Ano_Ret As Long, ByVal Empresa As Long, ByVal Gan_imponible As Single, ByVal Por_Deduccion As Single, ByVal Aux_Fecha_A_Utilizar As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Completa los datos en la tabla reporte19, necesarios para la impresion de la
'               parte posterior del fromulario
' Autor      : Javie Iraztorza
' Fecha      : 09/05/2000
' Modif      : H.J.I 16/03/2001
' Traduccion : FGZ
' Fecha      : 22/04/2004
' Ult. Mod   : FGZ - 15/03/2007 Problemas con los items 6 y 13
' --------------------------------------------------------------------------------------------
Dim I          As Integer
Dim J          As Integer


Dim rs_Rep19 As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_DesmenAux As New ADODB.Recordset


Dim Cuit1 As String
Dim Entidad1 As String
Dim Monto1 As Double
Dim Cuit2 As String
Dim Entidad2 As String
Dim Entidad3 As String
Dim Monto2 As Double
Dim Monto3 As Double
Dim TotalEnt1 As Double
Dim TotalEnt2 As Double
Dim TotalEnt3 As Double
Dim TotalEnt4 As Double
Dim ValorDDJJ As Double
Dim MaxValorDDJJ As Double
Dim CantMayor As Integer
Dim TopeItemActual As Double
Dim SumaDesmen As Double
Dim ItemActual As Integer
Dim ItemDesc As String
Dim EntDesc As String
Dim CUITDesc As String

Flog.writeline "SueGan21 -------------------------------------------------------------------"

On Error GoTo CE

StrSql = "SELECT * FROM  rep19"
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
OpenRecordset StrSql, rs_Rep19

Flog.writeline ""

Flog.writeline "ITEM 6 y 13 "
Flog.writeline "CUOTAS MEDICO ASISTENCIALES"

Cuit1 = " "
Entidad1 = " "
Monto1 = 0
Cuit2 = " "
Entidad2 = " "
Monto2 = 0
TotalEnt1 = 0
ValorDDJJ = 0
MaxValorDDJJ = 0


'----------------------------------------------------------------------------------
'Busqueda del item 6 para el primer reglon
'----------------------------------------------------------------------------------

'---------------------mdf inicio se comenta el item 6 ya q no se deduse mas en el rubro 11

'ItemActual = 6
'Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
'If Items_TOPE(ItemActual) <> 0 Then
        
'    Monto1 = Abs(Items_TOPE(ItemActual))
    
    'Busca la obra social para la entidad y cuit 1
    
'    StrSql = "SELECT estructura.estrdabr, ter_doc.nrodoc FROM his_estructura "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
'    StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = his_estructura.ternro AND ter_doc.tidnro = 6"
'    StrSql = StrSql & " WHERE his_estructura.ternro = " & Tercero & " AND"
'    StrSql = StrSql & " his_estructura.tenro = 17 AND "
'    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(rs_Rep19!Hasta) & ") AND"
'    StrSql = StrSql & " ((" & ConvFecha(rs_Rep19!Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
'    OpenRecordset StrSql, rs_DesmenAux
'    If Not rs_DesmenAux.EOF Then
'        Cuit1 = Format_Str(rs_DesmenAux!NroDoc, 13, False, " ")
'        Entidad1 = Format_Str(rs_DesmenAux!estrdabr, 70, False, " ")
'        Flog.writeline Espacios(3) & "Se encontro Obra social = " & rs_DesmenAux!estrdabr
'    Else
'        Flog.writeline Espacios(3) & "No se encontro Obra social"
'    End If
'    rs_DesmenAux.Close

'End If


'----------------------------------------------------------------------------------
'Busqueda del item 13 para el segundo reglon
'----------------------------------------------------------------------------------
ItemActual = 13
Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
If Items_TOPE(ItemActual) <> 0 Then
        
    'Monto2 = Abs(Items_TOPE(ItemActual)) mdf
    Monto1 = Abs(Items_TOPE(ItemActual)) 'mdf
    'Miro cuantas DDJJ hay
    Flog.writeline Espacios(3) & "Buscando DDJJ para el item en el año"
    StrSql = "SELECT * FROM desmen "
    StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
    StrSql = StrSql & " AND desmen.empleado = " & Tercero
    StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
    StrSql = StrSql & " ORDER BY desmondec DESC "
    OpenRecordset StrSql, rs_Desmen
    
    'Hay DDJJ
    If rs_Desmen.RecordCount > 0 Then
            
        'Hay Una sola, tomo la entidad y cuit de esa
        If rs_Desmen.RecordCount = 1 Then
        
            Flog.writeline Espacios(4) & "Se encontro 1 DDJJ"
            'Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ") mdf
             Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
            'Entidad2 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ") mdf
             Entidad1 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ") 'mdf
             Cuit2 = " "
             Entidad2 = ""
        Else 'Hay mas de una DDJJ
        
            Flog.writeline Espacios(4) & "Se encontraron mas de una DDJJ. Busca cuantas mayor que tope y mayor"
            
            'Veo cuantas DDJJ hay mayor que el tope y cual es la mayor
            ValorDDJJ = 0
            MaxValorDDJJ = 0
            CantMayor = 0
            Do While Not rs_Desmen.EOF
                ValorDDJJ = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                ValorDDJJ = Abs(ValorDDJJ)
                
                If ValorDDJJ >= Monto1 Then 'mdf
                'If ValorDDJJ >= Monto2 Then mdf
                    CantMayor = CantMayor + 1
                    If ValorDDJJ > MaxValorDDJJ Then MaxValorDDJJ = ValorDDJJ
                End If
                
                rs_Desmen.MoveNext
            Loop
            
            'De acuerdo a la cantidad de registros pongo el nombre a entidad y cuit
            
            If CantMayor >= 1 Then 'hay mayores
                rs_Desmen.MoveFirst
                Do While Not rs_Desmen.EOF
                       If rs_Desmen!desmondec = MaxValorDDJJ Then
                           Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ") 'mdf
                           Entidad1 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ") 'mdf
                       End If
                     rs_Desmen.MoveNext
                 Loop
                 
                If rs_Desmen.RecordCount = 2 Then 'aca va la descripcion de la segunda
                  rs_Desmen.MoveFirst
                  Do While Not rs_Desmen.EOF
                     If rs_Desmen!desmondec <> MaxValorDDJJ Then
                       Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                       Entidad2 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ")
                     End If
                     rs_Desmen.MoveNext
                  Loop
                End If
                
                If rs_Desmen.RecordCount > 2 Then
                  Cuit2 = " "
                  Entidad2 = "Otras cuotas medicas"
                End If
                
            Else
                '08/10/2014 - Carmen Quintero
                'caso cuando los montos son menores que lo calculado por el proceso
                rs_Desmen.MoveFirst
                Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                Entidad1 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ")
                Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                
                If rs_Desmen.RecordCount > 2 Then
                    Cuit2 = " "
                    Entidad2 = "Cuotas Médicas Asistenciales"
                    Monto2 = CDbl(Abs(Items_TOPE(ItemActual))) - CDbl(Monto1)
                End If
                
                If rs_Desmen.RecordCount = 2 Then
                    rs_Desmen.MoveNext
                    Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad2 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ")
                    Monto2 = CDbl(Abs(Items_TOPE(ItemActual))) - CDbl(Monto1)
                End If
                'fin
            End If
        
        End If
        
    'No hay ninguna DDJJ
    Else
        Cuit1 = " "
        Entidad1 = "Adic. Obra Social"
        Cuit2 = " "
        Entidad2 = " "
    End If
    
    rs_Desmen.Close
    
End If

'----------------------------------------------------------------------------------
'Se analiza que poner en cada reglon
'----------------------------------------------------------------------------------
'TotalEnt1 = Monto1 + Monto2
TotalEnt1 = CDbl(Abs(Items_TOPE(ItemActual)))
If TotalEnt1 > 0 Then

    'If (Monto1 > 0) And (Monto2 > 0) Then
        'Tengo los 2 reglones completos
        StrSql = "UPDATE rep19 SET"
        StrSql = StrSql & " monto_entidad1 = " & Monto1
        StrSql = StrSql & ",cuit_entidad1 = '" & Cuit1 & "'"
        StrSql = StrSql & ",entidad1 = '" & Entidad1 & "'"
        StrSql = StrSql & ",monto_entidad2 = " & Monto2
        StrSql = StrSql & ",cuit_entidad2 = '" & Cuit2 & "'"
        StrSql = StrSql & ",entidad2 = '" & Entidad2 & "'"
        StrSql = StrSql & ",total_entidad1 = " & TotalEnt1
        StrSql = StrSql & " WHERE bpronro = " & Bpronro
        StrSql = StrSql & " AND ternro = " & Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
    'Else
    '    If (Monto1 > 0) Then
    '        'Tengo solo entidad 1, la pongo en el primer reglon
    '        StrSql = "UPDATE rep19 SET"
    '        StrSql = StrSql & " monto_entidad1 = " & Monto1
    '        StrSql = StrSql & ",cuit_entidad1 = '" & Cuit1 & "'"
    '        StrSql = StrSql & ",entidad1 = '" & Entidad1 & "'"
    '        StrSql = StrSql & ",total_entidad1 = " & TotalEnt1
    '        StrSql = StrSql & " WHERE bpronro = " & Bpronro
    '        StrSql = StrSql & " AND ternro = " & Tercero
    '        objConn.Execute StrSql, , adExecuteNoRecords
    '    Else
            'Tengo solo entidad 2, la pongo en el primer reglon
    '        StrSql = "UPDATE rep19 SET"
    '        StrSql = StrSql & " monto_entidad1 = " & Monto2
    '        StrSql = StrSql & ",cuit_entidad1 = '" & Cuit2 & "'"
    '        StrSql = StrSql & ",entidad1 = '" & Entidad2 & "'"
    '        StrSql = StrSql & ",total_entidad1 = " & TotalEnt1
    '        StrSql = StrSql & " WHERE bpronro = " & Bpronro
    '        StrSql = StrSql & " AND ternro = " & Tercero
    '        objConn.Execute StrSql, , adExecuteNoRecords
    '    End If
    'End If
Else
    'No hay reglon 1 ni 2
    StrSql = "UPDATE rep19 SET "
    StrSql = StrSql & "cuit_entidad1 =' '"
    StrSql = StrSql & ",entidad1 =' '"
    StrSql = StrSql & ",monto_entidad1 = 0 "
    StrSql = StrSql & ",total_entidad1 = 0 "
    StrSql = StrSql & ",cuit_entidad2 =' '"
    StrSql = StrSql & ",entidad2 =' '"
    StrSql = StrSql & ",monto_entidad2 = 0 "
    StrSql = StrSql & " WHERE bpronro=" & Bpronro
    StrSql = StrSql & " AND ternro=" & Tercero
    objConn.Execute StrSql, , adExecuteNoRecords
End If
   
   
Flog.writeline "ITEM 8 Seguro de Vida "
Flog.writeline "PRIMAS DE SEGURO "


Cuit1 = ""
Entidad1 = ""
Monto1 = 0
ValorDDJJ = 0
MaxValorDDJJ = 0
ItemActual = 8
'----------------------------------------------------------------------------------
'Busqueda del item 8 para el unico reglon
'----------------------------------------------------------------------------------
Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
If Items_TOPE(ItemActual) <> 0 Then

    Monto1 = Abs(Items_TOPE(ItemActual))
    
    'Miro cuantas DDJJ hay
    Flog.writeline Espacios(3) & "Buscando DDJJ para el item en el año"
    StrSql = "SELECT * FROM desmen "
    StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
    StrSql = StrSql & " AND desmen.empleado = " & Tercero
    StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
    OpenRecordset StrSql, rs_Desmen
    
    'Hay DDJJ
    If rs_Desmen.RecordCount > 0 Then
            
        'Hay Una sola, tomo la entidad y cuit de esa
        If rs_Desmen.RecordCount = 1 Then
            
            Flog.writeline Espacios(4) & "Se encontro 1 DDJJ"
            Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
            Entidad1 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ")
        
        Else 'Hay mas de una DDJJ
            
            'Veo cuantas DDJJ hay mayor que el tope y cual es la mayor
            Flog.writeline Espacios(4) & "Se encontraron mas de una DDJJ. Busca cuantas mayor que tope y mayor"
            ValorDDJJ = 0
            MaxValorDDJJ = 0
            CantMayor = 0
            Do While Not rs_Desmen.EOF
                ValorDDJJ = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                ValorDDJJ = Abs(ValorDDJJ)
                
                If ValorDDJJ >= Monto1 Then
                    CantMayor = CantMayor + 1
                    If ValorDDJJ > MaxValorDDJJ Then MaxValorDDJJ = ValorDDJJ
                End If
                
                rs_Desmen.MoveNext
            Loop
            
            'De acuerdo a la cantidad de registros pongo el nombre a entidad y cuit
            If CantMayor = 1 Then
            
                Flog.writeline Espacios(4) & "1 sola mayor al tope. Se muestra la misma"
                'Busco la de descripcion y cuit de la mayor
                rs_Desmen.MoveFirst
                Do While Not rs_Desmen.EOF
                    If rs_Desmen!desmondec = MaxValorDDJJ Then
                        Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                        Entidad1 = Format_Str(rs_Desmen!desrazsoc, 70, False, " ")
                    End If
                    
                    rs_Desmen.MoveNext
                Loop
                
            Else
                Flog.writeline Espacios(4) & "Varias Mayor al tope"
                Cuit1 = " "
                Entidad1 = "Seguro de Vida"
            End If
            
        End If
        
    'No hay ninguna DDJJ
    Else
        Flog.writeline Espacios(4) & "No se encontraron DDJJ"
        Cuit1 = " "
        Entidad1 = "Seguro de Vida"
    End If
    
    rs_Desmen.Close
    
End If

'----------------------------------------------------------------------------------
'Inserto en el unico reglon
'----------------------------------------------------------------------------------
If Monto1 > 0 Then
    'Tengo solo entidad 1, la pongo en el primer reglon
    StrSql = "UPDATE rep19 SET"
    StrSql = StrSql & " monto_entidad3 = " & Monto1
    StrSql = StrSql & ",cuit_entidad3 = '" & Cuit1 & "'"
    StrSql = StrSql & ",entidad3 = '" & Entidad1 & "'"
    StrSql = StrSql & " WHERE bpronro = " & Bpronro
    StrSql = StrSql & " AND ternro = " & Tercero
    objConn.Execute StrSql, , adExecuteNoRecords
Else
    'No hay datos
    StrSql = "UPDATE rep19 SET"
    StrSql = StrSql & " monto_entidad3 = 0"
    StrSql = StrSql & ",cuit_entidad3 = ' '"
    StrSql = StrSql & ",entidad3 = ' '"
    StrSql = StrSql & " WHERE bpronro = " & Bpronro
    StrSql = StrSql & " AND ternro = " & Tercero
    objConn.Execute StrSql, , adExecuteNoRecords
End If

Flog.writeline "ITEM 9 Sepelio "
Flog.writeline "GASTOS DE SEPELIO "
'----------------------------------------------------------------------------------
'Busqueda del item 9 para los dos reglones
'----------------------------------------------------------------------------------
Cuit1 = " "
Entidad1 = " "
Monto1 = 0
Cuit2 = " "
Entidad2 = " "
Monto2 = 0
TotalEnt1 = 0
TotalEnt2 = 0
ItemActual = 9

Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
If Items_TOPE(ItemActual) <> 0 Then
    
    TopeItemActual = Abs(Items_TOPE(ItemActual))
    
    'Miro cuantas DDJJ hay
    Flog.writeline Espacios(3) & "Buscando DDJJ para el item en el año"
    StrSql = "SELECT * FROM desmen "
    StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
    StrSql = StrSql & " AND desmen.empleado = " & Tercero
    StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
    StrSql = StrSql & " ORDER BY ABS(desmondec) DESC"
    OpenRecordset StrSql, rs_Desmen
    
    'Hay DDJJ
    If rs_Desmen.RecordCount > 0 Then
            
        If rs_Desmen.RecordCount = 1 Then
        
            Flog.writeline Espacios(4) & "Se encontro 1 DDJJ"
            
            Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
            Monto1 = Abs(Monto1)
            Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
            Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
            TotalEnt1 = TopeItemActual
        
        Else
            
            Flog.writeline Espacios(4) & "Se encontro 2 o mas DDJJ"
                
            'Verifico si la suma supera el tope
            SumaDesmen = 0
            StrSql = "SELECT SUM(ABS(desmondec)) suma FROM desmen "
            StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
            StrSql = StrSql & " AND desmen.empleado = " & Tercero
            StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
            OpenRecordset StrSql, rs_DesmenAux
            If Not rs_DesmenAux.EOF Then
                SumaDesmen = IIf(EsNulo(rs_DesmenAux!suma), 0, rs_DesmenAux!suma)
            End If
            rs_DesmenAux.Close
            
            Flog.writeline Espacios(4) & "Suma de las DDJJ = " & SumaDesmen
            
            If SumaDesmen <= TopeItemActual Then
            
                If rs_Desmen.RecordCount = 2 Then
                
                    Flog.writeline Espacios(4) & "La suma no supera el tope y hay dos DDJJ muestro las dos"
                    
                    Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                    Monto1 = Abs(Monto1)
                    Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    TotalEnt1 = Monto1
                    
                    rs_Desmen.MoveNext
                    
                    Monto2 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                    Monto2 = Abs(Monto2)
                    Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad2 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    TotalEnt2 = Monto2
                
                Else
                
                    Flog.writeline Espacios(4) & "Hay mas de dos DDJJ y la suma no supera el tope"
                    Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                    Monto1 = Abs(Monto1)
                    Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    TotalEnt1 = Monto1
                
                    rs_Desmen.MoveNext
                
                    Monto2 = SumaDesmen - TotalEnt1
                    Cuit2 = " "
                    Entidad2 = "Otros Gastos Sepelio"
                    TotalEnt2 = Monto2
                    
                End If
            Else
            
                Flog.writeline Espacios(4) & "La suma es mayor que el tope, verifico si la primera es mayor al tope para ver si va completa"
                Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                Monto1 = Abs(Monto1)
                Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                
                If Monto1 > TopeItemActual Then
                    'Gasto todo en la primera
                    TotalEnt1 = TopeItemActual
                    
                    rs_Desmen.MoveNext
                    
                    Monto2 = SumaDesmen - Monto1
                    If rs_Desmen.RecordCount > 2 Then
                        Cuit2 = " "
                        Entidad2 = "Otros Gastos Sepelio"
                    Else
                        Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                        Entidad2 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    End If
                    TotalEnt2 = 0
                Else
                    TotalEnt1 = Monto1
                    
                    rs_Desmen.MoveNext
                    
                    Monto2 = SumaDesmen - Monto1
                    If (rs_Desmen.RecordCount > 2) Then
                        Entidad2 = "Otros Gastos Sepelio"
                        Cuit2 = " "
                    Else
                        Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                        Entidad2 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    End If
                    
                    TotalEnt2 = TopeItemActual - Monto1
                    
                End If
                
            End If
                
        End If 'If rs_Desmen.RecordCount = 1 Then
        
    'No hay ninguna DDJJ
    Else
        Cuit1 = " "
        Entidad1 = "Otros Gastos Sepelio"
        Monto1 = 0
        TotalEnt1 = TopeItemActual
    End If
    
    rs_Desmen.Close
    
End If

StrSql = "UPDATE rep19 SET"
StrSql = StrSql & " monto_entidad4 = " & Monto1
StrSql = StrSql & ",cuit_entidad4 = '" & Cuit1 & "'"
StrSql = StrSql & ",entidad4 = '" & Entidad1 & "'"
StrSql = StrSql & ",monto_entidad5 = " & Monto2
StrSql = StrSql & ",cuit_entidad5 = '" & Cuit2 & "'"
StrSql = StrSql & ",entidad5 = '" & Entidad2 & "'"
StrSql = StrSql & ",total_entidad10 = " & TotalEnt1
StrSql = StrSql & ",total_entidad11 = " & TotalEnt2
StrSql = StrSql & ",total_entidad12 = " & TotalEnt1 + TotalEnt2
StrSql = StrSql & " WHERE bpronro = " & Bpronro
StrSql = StrSql & " AND ternro = " & Tercero
objConn.Execute StrSql, , adExecuteNoRecords

Flog.writeline "ITEM 15 Donaciones "
Flog.writeline "DONACIONES "
'----------------------------------------------------------------------------------
'Busqueda del item 15 para los dos reglones
'----------------------------------------------------------------------------------
Cuit1 = " "
Entidad1 = " "
Monto1 = 0
Cuit2 = " "
Entidad2 = " "
Monto2 = 0
TotalEnt1 = 0
TotalEnt2 = 0
ItemActual = 15


Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
If Items_TOPE(ItemActual) <> 0 Then
    
    TopeItemActual = Abs(Items_TOPE(ItemActual))
    
    'Miro cuantas DDJJ hay
    Flog.writeline Espacios(3) & "Buscando DDJJ para el item en el año"
    StrSql = "SELECT * FROM desmen "
    StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
    StrSql = StrSql & " AND desmen.empleado = " & Tercero
    StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
    StrSql = StrSql & " ORDER BY ABS(desmondec) DESC"
    OpenRecordset StrSql, rs_Desmen
    
    'Hay DDJJ
    If rs_Desmen.RecordCount > 0 Then
            
        If rs_Desmen.RecordCount = 1 Then
        
            Flog.writeline Espacios(4) & "Se encontro 1 DDJJ"
            
            Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
            Monto1 = Abs(Monto1)
            Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
            Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
            TotalEnt1 = TopeItemActual
        
        Else
            
            Flog.writeline Espacios(4) & "Se encontro 2 o mas DDJJ"
                
            'Verifico si la suma supera el tope
            SumaDesmen = 0
            StrSql = "SELECT SUM(ABS(desmondec)) suma FROM desmen "
            StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
            StrSql = StrSql & " AND desmen.empleado = " & Tercero
            StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
            OpenRecordset StrSql, rs_DesmenAux
            If Not rs_DesmenAux.EOF Then
                SumaDesmen = IIf(EsNulo(rs_DesmenAux!suma), 0, rs_DesmenAux!suma)
            End If
            rs_DesmenAux.Close
            
            Flog.writeline Espacios(4) & "Suma de las DDJJ = " & SumaDesmen
            
            If SumaDesmen <= TopeItemActual Then
            
                If rs_Desmen.RecordCount = 2 Then
                
                    Flog.writeline Espacios(4) & "La suma no supera el tope y hay dos DDJJ muestro las dos"
                    
                    Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                    Monto1 = Abs(Monto1)
                    Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    TotalEnt1 = Monto1
                    
                    rs_Desmen.MoveNext
                    
                    Monto2 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                    Monto2 = Abs(Monto2)
                    Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad2 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    TotalEnt2 = Monto2
                
                Else
                
                    Flog.writeline Espacios(4) & "Hay mas de dos DDJJ y la suma no supera el tope"
                    Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                    Monto1 = Abs(Monto1)
                    Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                    Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    TotalEnt1 = Monto1
                
                    rs_Desmen.MoveNext
                
                    Monto2 = SumaDesmen - TotalEnt1
                    Cuit2 = " "
                    Entidad2 = "Otras Donaciones"
                    TotalEnt2 = Monto2
                    
                End If
            Else
            
                Flog.writeline Espacios(4) & "La suma es mayor que el tope, verifico si la primera es mayor al tope para ver si va completa"
                Monto1 = IIf(EsNulo(rs_Desmen!desmondec), 0, rs_Desmen!desmondec)
                Monto1 = Abs(Monto1)
                Cuit1 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                Entidad1 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                
                If Monto1 > TopeItemActual Then
                    'Gasto todo en la primera
                    TotalEnt1 = TopeItemActual
                    
                    rs_Desmen.MoveNext
                    
                    Monto2 = SumaDesmen - Monto1
                    If rs_Desmen.RecordCount > 2 Then
                        Cuit2 = " "
                        Entidad2 = "Otras Donaciones"
                    Else
                        Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                        Entidad2 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    End If
                    TotalEnt2 = 0
                Else
                    'ACA HAY QUE VER SI HAY DOS O NO
                    TotalEnt1 = Monto1
                    
                    rs_Desmen.MoveNext
                    
                    Monto2 = SumaDesmen - Monto1
                    If (rs_Desmen.RecordCount > 2) Then
                        Entidad2 = "Otras Donaciones"
                        Cuit2 = " "
                    Else
                        Cuit2 = Format_Str(rs_Desmen!descuit, 13, False, " ")
                        Entidad2 = Format_Str(rs_Desmen!desrazsoc, 40, False, " ")
                    End If
                    
                    TotalEnt2 = TopeItemActual - Monto1
                End If
                
            End If
                
        End If 'If rs_Desmen.RecordCount = 1 Then
        
    'No hay ninguna DDJJ
    Else
        Cuit1 = " "
        Entidad1 = "Otras Donaciones"
        Monto1 = 0
        TotalEnt1 = TopeItemActual
    End If
    
    rs_Desmen.Close
    
End If


StrSql = "UPDATE rep19 SET"
StrSql = StrSql & " monto_entidad6 = " & Monto1
StrSql = StrSql & ",cuit_entidad6 = '" & Cuit1 & "'"
StrSql = StrSql & ",entidad6 = '" & Entidad1 & "'"
StrSql = StrSql & ",monto_entidad7 = " & Monto2
StrSql = StrSql & ",cuit_entidad7 = '" & Cuit2 & "'"
StrSql = StrSql & ",entidad7 = '" & Entidad2 & "'"
StrSql = StrSql & ",total_entidad13 = " & TotalEnt1
StrSql = StrSql & ",total_entidad14 = " & TotalEnt2
StrSql = StrSql & ",total_entidad6 = " & TotalEnt1 + TotalEnt2
StrSql = StrSql & " WHERE bpronro = " & Bpronro
StrSql = StrSql & " AND ternro = " & Tercero
objConn.Execute StrSql, , adExecuteNoRecords



Flog.writeline
Flog.writeline "ITEM 14 Aportes Voluntarios Jubilación "
Flog.writeline "OTRAS DEDUCCIONES "


J = 0
TotalEnt1 = 0
Monto1 = 0
Entidad1 = " "
TotalEnt2 = 0
Monto2 = 0
Entidad2 = " "
TotalEnt3 = 0
Monto3 = 0
Entidad3 = " "

TotalEnt4 = 0


'codigo seba 19/03/2012
'ItemActual = 7
'Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
'If Items_TOPE(ItemActual) <> 0 Then
'
'    Monto1 = Abs(Items_TOPE(ItemActual))
'
'    'Busca la obra social para la entidad y cuit 1
'
'    StrSql = "SELECT estructura.estrdabr, ter_doc.nrodoc FROM his_estructura "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
'    StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = his_estructura.ternro AND ter_doc.tidnro = 6"
'    StrSql = StrSql & " WHERE his_estructura.ternro = " & Tercero & " AND"
'    StrSql = StrSql & " his_estructura.tenro = 16 AND "
'    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(rs_Rep19!Hasta) & ") AND"
'    StrSql = StrSql & " ((" & ConvFecha(rs_Rep19!Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
'    OpenRecordset StrSql, rs_DesmenAux
'    If Not rs_DesmenAux.EOF Then
'        'Cuit1 = Format_Str(rs_DesmenAux!NroDoc, 13, False, " ")
'        Entidad1 = Format_Str(rs_DesmenAux!estrdabr, 70, False, " ")
'        Flog.writeline Espacios(3) & "Se encontro Sindicato = " & rs_DesmenAux!estrdabr
'    Else
'        Flog.writeline Espacios(3) & "No se encontro Sindicato"
'    End If
'    rs_DesmenAux.Close
'End If


'StrSql = "UPDATE rep19 SET "
'StrSql = StrSql & " monto_entidad8 = " & Monto1 & ""
'StrSql = StrSql & ",entidad8 = '" & Format_Str(Entidad1, 70, False, " ") & "'"
'StrSql = StrSql & ",cuit_entidad8 = '" & Monto1 & "'"
''StrSql = StrSql & ",monto_entidad9 = " & TotalEnt2
''StrSql = StrSql & ",entidad9 = '" & Format_Str(Entidad2, 70, False, " ") & "'"
''StrSql = StrSql & ",cuit_entidad9 = '" & Monto2 & "'"
''StrSql = StrSql & ",monto_entidad10 = " & TotalEnt3
''StrSql = StrSql & ",entidad10 = '" & Format_Str(Entidad3, 70, False, " ") & "'"
''StrSql = StrSql & ",cuit_entidad10 = '" & Monto3 & "'"
''StrSql = StrSql & ",total_entidad5 = " & TotalEnt4
'StrSql = StrSql & " WHERE bpronro=" & Bpronro
'StrSql = StrSql & " AND ternro=" & Tercero
'Flog.writeline " Update SQL " & StrSql
'objConn.Execute StrSql, , adExecuteNoRecords

'hasta aca
J = 0
For I = 1 To 7
    
    Select Case I
        Case 1
            ItemActual = 7
        Case 2
            ItemActual = 20
        Case 3
            ItemActual = 21
        Case 4
            ItemActual = 31
        'FGZ - 28/06/2011 -- se agregó --------
        Case 5
            ItemActual = 22
        Case 6
            ItemActual = 24
        Case 7
            ItemActual = 23
        'FGZ - 02/10/2013 -----------------------------------------------
        'Se sacó el item 56 pues, solo va en el anverso en el rubro 9b.
        'MDZ - 27/06/2013 - CAS-19228 - se agregó
        'Case 8
        '    ItemActual = 56
        'FGZ - 02/10/2013 -----------------------------------------------
    End Select
    

'codigo seba 19/03/2012
'Flog.writeline Espacios(2) & "Items_TOPE(" & ItemActual & ") = " & Abs(Items_TOPE(ItemActual))
'If Items_TOPE(ItemActual) <> 0 Then
        
'    Monto1 = Abs(Items_TOPE(ItemActual))
    
    'Busca la obra social para la entidad y cuit 1
    
'    StrSql = "SELECT estructura.estrdabr, ter_doc.nrodoc FROM his_estructura "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
'    StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = his_estructura.ternro AND ter_doc.tidnro = 6"
'    StrSql = StrSql & " WHERE his_estructura.ternro = " & Tercero & " AND"
'    StrSql = StrSql & " his_estructura.tenro = 16 AND "
'    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(rs_Rep19!Hasta) & ") AND"
'    StrSql = StrSql & " ((" & ConvFecha(rs_Rep19!Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
'    OpenRecordset StrSql, rs_DesmenAux
'    If Not rs_DesmenAux.EOF Then
'        Cuit1 = Format_Str(rs_DesmenAux!NroDoc, 13, False, " ")
'        Entidad1 = Format_Str(rs_DesmenAux!estrdabr, 70, False, " ")
'        Flog.writeline Espacios(3) & "Se encontro Sindicato = " & rs_DesmenAux!estrdabr
'    Else
'        Flog.writeline Espacios(3) & "No se encontro Sindicato"
'    End If
'    rs_DesmenAux.Close

'End If
'hasta aca

    If Items_TOPE(ItemActual) <> 0 Then
        
        EntDesc = ""
        CUITDesc = ""
        ItemDesc = ""
        
        If ItemActual = 7 Then 'Si el item es 7 busco la descripcion del sindicato.. por estructura
            If Items_TOPE(ItemActual) <> 0 Then
                Monto1 = Abs(Items_TOPE(ItemActual))
                SumaDesmen = Monto1
                EntDesc = ""
                'Busca la obra social para la entidad y cuit 1
                StrSql = "SELECT estructura.estrdabr, ter_doc.nrodoc FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = his_estructura.ternro AND ter_doc.tidnro = 6"
                StrSql = StrSql & " WHERE his_estructura.ternro = " & Tercero & " AND"
                StrSql = StrSql & " his_estructura.tenro = 16 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(rs_Rep19!Hasta) & ") AND"
                StrSql = StrSql & " ((" & ConvFecha(rs_Rep19!Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_DesmenAux
                If Not rs_DesmenAux.EOF Then
                    Cuit1 = Format_Str(rs_DesmenAux!NroDoc, 13, False, " ")
                    If Cuit1 <> "" Then
                        EntDesc = Format_Str(rs_DesmenAux!estrdabr & "(" & Cuit1 & ")", 70, False, " ")
                    Else
                        EntDesc = Format_Str(rs_DesmenAux!estrdabr, 70, False, " ")
                    End If
                    Flog.writeline Espacios(3) & "Se encontro Sindicato = " & Entidad1
                    '_____________________________________________________________________________
                    'EntDesc = Entidad1 'IIf(EsNulo(rs_DesmenAux!desrazsoc), 0, rs_DesmenAux!desrazsoc)
                    ItemDesc = Format_Str(EntDesc, 70, False, " ")
                Else
                    Flog.writeline Espacios(3) & "No se encontro Sindicato"
                End If
                rs_DesmenAux.Close
                J = J + 1
            End If
        Else
            EntDesc = ""
            CUITDesc = ""
            StrSql = "SELECT descuit, desrazsoc"
            StrSql = StrSql & " FROM desmen"
            StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
            StrSql = StrSql & " AND desmen.empleado = " & Tercero
            StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
            OpenRecordset StrSql, rs_DesmenAux
            If (Not rs_DesmenAux.EOF) And (rs_DesmenAux.RecordCount = 1) Then
                EntDesc = IIf(EsNulo(rs_DesmenAux!desrazsoc), "", rs_DesmenAux!desrazsoc)
                CUITDesc = IIf(EsNulo(rs_DesmenAux!descuit), 0, rs_DesmenAux!descuit)
            End If
            rs_DesmenAux.Close
        
        
            SumaDesmen = 0
            StrSql = "SELECT sum(ABS(desmondec)) suma, item.itenom"
            StrSql = StrSql & " FROM desmen"
            StrSql = StrSql & " INNER JOIN item ON item.itenro = desmen.itenro"
            StrSql = StrSql & " WHERE desmen.itenro = " & ItemActual
            StrSql = StrSql & " AND desmen.empleado = " & Tercero
            StrSql = StrSql & " AND desmen.desano = " & Year(rs_Rep19!Desde)
            StrSql = StrSql & " GROUP BY item.itenom"
            OpenRecordset StrSql, rs_DesmenAux
            If Not rs_DesmenAux.EOF Then
                SumaDesmen = IIf(EsNulo(rs_DesmenAux!suma), 0, rs_DesmenAux!suma)
                ItemDesc = IIf(EsNulo(rs_DesmenAux!itenom), " ", rs_DesmenAux!itenom)
                If ((EntDesc <> "") Or (CUITDesc <> "")) Then
                    ItemDesc = ItemDesc & " (" & EntDesc & " " & CUITDesc & ")"
                End If
                
            End If
            rs_DesmenAux.Close
            J = J + 1
        End If
        'Cada linea es uno de los 3 reglones
        If SumaDesmen <> 0 Then
            Select Case J
                Case 0
                Case 1
                    TotalEnt1 = Abs(Items_TOPE(ItemActual))
                    Monto1 = SumaDesmen
                    Entidad1 = ItemDesc
                Case 2
                    TotalEnt2 = Abs(Items_TOPE(ItemActual))
                    Monto2 = SumaDesmen
                    Entidad2 = ItemDesc
                Case 3
                    TotalEnt3 = Abs(Items_TOPE(ItemActual))
                    Monto3 = SumaDesmen
                    Entidad3 = ItemDesc
                Case Else
                    TotalEnt3 = TotalEnt3 + Abs(Items_TOPE(ItemActual))
                    Monto3 = Monto3 + SumaDesmen
                    Entidad3 = "Otros"
            End Select
        End If
        TotalEnt4 = TotalEnt4 + Abs(Items_TOPE(ItemActual))
        
    End If

Next I
       
StrSql = "UPDATE rep19 SET "
StrSql = StrSql & " monto_entidad8 = " & TotalEnt1
StrSql = StrSql & ",entidad8 = '" & Format_Str(Entidad1, 70, False, " ") & "'"
StrSql = StrSql & ",cuit_entidad8 = '" & Monto1 & "'"
StrSql = StrSql & ", monto_entidad9 = " & TotalEnt2
StrSql = StrSql & ",entidad9 = '" & Format_Str(Entidad2, 70, False, " ") & "'"
StrSql = StrSql & ",cuit_entidad9 = '" & Monto2 & "'"
StrSql = StrSql & ",monto_entidad10 = " & TotalEnt3
StrSql = StrSql & ",entidad10 = '" & Format_Str(Entidad3, 70, False, " ") & "'"
StrSql = StrSql & ",cuit_entidad10 = '" & Monto3 & "'"
StrSql = StrSql & ",total_entidad5 = " & TotalEnt4
StrSql = StrSql & " WHERE bpronro=" & Bpronro
StrSql = StrSql & " AND ternro=" & Tercero
Flog.writeline " Update SQL " & StrSql
objConn.Execute StrSql, , adExecuteNoRecords


Flog.writeline " Fin SueGan21 ---------------------------------------------------------------"
Flog.writeline
'faltaria liberar todo y cerra
If rs_Rep19.State = adStateOpen Then rs_Rep19.Close
If rs_DesmenAux.State = adStateOpen Then rs_DesmenAux.Close
If rs_Desmen.State = adStateOpen Then rs_Desmen.Close

Set rs_Rep19 = Nothing
Set rs_DesmenAux = Nothing
Set rs_Desmen = Nothing

Exit Sub

CE:
    'Resume Next
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub



Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!Empleado <> Anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function

