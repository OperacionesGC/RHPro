Attribute VB_Name = "MdlImpERecruiting"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "25/03/2009" 'Martin Ferraro - Version Inicial
'-----------------------------------------------

'Const Version = "1.01"
'Const FechaVersion = "23/04/2009" 'Martin Ferraro - CtrlNuloTXT: se agrego el mid dentro de la funcion
'                                                   Importar10: emparant inicializo en 0
'                                                   CtrlNuloNUM: se agrego el contro de si es numero
'-----------------------------------------------

'Const Version = "1.02"
'Const FechaVersion = "02/06/2009" 'Martin Ferraro - Importacion desde RHPro a ER busquedas, puestos y como nos conocio
'
'
'-----------------------------------------------

'Const Version = "1.03"
'Const FechaVersion = "20/08/2009" 'Martin Ferraro - Encriptacion conexion
'
'
'-----------------------------------------------


'Const Version = "1.04"
'Const FechaVersion = "27/11/2009" 'Martin Ferraro - Pasar Postulaciones de ER a RHPro
'
'
'-----------------------------------------------


'Const Version = "1.05"
'Const FechaVersion = "28/12/2009" 'Martin Ferraro - Se cambio la politica de alcance de empleos anteriores de 15 a 18
'
'
'-----------------------------------------------
'
'Const Version = "1.06"
'Const FechaVersion = "29/12/2009" 'FGZ - Se agregó la replica de postulantes para Legajo Electronico
'
'
'-----------------------------------------------

'Const Version = "1.07"
'Const FechaVersion = "06/01/2009" 'EAM - Se agregó la replica de postulantes para Legajo Electronico
''
''
''-----------------------------------------------

'Const Version = "1.08"
'Const FechaVersion = "24/02/2011" 'EAM - Se dejó solo la parte que migra de rhpro (Busquedas) a ER y se comento las importaciones de ER a LE y RHPRO


'Const Version = "1.09"
'Const FechaVersion = "26/05/2011" 'JPB - Se reincorporaron las importaciones de ER a LE y RHPRO. También se incorporó una funcion "importar13" para importar
'                                        de e-rec a RhPro la seccion "Sykes Academy"

'Const Version = "1.10"             'Zamarbide Juan A. - CAS-11901 - Heidt y Asociados - Modificación de EyP
'Const FechaVersion = "14/07/2011"  'En el Sub Importar2 se agregó el código para que envíe el dato del textfield "Puesto al que Aspira" en el campo pospuestoaspira la tabla postulantes


'Const Version = "1.11"
'Const FechaVersion = "10/11/2011" 'JPB - Se solucionó un bug cuando verificaba, si al querer pasar los datos desde E-R a Rhpro, el postulante existía como Ampleado en Rhpro.
''                                        Se modificó la forma que verifica si se usa Leg. Electrónico. (Antes se fijaba en Conf. General y ahora se configura en el codigo 9 de la Conf. de la Empresa)
                                         
'**************************************
' ** Version No liberada
'Const Version = "1.12"
'Const FechaVersion = "27/08/2013" 'FGZ - Generacion de archivo de log con detalles de sincronizacion de tablas EyP <--> ER
'                                   En la bd de ER los domicilios no tienen tipo asignado ==>
'                                   Se agregó un parametro nuevo al confrep (COlomna 2) para configurar el modelo de domicilio donde se actualiza de ER a EyP.
'                                   Defualt tipo2 (Particular)


'Const Version = "1.13"
'Const FechaVersion = "09/10/2013" 'JPB - CAS-21010 - ANDREANI -  BUG TESTEOS E-RECRUITING (Habia error cuando importaba datos de importar2)
 
 
'Const Version = "1.14"
'Const FechaVersion = "16/10/2013" 'JPB - CAS-21010 - ANDREANI -  BUG TESTEOS E-RECRUITING
                                  'Se soluciono bug. No pasaba a RHPro el telefono Celular.
 
'Const Version = "1.15"
'Const FechaVersion = "13/02/2014" 'JPB - CAS-CAS-19564 - Raffo - Recruiting Básicas
'                                  'Se creo un procedimiento encargado de subir la imagen a la carpeta fotos y la asocia en la tabla ter_imag
'Const Version = "1.16"
'Const FechaVersion = "19/02/2014" 'JPB - CAS-CAS-19564 - Raffo - Recruiting Básicas
'                                  'Se customizo la seccion 4 de Estudios Formales

Const Version = "1.17"
Const FechaVersion = "19/02/2014" 'JPB - CAS-CAS-19564 - Raffo - Recruiting Básicas
                                  'Se constola que una fecha no venga con la forma 01/01/1900

'
'-----------------------------------------------

'-------------------------------------------------------------------------------------------------
'Conexion Externa
'-------------------------------------------------------------------------------------------------
Global ExtConn As New ADODB.Connection
Global ConnLE As New ADODB.Connection
Global Usa_LE As Boolean
Global Misma_BD As Boolean
Global FSinc
Global NroTipoDomi As Long



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Importacion.
' Autor      : Martin Ferraro
' Fecha      : 25/03/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim Nombre_Arch2 As String
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
'        If IsNumeric(strCmdLine) Then
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

    Nombre_Arch = PathFLog & "Import_E_Recruiting" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch2 = PathFLog & "Sinc_ER-EyP" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FSinc = fs.CreateTextFile(Nombre_Arch2, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 234 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ImpERecruiting(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
    
Fin:
    Flog.Close
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
End Sub


Public Sub ImpERecruiting(ByVal bpronro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de Importacion ERecruiting
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim Manual As Boolean
Dim FecDesde As Date
Dim FecHasta As Date


'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------
Dim ArrPar
Dim NroConnExt As Long
Dim StrConnExt As String
Dim ternroER As Long
Dim DomNro As Long
Dim CantRHPuesto As Long
Dim CantRHMedio As Long
Dim CantRHBusq As Long
Dim CantRHExito As Long

Dim Indice As Long
'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Postulantes As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_RHPuesto As New ADODB.Recordset
Dim rs_RHMedio As New ADODB.Recordset
Dim rs_RHBusq As New ADODB.Recordset
Dim rs_con As New ADODB.Recordset


'Inicio codigo ejecutable
On Error GoTo E_ImpERecruiting


'-------------------------------------------------------------------------------------------------
'Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, "@")
        
        Select Case ArrPar(0)
            Case 0:
                'Disparo planificado
                Manual = False
                Flog.writeline Espacios(Tabulador * 1) & "Disparo Planificado"
                
            Case -1:
                'Disparo manual, viene la fecha desde y hasta
                Manual = True
                Flog.writeline Espacios(Tabulador * 1) & "Disparo Manual"
                
'                FecDesde = CDate(ArrPar(1))
'                Flog.writeline Espacios(Tabulador * 1) & "Parametro Fecha Desde = " & FecDesde
'
'                FecHasta = CDate(ArrPar(2))
'                Flog.writeline Espacios(Tabulador * 1) & "Parametro Fecha Hasta = " & FecHasta
            
            Case Else:
                Flog.writeline Espacios(Tabulador * 1) & "ERROR. Numero de parametros erroneo."
                Exit Sub
        End Select
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    HuboError = True
    Exit Sub
End If
Flog.writeline

FSinc.writeline Espacios(Tabulador * 0) & "Inicio "
FSinc.writeline

'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
FSinc.writeline
FSinc.writeline Espacios(Tabulador * 0) & "------------------------------------------------------"
FSinc.writeline Espacios(Tabulador * 0) & "Generales"
FSinc.writeline
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte 253."
FSinc.writeline Espacios(Tabulador * 1) & "Configuracion: Reporte 253."
FSinc.writeline Espacios(Tabulador * 2) & "Columna 1: Conexion a BD de ER"
NroConnExt = 0
StrSql = "SELECT * FROM confrep WHERE repnro = 253 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte 253."
    HuboError = True
    Exit Sub
Else
    NroTipoDomi = 2 'Particular
    Do While Not rs_Consult.EOF
    
    
        Select Case rs_Consult!confnrocol
            Case 1:
                NroConnExt = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case 1:
                NroTipoDomi = IIf(EsNulo(rs_Consult!confval), NroTipoDomi, rs_Consult!confval)
        End Select
    
        rs_Consult.MoveNext
    Loop
End If
rs_Consult.Close


If NroConnExt = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el valor de la conexion externa en el valor numerico de la columna 1."
    HuboError = True
    Exit Sub
End If


'-------------------------------------------------------------------------------------------------
'Busqueda de la conexion externa
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando la conexion externa " & NroConnExt
StrSql = "SELECT cnstring, cndesc FROM Conexion WHERE cnnro = " & NroConnExt
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    StrConnExt = IIf(EsNulo(rs_Consult!cnstring), "", rs_Consult!cnstring)
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la conexion externa."
    HuboError = True
    Exit Sub
End If
rs_Consult.Close

If Len(StrConnExt) = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la conexion externa."
    HuboError = True
    Exit Sub
End If
Flog.writeline

FSinc.writeline Espacios(Tabulador * 1) & "Sincronizacion: RHPRO --> ER"
FSinc.writeline
FSinc.writeline Espacios(Tabulador * 2) & "Tablas RHPRO:"
'-------------------------------------------------------------------------------------------------
'Busca si hay datos para pasar de RH a ER
'-------------------------------------------------------------------------------------------------
FSinc.writeline Espacios(Tabulador * 3) & "puestoER"
StrSql = "SELECT estrnro, puedesc, migrar"
StrSql = StrSql & " FROM puestoER"
StrSql = StrSql & " WHERE migrar = -1"
OpenRecordset StrSql, rs_RHPuesto
CantRHPuesto = rs_RHPuesto.RecordCount
FSinc.writeline Espacios(Tabulador * 4) & "SQL: " & StrSql

FSinc.writeline Espacios(Tabulador * 3) & "pos_medioER"
StrSql = "SELECT mednro, meddesabr, migrar"
StrSql = StrSql & " FROM pos_medioER"
StrSql = StrSql & " WHERE migrar = -1"
OpenRecordset StrSql, rs_RHMedio
CantRHMedio = rs_RHMedio.RecordCount
FSinc.writeline Espacios(Tabulador * 4) & "SQL: " & StrSql

FSinc.writeline Espacios(Tabulador * 3) & "pos_busquedaER"
StrSql = "SELECT busnro, busdesabr, busdesext, busfecha, busfecplanent, busq_requerimientos"
StrSql = StrSql & " FROM pos_busquedaER"
StrSql = StrSql & " WHERE migrar = -1"
OpenRecordset StrSql, rs_RHBusq
CantRHBusq = rs_RHBusq.RecordCount
FSinc.writeline Espacios(Tabulador * 4) & "SQL: " & StrSql

'-------------------------------------------------------------------------------------------------
'Abro la nueva conexion
'-------------------------------------------------------------------------------------------------
OpenConnExt StrConnExt, ExtConn


'-------------------------------------------------------------------------------------------------
'Abro la consulta de cambios de ER
'-------------------------------------------------------------------------------------------------
FSinc.writeline Espacios(Tabulador * 2) & "Tablas ER:"
FSinc.writeline Espacios(Tabulador * 3) & "recUSUSECAUDI"
FSinc.writeline Espacios(Tabulador * 3) & "recUSUARIOS"
    StrSql = "SELECT recUSUARIOS.usuUSER, recUSUARIOS.usuTERCEROID, usaID, usaSECCION"
    StrSql = StrSql & " FROM recUSUSECAUDI"
    StrSql = StrSql & " INNER JOIN recUSUARIOS ON recUSUSECAUDI.usaID = recUSUARIOS.usuID"
    StrSql = StrSql & " WHERE usaMIGRA = 1"
    StrSql = StrSql & " ORDER BY usaID, usaSECCION"
OpenRecordsetExt StrSql, rs_Postulantes, ExtConn
FSinc.writeline Espacios(Tabulador * 4) & "SQL: " & StrSql

'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Postulantes.RecordCount
If CEmpleadosAProc = 0 Then
    CEmpleadosAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Cambios a migrar desde ER a RHPRO: " & CEmpleadosAProc
End If
IncPorc = (100 / (CEmpleadosAProc + CantRHPuesto + CantRHMedio + CantRHBusq))
Flog.writeline
        
FSinc.writeline
FSinc.writeline Espacios(Tabulador * 0) & "------------------------------------------------------"
FSinc.writeline Espacios(Tabulador * 0) & "Detalle"
FSinc.writeline
FSinc.writeline Espacios(Tabulador * 1) & "RHPRO --> ER"

'-------------------------------------------------------------------------------------------------
'Comienzo de migracion de datos desde RHPRO a ER
'-------------------------------------------------------------------------------------------------
If (CantRHPuesto + CantRHMedio + CantRHBusq) <> 0 Then
    
    'Procesando Puestos
    '-------------------------------------------------------------------------------------
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 2) & "Puestos"
    FSinc.writeline Espacios(Tabulador * 3) & "recASPIRA (ER): Verifico si el codigo del puesto existe en ER. Si no existe ==> lo inserta, si existe ==> lo actualiza"
    
    If CantRHPuesto <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "Migrando Puestos desde RHPRO a ER"
        CantRHExito = 0
        
        Do While Not rs_RHPuesto.EOF
            'Verifico si el codigo del puesto existe en ER
            StrSql = "SELECT aspID, aspNOMBRE FROM recASPIRA"
            StrSql = StrSql & " WHERE aspID = " & rs_RHPuesto!estrnro
            OpenRecordsetExt StrSql, rs_Consult, ExtConn
            If rs_Consult.EOF Then
                CantRHExito = CantRHExito + 1
                
                'Creo el registro en ER
                StrSql = "SET IDENTITY_INSERT recASPIRA ON"
                StrSql = StrSql & " INSERT INTO recASPIRA("
                StrSql = StrSql & " aspID, aspNOMBRE"
                StrSql = StrSql & " )VALUES("
                StrSql = StrSql & " " & rs_RHPuesto!estrnro
                StrSql = StrSql & "," & CtrlNuloTXT(rs_RHPuesto!puedesc, 50)
                StrSql = StrSql & " )"
                StrSql = StrSql & " SET IDENTITY_INSERT recASPIRA OFF"
                ExtConn.Execute StrSql, , adExecuteNoRecords

                'Marco registro migrado
                StrSql = "UPDATE puestoER"
                StrSql = StrSql & " SET  migrar = 0"
                StrSql = StrSql & " WHERE estrnro = " & rs_RHPuesto!estrnro
                StrSql = StrSql & " AND puedesc = '" & rs_RHPuesto!puedesc & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Ya existe el codigo de puesto " & rs_Consult!aspID & " - " & rs_Consult!aspNOMBRE & " No se inserto el registro."
                
                'Marco registro de error
                StrSql = "UPDATE puestoER"
                StrSql = StrSql & " SET  migrar = 1"
                StrSql = StrSql & " WHERE estrnro = " & rs_RHPuesto!estrnro
                StrSql = StrSql & " AND puedesc = '" & rs_RHPuesto!puedesc & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_RHPuesto.MoveNext
        Loop
        
        'Progreso
        Progreso = Progreso + (IncPorc * CantRHPuesto)
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline Espacios(Tabulador * 1) & CantRHExito & " Puestos migrados desde RHPRO a ER"
        
    End If
    
    'Procesando Medios
    '-------------------------------------------------------------------------------------
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 2) & "Medios"
    FSinc.writeline Espacios(Tabulador * 3) & "pos_medio (ER): Verifico si el codigo del medio existe en ER. Si no existe ==> lo inserta, si existe ==> lo actualiza"
    
    If CantRHMedio <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "Migrando Medios desde RHPRO a ER"
        CantRHExito = 0
        
        Do While Not rs_RHMedio.EOF
            
            'Verifico si el codigo del medio existe en ER
            StrSql = "SELECT mednro, meddesabr FROM pos_medio"
            StrSql = StrSql & " WHERE mednro = " & rs_RHMedio!mednro
            OpenRecordsetExt StrSql, rs_Consult, ExtConn
            If rs_Consult.EOF Then
                CantRHExito = CantRHExito + 1
                
                'Creo el registro en ER
                StrSql = "SET IDENTITY_INSERT pos_medio ON"
                StrSql = StrSql & " INSERT INTO pos_medio("
                StrSql = StrSql & " mednro, meddesabr"
                StrSql = StrSql & " )VALUES("
                StrSql = StrSql & " " & rs_RHMedio!mednro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_RHMedio!meddesabr, 50)
                StrSql = StrSql & " )"
                StrSql = StrSql & " SET IDENTITY_INSERT pos_medio OFF"
                ExtConn.Execute StrSql, , adExecuteNoRecords

                'Marco registro migrado
                StrSql = "UPDATE pos_medioER"
                StrSql = StrSql & " SET  migrar = 0"
                StrSql = StrSql & " WHERE mednro = " & rs_RHMedio!mednro
                StrSql = StrSql & " AND meddesabr = '" & rs_RHMedio!meddesabr & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Ya existe el codigo de medio " & rs_RHMedio!mednro & " - " & rs_RHMedio!meddesabr & " No se inserto el registro."
                
                'Marco registro de error
                StrSql = "UPDATE pos_medioER"
                StrSql = StrSql & " SET  migrar = 1"
                StrSql = StrSql & " WHERE mednro = " & rs_RHMedio!mednro
                StrSql = StrSql & " AND meddesabr = '" & rs_RHMedio!meddesabr & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_RHMedio.MoveNext
        Loop
        
        'Progreso
        Progreso = Progreso + (IncPorc * CantRHMedio)
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline Espacios(Tabulador * 1) & CantRHExito & " Medios migrados desde RHPRO a ER"
    End If
    
    'Procesando Busquedas
    '-------------------------------------------------------------------------------------
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 2) & "Busquedas"
    FSinc.writeline Espacios(Tabulador * 3) & "rec_BUSQUEDA (ER): Verifico si el codigo de busqueda existe en ER. Si no existe ==> lo inserta, si existe ==> lo actualiza"
    
    If CantRHBusq <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "Migrando Busquedas desde RHPRO a ER"
        CantRHExito = 0
        
        Do While Not rs_RHBusq.EOF
            
            'Verifico si el codigo de la busqueda existe en ER
            StrSql = "SELECT reqpernro, busq_titulo FROM rec_BUSQUEDA"
            StrSql = StrSql & " WHERE id_rhpro = " & rs_RHBusq!busnro
            OpenRecordsetExt StrSql, rs_Consult, ExtConn
            'Si la busqueda NO existe en ER
            If rs_Consult.EOF Then
                CantRHExito = CantRHExito + 1
                
                'Creo el registro en ER
                StrSql = "INSERT INTO rec_BUSQUEDA"
                StrSql = StrSql & " (id_rhpro,busq_titulo,busq_desc_abreviada,busq_desc_extendida"
                StrSql = StrSql & " ,busq_fec_alta,  busq_fec_fin,activa,destacada,modelo, busq_requerimientos)"
                StrSql = StrSql & " VALUES("
                StrSql = StrSql & " " & rs_RHBusq!busnro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_RHBusq!busdesabr, 100)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_RHBusq!busdesabr, 1000)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_RHBusq!busdesext, 5000)
               ' StrSql = StrSql & " ," & cambiaFecha(Date)
               ' StrSql = StrSql & " ,NULL"
                StrSql = StrSql & " ," & cambiaFecha(rs_RHBusq!busfecha)
                StrSql = StrSql & " ," & cambiaFecha(rs_RHBusq!busfecplanent)

                StrSql = StrSql & " ,-1"
                StrSql = StrSql & " ,-1"
                StrSql = StrSql & " ,1"
                 StrSql = StrSql & " ," & CtrlNuloTXT(rs_RHBusq!busq_requerimientos, 5000)
                StrSql = StrSql & " )"
                ExtConn.Execute StrSql, , adExecuteNoRecords

                'Marco registro migrado
                StrSql = "UPDATE pos_busquedaER"
                StrSql = StrSql & " SET  migrar = 0"
                StrSql = StrSql & " WHERE busnro = " & rs_RHBusq!busnro
                StrSql = StrSql & " AND busdesabr = '" & rs_RHBusq!busdesabr & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Ya existe el codigo de busqueda " & rs_RHBusq!busnro & " - " & rs_RHBusq!busdesabr & " No se inserto el registro."
                
                'Marco registro de error
                StrSql = "UPDATE pos_busquedaER"
                StrSql = StrSql & " SET  migrar = -1"
                StrSql = StrSql & " WHERE busnro = " & rs_RHBusq!busnro
                StrSql = StrSql & " AND busdesabr = '" & rs_RHBusq!busdesabr & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_RHBusq.MoveNext
        Loop
        
        'Progreso
        Progreso = Progreso + (IncPorc * CantRHBusq)
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline Espacios(Tabulador * 1) & CantRHExito & " Busquedas migrados desde RHPRO a ER"
    End If
    
Else
    Flog.writeline Espacios(Tabulador * 0) & "No hay Cambios a migrar desde RHPRO a ER"
End If
        
        
FSinc.writeline
FSinc.writeline
FSinc.writeline Espacios(Tabulador * 1) & "ER --> RHPRO"
FSinc.writeline Espacios(Tabulador * 2) & "*** Si Se usa Legajo Electronico LE (configuracion en CONFPER 9) replica algunos datos en las tablas que pueden estar en otra BD"
FSinc.writeline
FSinc.writeline
''-------------------------------------------------------------------------------------------------
''Comienzo de migracion de datos desde ER a RHPRO
''-------------------------------------------------------------------------------------------------
If rs_Postulantes.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay Cambios a migrar desde ER a RHPRO"
Else
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Comienza el Procesamiento de Postulantes de ER"
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
End If


'FGZ - levanta configuracion para ver si se usa o no LE
Usa_LE = False

'StrSql = "SELECT eltoconfnro, eltoopnro "
'StrSql = StrSql & " FROM conf_general "
'StrSql = StrSql & " WHERE eltoconfnro = 6 "
'OpenRecordset StrSql, rs_Consult
'If Not rs_Consult.EOF Then
'    If rs_Consult!eltoopnro = 13 Then
'        Usa_LE = True
'    End If
'End If

'JPB - Recupero de la configuracion de empresa la columna 9 para ver si tiene activo el uso del legajo electrónico
StrSql = "SELECT * "
StrSql = StrSql & " FROM confper "
StrSql = StrSql & " WHERE confnro = 9 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    If rs_Consult!confactivo = -1 Then
        Usa_LE = True
    End If
End If


If Usa_LE Then
    'Establezco conexion a la BD temporal
    '---------------------------------------------------------------------
    'Establecer la conexion a la BD temporal
    StrSql = " SELECT cnnro, cnstring FROM conexion WHERE cnnro = 2 "
    OpenRecordset StrSql, rs_con
    If rs_con.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No se encuentra la conexion a la BD temporal para LE."
        Flog.writeline Espacios(Tabulador * 0) & "Se se replicaran los postulantes."
        Usa_LE = False
    Else
        On Error Resume Next
        'Abro la conexion a la BD Temporal
        OpenConnection rs_con!cnstring, ConnLE
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion. Debe Configurar bien la conexion a la BD temporal para LE."
            Flog.writeline Espacios(Tabulador * 0) & "Se se replicaran los postulantes."
            Usa_LE = False
        End If
    End If
    '---------------------------------------------------------------------

    'Si la conexion apunta a la misma BD que la productiva ==> No debe replicar ... solamente actualizar un campo en la tabla pos_postulante
    If Usa_LE Then
        Misma_BD = True
        Indice = 15 'DataSource Name
        If objConn.Properties(Indice) <> ConnLE.Properties(Indice) Then
            Misma_BD = False
        End If

        Indice = 49 'Server Name
        If objConn.Properties(Indice) <> ConnLE.Properties(Indice) Then
            Misma_BD = False
        End If

        Indice = 69 'Catalogo
        If objConn.Properties(Indice) <> ConnLE.Properties(Indice) Then
            Misma_BD = False
        End If

        Indice = 70 'Data Source
        If objConn.Properties(Indice) <> ConnLE.Properties(Indice) Then
            Misma_BD = False
        End If
    End If
End If



 

Do While Not rs_Postulantes.EOF

    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO : " & rs_Postulantes!usuUSER & " Seccion " & rs_Postulantes!usaSECCION

    '-------------------------------------------------------------------------------------------------
    'Segun la seccion modificada llamo al correspondiente procedimiento que se encarga de su actulizacion
    '-------------------------------------------------------------------------------------------------
    Select Case CInt(rs_Postulantes!usaSECCION)
        Case 1:
            Call Importar1(rs_Postulantes!usuTERCEROID)
        Case 2:
            Call Importar2(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 3:
            Call Importar3(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 4:
            Call Importar4(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 5:
            Call Importar5(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 6:
            Call Importar6(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 7:
            Call Importar7(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 8:
            Call Importar8(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 9:
            Call Importar9(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 10:
            Call Importar10(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 11:
            Call Importar11(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 12:
            Call Importar12(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 12:
            Call Importar13(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case 100:
            Call Importar100(rs_Postulantes!usuUSER, rs_Postulantes!usuTERCEROID)
        Case Else:
            Flog.writeline Espacios(Tabulador * 0) & "No existe la seccion"
    End Select


    '---------------------------------------------------------------------------------------------------------------
    'ACTUALIZO EL PROGRESO------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    rs_Postulantes.MoveNext

Loop


If rs_Postulantes.State = adStateOpen Then rs_Postulantes.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_RHPuesto.State = adStateOpen Then rs_RHPuesto.Close
If rs_RHMedio.State = adStateOpen Then rs_RHMedio.Close
If rs_RHBusq.State = adStateOpen Then rs_RHBusq.Close
Set rs_Postulantes = Nothing
Set rs_Consult = Nothing
Set rs_RHPuesto = Nothing
Set rs_RHMedio = Nothing
Set rs_RHBusq = Nothing

'Libero la conexion externo
ExtConn.Close
Set ExtConn = Nothing

Exit Sub

E_ImpERecruiting:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: ImpERecruiting"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


Public Function cambiaFecha(ByVal Fecha As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea la fecha al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    'If EsNulo(Fecha) Then
    If EsNulo(Fecha) Or Trim(Fecha) = "" Or (InStr(1, Fecha, "1900") > 0) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(Fecha)
    End If

End Function


Public Function CtrlNuloNUM(ByVal Valor) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea en numero al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If EsNulo(Valor) Then
        CtrlNuloNUM = 0
    Else
        If UCase(Valor) = "NULL" Then
            CtrlNuloNUM = 0
        Else
            If Not EsNum(CStr(Valor)) Then
                CtrlNuloNUM = 0
            Else
                CtrlNuloNUM = Valor
            End If
        End If
    End If
    
End Function

Private Function EsNum(ByVal a) As Boolean
Dim I As Long

    For I = 1 To Len(a)
        If ((Asc(Mid(a, I, 1)) < 48) Or (Asc(Mid(a, I, 1)) > 57)) And (Asc(Mid(a, I, 1)) <> 46) Then
            EsNum = False
            Exit Function
        End If
    Next I
    
    EsNum = True

End Function


Public Function CtrlNuloTXT(ByVal Valor, Optional ByVal Longitud As Long = 0) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea la texto al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If EsNulo(Valor) Then
        CtrlNuloTXT = "NULL"
    Else
        If UCase(Valor) = "NULL" Then
            CtrlNuloTXT = "NULL"
        Else
            If Longitud <> 0 Then
                CtrlNuloTXT = "'" & Mid(Valor, 1, Longitud) & "'"
            Else
                CtrlNuloTXT = "'" & Valor & "'"
            End If
        End If
    End If
    
End Function

Public Function CtrlNuloBOOL(ByVal Valor) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea el bool al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If IsNull(Valor) Then
        CtrlNuloBOOL = -1
    Else
        If UCase(Valor) = "NULL" Then
            CtrlNuloBOOL = -1
        Else
            If Valor = 1 Then
                CtrlNuloBOOL = -1
            Else
                CtrlNuloBOOL = 0
            End If
        End If
    End If
    
End Function


Public Sub OpenConnExt(strConnectionString As String, ByRef objConn As ADODB.Connection)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If objConn.State <> adStateClosed Then objConn.Close
    objConn.CursorLocation = adUseClient
    
    'objConn.IsolationLevel = adXactCursorStability
    'Indica que desde una transacción se pueden ver cambios que no se han producido
    'en otras transacciones.
    objConn.IsolationLevel = adXactReadUncommitted
    
    'objConn.IsolationLevel = adXactBrowse
    objConn.CommandTimeout = 3600 'segundos
    objConn.ConnectionTimeout = 60 'segundos
    objConn.Open strConnectionString
End Sub


Public Sub OpenRecordsetExt(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByVal objConnE As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre recordset de conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim pos1 As Long
Dim pos2 As Long
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    'Algunas propiedades de prueba
'    objRs.CursorType = 0 'adForwardOnly
'    objRs.CursorLocation = adUseServer
'    objRs.lockType = adLockReadOnly
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, objConnE, adOpenDynamic, lockType, adCmdText
    
'    pos1 = InStr(1, strSQLQuery, "from", vbTextCompare) + 5
'    If pos1 > 5 Then
'        pos2 = InStr(pos1, strSQLQuery, " ")
'        If pos2 = 0 Then
'            pos2 = Len(strSQLQuery)
'        End If
'        aux = Mid(strSQLQuery, pos1, pos2 - pos1)
'        Flog.writeline Espacios(Tabulador * 4) & "Tabla: " & aux
'    End If
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
End Sub


Public Function BuscarTerceroXMail(ByVal Mail As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion encargada de dado un mail obtener un tercero.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim rs_tercero As New ADODB.Recordset

    StrSql = "SELECT tercero.ternro, teremail"
    StrSql = StrSql & " FROM tercero"
    StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro"
    StrSql = StrSql & " AND ter_tip.tipnro = 14"
    StrSql = StrSql & " WHERE teremail = '" & Mail & "'"
    OpenRecordset StrSql, rs_tercero
    
    If rs_tercero.EOF Then
        BuscarTerceroXMail = 0
    Else
        BuscarTerceroXMail = rs_tercero!ternro
    End If
    
    rs_tercero.Close

Set rs_tercero = Nothing

End Function


Public Function BuscarTerceroTempXMail(ByVal Mail As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Dado un mail obtiene el campo tercerotemp.
' Autor      : Margiotta, Emanuel
' Fecha      : 04/01/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_terceroTemp As New ADODB.Recordset

    'Si LE esta instalado y esta en otra BD copia los datos
    If (Usa_LE) Then
        StrSql = "SELECT pos_postulante.tercerotemp"
        StrSql = StrSql & " FROM tercero"
        StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro"
        StrSql = StrSql & " AND ter_tip.tipnro = 14"
        StrSql = StrSql & " INNER JOIN pos_postulante ON tercero.ternro = pos_postulante.ternro"
        StrSql = StrSql & " WHERE teremail = '" & Mail & "'"
        OpenRecordset StrSql, rs_terceroTemp
    
        If rs_terceroTemp.EOF Then
            BuscarTerceroTempXMail = 0
        Else
            BuscarTerceroTempXMail = rs_terceroTemp!terceroTemp
        End If
    
        rs_terceroTemp.Close

        Set rs_terceroTemp = Nothing
    Else
    
        BuscarTerceroTempXMail = 0
    End If

End Function


Public Function MarcarSeccion(ByVal Mail As String, ByVal Seccion As Long, ByVal Valor As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Marcar la seccion de la auditoria de migracion.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Valor = -1 Modificacion OK
'Valor =  2 ERROR
Dim rs_Consult As New ADODB.Recordset
    
    'Busco el id del postulante
    StrSql = "SELECT usuid FROM recUSUARIOS  "
    StrSql = StrSql & " WHERE usuUSER = '" & Mail & "'"
    OpenRecordsetExt StrSql, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
    
        StrSql = "UPDATE recUSUSECAUDI"
        StrSql = StrSql & " SET usaMIGRA = " & Valor
        StrSql = StrSql & " WHERE usaID = " & rs_Consult!UsuId
        StrSql = StrSql & " AND usaSECCION = " & Seccion
        ExtConn.Execute StrSql, , adExecuteNoRecords
    
    End If
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

End Function

 
Public Sub Cargar_Imagen_Postulante(ByVal TernroRHPro As Long, ByVal rsPos As ADODB.Recordset)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de subir la imagen a la carpeta fotos y la asocia en la tabla ter_imag
' Autor      : Brzozowski Juan Pablo
' Fecha      : 12/02/2014
' ---------------------------------------------------------------------------------------------
      Dim rs_Consult As New ADODB.Recordset
      Dim mystream As New ADODB.Stream
      Dim tmp As String
      Dim Fn As Integer
      Dim sField As String
      Dim sObject As Object
      Dim sisdir As String
      Dim NuevoNombreImagen As String
      
      ' Controla la imagen que traigo de la base tenga algun valor
      If IsNull(rsPos!terFoto) Or rsPos.EOF = True Then
        Exit Sub
      End If
      
      'Si no esta el tipo de documente alta, sino modificadion
      StrSql = "SELECT  sisdir FROM sistema"
      OpenRecordset StrSql, rs_Consult
      If Not rs_Consult.EOF Then
       sisdir = rs_Consult!sisdir
      End If
      
      'Armo la ruta con el nombre de la imagen que se va a almacenar en la carpeta fotos
      NuevoNombreImagen = Format(Time, "hhmmss") & ".jpg"
      tmp = sisdir & "\fotos\" & NuevoNombreImagen
    
      'Asocio la imagen al postulante----------------------------------------------------------------
      StrSql = " DELETE ter_imag WHERE ternro= " & TernroRHPro & " AND ter_imag.tipimnro = 4  "
      StrSql = StrSql & " INSERT INTO ter_imag(ternro, tipimnro,terimnombre, terimpag, terimdesabr, terimfecha, terimvalid) "
      StrSql = StrSql & " VALUES(" & TernroRHPro & ",4,'" & NuevoNombreImagen & "',1,'Postulante Erecruiting',GETDATE(),-1)"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.writeline Espacios(Tabulador * 2) & "Asocio la imagen al postulante"
      '-----------------------------------------------------------------------------------------------
     
      ' Trae el proximo numero de archivo libre
      Fn = FreeFile
      ' Limpia el archivo de la imagen
      Open tmp For Random As Fn
      Close Fn
      'Copio en el stream la imagen de la base
      mystream.Open
      mystream.Type = adTypeBinary
      mystream.Write rsPos!terFoto
      ' Guardo el stream en el archivo
      mystream.SaveToFile tmp, adSaveCreateOverWrite
      mystream.Close
      Flog.writeline Espacios(Tabulador * 2) & "Se inserto la imagen del postulante en la carpeta fotos"
            
End Sub



Public Sub Importar1(ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 1.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim DomNro As Long
Dim TernroRHPro As Long
Dim TernroTemp As Long
Dim rs_Consult As New ADODB.Recordset
Dim rs_Post As New ADODB.Recordset

Dim CrearTercero As Integer     '0 - Error
                                '1 - Alta
                                '2 - Update
                                
On Error GoTo E_Importar1

MyBeginTrans

    CrearTercero = 0
    
    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 1 para Postulante de ER " & ternroER
    
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_1_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Trae datos Basicos (tercero, tipo de tercero, pos_postulante, documentos e imagen)"

    
    '---------------------------------------------------------------------------------------------------------------
    'Busco los datos del postulante
    '---------------------------------------------------------------------------------------------------------------
    OpenRecordsetExt "EXEC REC_MIGRA_SP_1_POS " & ternroER, rs_Post, ExtConn
    
    '---------------------------------------------------------------------------------------------------------------
    'Busco si en RhPro exista otro postulante con el mismo mail
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Validaciones sobre tercero."
    
    
    'Busco si en RhPro exista el mismo mail
    StrSql = "SELECT tercero.ternro, tercero.terape,tercero.ternom, pos_postulante.tercerotemp FROM tercero"
    StrSql = StrSql & " INNER JOIN pos_postulante ON pos_postulante.ternro = tercero.ternro"
    StrSql = StrSql & " WHERE tercero.teremail = '" & rs_Post!teremail & "'"
    OpenRecordset StrSql, rs_Consult
    If rs_Consult.EOF Then
        
        'No existe el mail Busco que si en RhPro existe el Documento
        StrSql = "SELECT ternro, tidnro, nrodoc"
        StrSql = StrSql & " FROM ter_doc"
        StrSql = StrSql & " WHERE nrodoc = " & "'" & rs_Post!terdoc & "'"
        StrSql = StrSql & " AND tidnro = " & rs_Post!tidnro
        OpenRecordset StrSql, rs_Consult
        If rs_Consult.EOF Then
            
            'Se crea el tercero
            CrearTercero = 1
            Flog.writeline Espacios(Tabulador * 2) & "Se procede a crear el tercero " & rs_Post!terape & " " & rs_Post!ternom & "(1)"
            
        Else
            
            'No existe el mail, existe el Documento en algun tercero
            TernroRHPro = rs_Consult!ternro
            
            'Verifico si el tercero con el Documento es postulante
            StrSql = "SELECT tipnro, ternro"
            StrSql = StrSql & " FROM ter_tip"
            StrSql = StrSql & " WHERE ternro = " & TernroRHPro
            'StrSql = StrSql & " AND ternro = 14"
            StrSql = StrSql & " AND tipnro = 14"
            OpenRecordset StrSql, rs_Consult
            If rs_Consult.EOF Then
                
                'Error, el documento pertenece a un tercero no postulante
                CrearTercero = 0
                Flog.writeline Espacios(Tabulador * 2) & "ERROR! El documento tipo " & rs_Post!tidnro & " NRO " & rs_Post!terdoc & " del postulante a insertar " & rs_Post!terape & " " & rs_Post!ternom & " PERTENECE a un tercero No postulante en RHPro (2)"
                
            Else
                
                'Modificacion de un postulante
                CrearTercero = 2
                Flog.writeline Espacios(Tabulador * 2) & "Se procede a Modificar el tercero con mail " & rs_Post!teremail & "(3)"
                
            End If
            
        End If
        
    Else
    
        'Existe el mail en un tercero de RHPro
        TernroRHPro = rs_Consult!ternro
        '--------------------------------
        'FGZ - 30/12/2009
        If Usa_LE Then
            TernroTemp = IIf(Not EsNulo(rs_Consult!terceroTemp), rs_Consult!terceroTemp, 0)
        End If
        '--------------------------------
        
        'Verifico si el tercero con el mail es postulante
        StrSql = "SELECT tipnro, ternro"
        StrSql = StrSql & " FROM ter_tip"
        StrSql = StrSql & " WHERE ternro = " & TernroRHPro
        StrSql = StrSql & " AND tipnro = 14"
        OpenRecordset StrSql, rs_Consult
        If rs_Consult.EOF Then
        
            'Error, el mail pertenece a un tercero no postulante
            CrearTercero = 0
            Flog.writeline Espacios(Tabulador * 2) & "ERROR! El mail " & rs_Post!teremail & " del postulante a insertar " & rs_Post!terape & " " & rs_Post!ternom & " PERTENECE a otro tercero No postulante en RHPro (4)"
            
        Else
        
            'El mail pertenece a un postulante verifico si tb coincide el documento para ese tercero
            StrSql = "SELECT ternro, tidnro, nrodoc"
            StrSql = StrSql & " FROM ter_doc"
            StrSql = StrSql & " WHERE nrodoc = " & "'" & rs_Post!terdoc & "'"
            StrSql = StrSql & " AND tidnro = " & rs_Post!tidnro
            StrSql = StrSql & " AND ternro = " & TernroRHPro
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                
                'El tercero tiene el mismo mail y dni y es postulante - Modificacion
                CrearTercero = 2
                Flog.writeline Espacios(Tabulador * 2) & "Se procede a Modificar el tercero con mail " & rs_Post!teremail & " coincide doc y mail (5)"
            
            Else
                
                'Busco si el documento pertenece a algun tercero
                StrSql = "SELECT ternro, tidnro, nrodoc"
                StrSql = StrSql & " FROM ter_doc"
                StrSql = StrSql & " WHERE nrodoc = " & "'" & rs_Post!terdoc & "'"
                StrSql = StrSql & " AND tidnro = " & rs_Post!tidnro
                OpenRecordset StrSql, rs_Consult
                If rs_Consult.EOF Then
                
                    'El documento No pertenece a nadie entonces modifico el tercero
                    CrearTercero = 2
                    Flog.writeline Espacios(Tabulador * 2) & "Se procede a Modificar el tercero con mail " & rs_Post!teremail & " coincide mail y no doc (6)"
                
                Else
                
                    'Error, el mail pertenece a un tercero no postulante
                    CrearTercero = 0
                    Flog.writeline Espacios(Tabulador * 2) & "ERROR! El documento de " & rs_Post!teremail & " a insertar " & rs_Post!terape & " " & rs_Post!ternom & "  PERTENECE a otro tercero No postulante en RHPro (7)"
                
                End If
                
            End If
            
        End If
        
    End If
    
    '---------------------------------------------------------------------------------------------------------------
    'ABM de tercero segun sea el caso correspondiente
    '---------------------------------------------------------------------------------------------------------------
    Select Case CrearTercero
        
        Case 0:
            
            Flog.writeline Espacios(Tabulador * 1) & "No se realiza ningun cambio."
            
            'Marco la seccion como migrada sin error
            Call MarcarSeccion(rs_Post!teremail, 1, 2)
            
        Case 1:
            
             
            'Inserto Tercero
            StrSql = "INSERT INTO tercero (ternom, ternom2, terape, terape2, terfecnac, tersex, teremail, nacionalnro, estcivnro, terfecing, terfecestciv, paisnro) VALUES "
            StrSql = StrSql & "(" & CtrlNuloTXT(rs_Post!ternom, 25)
            StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!ternom2, 25)
            StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!terape, 25)
            StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!terape2, 25)
            StrSql = StrSql & "," & cambiaFecha(rs_Post!terfecnac)
            StrSql = StrSql & "," & IIf(rs_Post!tersex = 1, -1, 0)
            StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!teremail, 100)
            StrSql = StrSql & "," & CtrlNuloNUM(rs_Post!nacionalnro)
            StrSql = StrSql & "," & CtrlNuloNUM(rs_Post!estcivnro)
            StrSql = StrSql & "," & cambiaFecha(rs_Post!terfecing)
            StrSql = StrSql & "," & cambiaFecha(rs_Post!terfecestciv)
            StrSql = StrSql & "," & CtrlNuloNUM(rs_Post!paisnro)
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            
            
            
            'Obtengo el ternro
            TernroRHPro = getLastIdentity(objConn, "tercero")
            Flog.writeline Espacios(Tabulador * 2) & "Se inserto en tercero " & TernroRHPro

            'Creacion del tipo de tercero postulante
            StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & TernroRHPro & ",14)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 2) & "Inserto el tipo de tercero 14 en ter_tip"
            
            'Creacion del postulante vacio
            StrSql = "INSERT INTO pos_postulante(ternro,posfecpres,"
            StrSql = StrSql & " estposnro,posest)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & TernroRHPro
            StrSql = StrSql & " ," & cambiaFecha(Date)
            StrSql = StrSql & " ,4"
            StrSql = StrSql & " ,-1"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 2) & "Inserto el postulante"
        
            'Creacion del Documento
            StrSql = "INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & TernroRHPro & "," & rs_Post!tidnro & "," & CtrlNuloTXT(rs_Post!terdoc, 30) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 2) & "Inserto Documento"
            
            'Marco la seccion como migrada sin error
            Call MarcarSeccion(rs_Post!teremail, 1, -1)
            
            
            
            '--------------------------------------------------------------------------------------------------
            ' 30/12/2009 - FGZ
            'Si usa Legajo Electronico entonces habria que replicar los postulantes en la BD temporal (ets abase podria ser la misma que la BD productiva de rhpro)
            If Usa_LE Then
                If Misma_BD Then
                    StrSql = "UPDATE pos_postulante SET tercerotemp = " & TernroRHPro & " WHERE ternro = " & TernroRHPro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "postulante replicado"
                Else
                    'Tengo que hacer los inserts
                    
                    'Inserto Tercero
                    StrSql = "INSERT INTO tercero (ternom, ternom2, terape, terape2, terfecnac, tersex, teremail, nacionalnro, estcivnro, terfecing, terfecestciv, paisnro) VALUES "
                    StrSql = StrSql & "(" & CtrlNuloTXT(rs_Post!ternom, 25)
                    StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!ternom2, 25)
                    StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!terape, 25)
                    StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!terape2, 25)
                    StrSql = StrSql & "," & cambiaFecha(rs_Post!terfecnac)
                    StrSql = StrSql & "," & IIf(rs_Post!tersex = 1, -1, 0)
                    StrSql = StrSql & "," & CtrlNuloTXT(rs_Post!teremail, 100)
                    StrSql = StrSql & "," & CtrlNuloNUM(rs_Post!nacionalnro)
                    StrSql = StrSql & "," & CtrlNuloNUM(rs_Post!estcivnro)
                    StrSql = StrSql & "," & cambiaFecha(rs_Post!terfecing)
                    StrSql = StrSql & "," & cambiaFecha(rs_Post!terfecestciv)
                    StrSql = StrSql & "," & CtrlNuloNUM(rs_Post!paisnro)
                    StrSql = StrSql & ")"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    
                    
                    'Obtengo el ternro
                    TernroTemp = getLastIdentity(ConnLE, "tercero")
                    Flog.writeline Espacios(Tabulador * 2) & "Replica en tercero " & TernroTemp
        
                    'Creacion del tipo de tercero postulante
                    StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & TernroTemp & ",14)"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Replica del tipo de tercero 14 en ter_tip"
                    
                    'Creacion del postulante vacio
                    StrSql = "INSERT INTO pos_postulante(ternro,posfecpres,"
                    StrSql = StrSql & " estposnro,posest,tercerotemp )"
                    StrSql = StrSql & " VALUES"
                    StrSql = StrSql & " (" & TernroTemp
                    StrSql = StrSql & " ," & cambiaFecha(Date)
                    StrSql = StrSql & " ,4"
                    StrSql = StrSql & " ,-1"
                    StrSql = StrSql & " ," & TernroTemp
                    StrSql = StrSql & " )"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Replico en  pos_postulante"
                
                    'Creacion del Documento
                    StrSql = "INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                    StrSql = StrSql & " VALUES(" & TernroTemp & "," & rs_Post!tidnro & "," & CtrlNuloTXT(rs_Post!terdoc, 30) & ")"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Replico el Documento"
                    
                    'Actualizo la BD Prod
                    StrSql = "UPDATE pos_postulante SET tercerotemp = " & TernroTemp & " WHERE ternro = " & TernroRHPro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Nro de postulante replicado actualizado en BD Prod"
                End If
            End If
            '--------------------------------------------------------------------------------------------------
           
        Case 2:
               
        
            'Modificacion de tercero
            StrSql = "UPDATE tercero SET "
            StrSql = StrSql & "  ternom = " & CtrlNuloTXT(rs_Post!ternom, 25)
            StrSql = StrSql & " ,terape = " & CtrlNuloTXT(rs_Post!terape, 25)
            StrSql = StrSql & " ,ternom2 = " & CtrlNuloTXT(rs_Post!ternom2, 25)
            StrSql = StrSql & " ,terape2 = " & CtrlNuloTXT(rs_Post!terape2, 25)
            StrSql = StrSql & " ,terfecnac = " & cambiaFecha(rs_Post!terfecnac)
            StrSql = StrSql & " ,tersex = " & IIf(rs_Post!tersex = 1, -1, 0)
            StrSql = StrSql & " ,terfecing = " & cambiaFecha(rs_Post!terfecing)
            StrSql = StrSql & " ,paisnro = " & CtrlNuloNUM(rs_Post!paisnro)
            StrSql = StrSql & " ,estcivnro = " & CtrlNuloNUM(rs_Post!estcivnro)
            StrSql = StrSql & " ,terfecestciv = " & cambiaFecha(rs_Post!terfecestciv)
            StrSql = StrSql & " ,nacionalnro = " & CtrlNuloNUM(rs_Post!nacionalnro)
            StrSql = StrSql & " ,teremail = " & CtrlNuloTXT(rs_Post!teremail)
       '           StrSql = StrSql & " ,terFoto = " & rs_Post!terFoto
            
            StrSql = StrSql & "  WHERE ternro = " & TernroRHPro
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 2) & "Update en tercero " & TernroRHPro
            
            'Si no esta el tipo de documente alta, sino modificadion
            StrSql = "SELECT ternro, tidnro"
            StrSql = StrSql & " FROM ter_doc"
            StrSql = StrSql & " WHERE tidnro = " & rs_Post!tidnro
            StrSql = StrSql & " AND ternro = " & TernroRHPro
            OpenRecordset StrSql, rs_Consult
            If rs_Consult.EOF Then
                'Inserto
                StrSql = "INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & TernroRHPro & "," & rs_Post!tidnro & "," & CtrlNuloTXT(rs_Post!terdoc, 30) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 2) & "Inserto Documento"
            
            Else
                'Modifico
                StrSql = "UPDATE ter_doc"
                StrSql = StrSql & " SET nrodoc = " & CtrlNuloTXT(rs_Post!terdoc, 30)
                StrSql = StrSql & " WHERE ternro = " & TernroRHPro
                StrSql = StrSql & " AND tidnro = " & rs_Post!tidnro
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 2) & "Modifico Documento"
                
            End If
    
    
            '---------------JPB: SUBE LA IMAGEN A LA CARPETA FOTOS--------
            Call Cargar_Imagen_Postulante(TernroRHPro, rs_Post)
            '--------------------------------------------------------
     
            '--------------------------------------------------------------------------------------------------
            ' 30/12/2009 - FGZ
            'Si usa Legajo Electronico entonces habria que replicar los postulantes en la BD temporal (ets abase podria ser la misma que la BD productiva de rhpro)
            If Usa_LE Then
                If Misma_BD Then
                    'No tengo que hacer nada
                Else
                    'Modificacion de tercero
                    StrSql = "UPDATE tercero SET "
                    StrSql = StrSql & "  ternom = " & CtrlNuloTXT(rs_Post!ternom, 25)
                    StrSql = StrSql & " ,terape = " & CtrlNuloTXT(rs_Post!terape, 25)
                    StrSql = StrSql & " ,ternom2 = " & CtrlNuloTXT(rs_Post!ternom2, 25)
                    StrSql = StrSql & " ,terape2 = " & CtrlNuloTXT(rs_Post!terape2, 25)
                    StrSql = StrSql & " ,terfecnac = " & cambiaFecha(rs_Post!terfecnac)
                    StrSql = StrSql & " ,tersex = " & IIf(rs_Post!tersex = 1, -1, 0)
                    StrSql = StrSql & " ,terfecing = " & cambiaFecha(rs_Post!terfecing)
                    StrSql = StrSql & " ,paisnro = " & CtrlNuloNUM(rs_Post!paisnro)
                    StrSql = StrSql & " ,estcivnro = " & CtrlNuloNUM(rs_Post!estcivnro)
                    StrSql = StrSql & " ,terfecestciv = " & cambiaFecha(rs_Post!terfecestciv)
                    StrSql = StrSql & " ,nacionalnro = " & CtrlNuloNUM(rs_Post!nacionalnro)
                    StrSql = StrSql & " ,teremail = " & CtrlNuloTXT(rs_Post!teremail)
                    'StrSql = StrSql & " ,terFoto = " & rs_Post!terFoto
                    StrSql = StrSql & "  WHERE ternro = " & TernroTemp
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Update en replica tercero " & TernroTemp
                    
                    'Si no esta el tipo de documente alta, sino modificadion
                    StrSql = "SELECT ternro, tidnro"
                    StrSql = StrSql & " FROM ter_doc"
                    StrSql = StrSql & " WHERE tidnro = " & rs_Post!tidnro
                    StrSql = StrSql & " AND ternro = " & TernroTemp
                    'OpenRecordset StrSql, rs_Consult
                    OpenRecordsetWithConn StrSql, rs_Consult, ConnLE
                    If rs_Consult.EOF Then
                        'Inserto
                        StrSql = "INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                        StrSql = StrSql & " VALUES(" & TernroTemp & "," & rs_Post!tidnro & "," & CtrlNuloTXT(rs_Post!terdoc, 30) & ")"
                        ConnLE.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 2) & "Inserto Documento en replica"
                    Else
                        'Modifico
                        StrSql = "UPDATE ter_doc"
                        StrSql = StrSql & " SET nrodoc = " & CtrlNuloTXT(rs_Post!terdoc, 30)
                        StrSql = StrSql & " WHERE ternro = " & TernroTemp
                        StrSql = StrSql & " AND tidnro = " & rs_Post!tidnro
                        ConnLE.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 2) & "Modifico Documento en replica"
                    End If
                End If
            End If
            '--------------------------------------------------------------------------------------------------
            
            
            'Marco la seccion como migrada sin error
            Call MarcarSeccion(rs_Post!teremail, 1, -1)
            
    End Select
    
    

If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
If rs_Post.State = adStateOpen Then rs_Post.Close
Set rs_Post = Nothing

MyCommitTrans

Exit Sub
E_Importar1:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar1"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans
    
    
End Sub


Public Sub Importar2(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 2.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
'    30/12/2009 Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim rs_Consult2 As New ADODB.Recordset   'Agregado ver 1.10
Dim TernroTemp As Long

On Error GoTo E_Importar2

MyBeginTrans

    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 2 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_2_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Algunos datos extra del Postulantes y actualiza pos_postulante + puesto al que aspira"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 2, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
            
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Modificacion del postulante
    '---------------------------------------------------------------------------------------------------------------
    OpenRecordsetExt "EXEC REC_MIGRA_SP_2 " & ternroER, rs_Consult, ExtConn
    
    'Agregado ver 1.10
    StrSql = "SELECT aspID, aspNOMBRE FROM recASPIRA"
    StrSql = StrSql & " WHERE aspID = " & CtrlNuloNUM(rs_Consult!aspID)
    OpenRecordsetExt StrSql, rs_Consult2, ExtConn

           
    If Not rs_Consult.EOF Then
        
        StrSql = "UPDATE pos_postulante SET "
        StrSql = StrSql & "  posrempre = " & CtrlNuloNUM(rs_Consult!posrempre)
        StrSql = StrSql & " ,poscanhijos = " & CtrlNuloNUM(rs_Consult!poscanhijos)
        StrSql = StrSql & " ,posest = -1"
        StrSql = StrSql & " ,posfecpres = " & cambiaFecha(rs_Consult!posfecpres)
        StrSql = StrSql & " ,posref = " & CtrlNuloTXT(rs_Consult!posref, 250)
        StrSql = StrSql & " ,posfecdis = " & cambiaFecha(rs_Consult!posfecdis)
        StrSql = StrSql & " ,estposnro = 4"
        StrSql = StrSql & " ,posobjprofesional = " & CtrlNuloTXT(rs_Consult!posobjprofesional, 500)
        If Not rs_Consult2.EOF Then   'Agregado ver 1.10
            StrSql = StrSql & " ,pospuestoaspira = " & CtrlNuloTXT(rs_Consult2!aspNOMBRE, 200)
        End If
        StrSql = StrSql & " ,posdispviajar = " & CtrlNuloBOOL(rs_Consult!terviajar)
        StrSql = StrSql & " ,posdisreubicarse = " & CtrlNuloBOOL(rs_Consult!terreubicarse)
        
        'JPB - controlo que no sea nulo terfuma
        If CtrlNuloBOOL(rs_Consult!terfuma) <> -1 Then
          StrSql = StrSql & " ,posfumador = " & rs_Consult!terfuma
        Else
          StrSql = StrSql & " ,posfumador = 0 "
        End If
        
        StrSql = StrSql & " ,postrabajaact = " & CtrlNuloBOOL(rs_Consult!tertrabaja)
        
        'JPB - controlo que no sea nulo tervehiculo
        If CtrlNuloBOOL(rs_Consult!tervehiculo) <> -1 Then
            StrSql = StrSql & " ,posvehiculo = " & rs_Consult!tervehiculo
        Else
            StrSql = StrSql & " ,posvehiculo = 0 "
        End If
            
        
        StrSql = StrSql & " ,mednro = " & CtrlNuloNUM(rs_Consult!mednro)
        StrSql = StrSql & "  WHERE ternro = " & TernroRHPro
        objConn.Execute StrSql, , adExecuteNoRecords
            
        Flog.writeline Espacios(Tabulador * 1) & "Postulante Modificado"
                                        
        
        'Borro el estruc_aplica de puesto
        StrSql = "DELETE FROM estruc_aplica"
        StrSql = StrSql & " WHERE ternro = " & TernroRHPro
        StrSql = StrSql & " AND tenro = 4"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra estruc_aplica de puesto"
        
        If CtrlNuloNUM(rs_Consult!aspID) <> 0 Then
        
            'Verifico que la estructura a insertar exista
            StrSql = "SELECT estrnro, tenro FROM estructura WHERE tenro = 4 AND estrnro = " & rs_Consult!aspID
            OpenRecordset StrSql, rs_Estr
            If Not rs_Estr.EOF Then
            
                StrSql = "INSERT INTO estruc_aplica "
                StrSql = StrSql & "(tenro, ternro, estrnro) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & " ( 4"
                StrSql = StrSql & " , " & TernroRHPro
                StrSql = StrSql & " , " & rs_Consult!aspID
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 1) & "Puesto al que aspira creado."
                
            Else
                Flog.writeline Espacios(Tabulador * 1) & "No se crea puesto al que aspira porque en RHPro no existe la estructura " & rs_Consult!aspID
            End If
        End If
        
        'EAM - 30/12/2009 - LE
        'Si usa Legajo Electronico entonces habria que copiar los mismos datos en la BD temporal (esta base podria ser la misma que la BD productiva de rhpro)
        If Usa_LE Then
            If Not Misma_BD Then
                StrSql = "UPDATE pos_postulante SET "
                StrSql = StrSql & "  posrempre = " & CtrlNuloNUM(rs_Consult!posrempre)
                StrSql = StrSql & " ,poscanhijos = " & CtrlNuloNUM(rs_Consult!poscanhijos)
                StrSql = StrSql & " ,posest = -1"
                StrSql = StrSql & " ,posfecpres = " & cambiaFecha(rs_Consult!posfecpres)
                StrSql = StrSql & " ,posref = " & CtrlNuloTXT(rs_Consult!posref, 250)
                StrSql = StrSql & " ,posfecdis = " & cambiaFecha(rs_Consult!posfecdis)
                StrSql = StrSql & " ,estposnro = 4"
                StrSql = StrSql & " ,posobjprofesional = " & CtrlNuloTXT(rs_Consult!posobjprofesional, 500)
                If Not rs_Consult2.EOF Then       'Agregado ver 1.10
                    StrSql = StrSql & " ,pospuestoaspira = " & CtrlNuloTXT(rs_Consult2!aspNOMBRE, 200)
                End If
                StrSql = StrSql & " ,posdispviajar = " & CtrlNuloBOOL(rs_Consult!terviajar)
                StrSql = StrSql & " ,posdisreubicarse = " & CtrlNuloBOOL(rs_Consult!terreubicarse)
                
                'JPB - controlo que no sea nulo terfuma
                If CtrlNuloBOOL(rs_Consult!terfuma) <> -1 Then
                   StrSql = StrSql & " ,posfumador = " & CtrlNuloBOOL(rs_Consult!terfuma)
                Else
                  StrSql = StrSql & " ,posfumador = 0 "
                End If
                'StrSql = StrSql & " ,posfumador = " & CtrlNuloBOOL(rs_Consult!terfuma)
                
                
                StrSql = StrSql & " ,postrabajaact = " & CtrlNuloBOOL(rs_Consult!tertrabaja)
                
                'JPB - controlo que no sea nulo tervehiculo
                If CtrlNuloBOOL(rs_Consult!tervehiculo) <> -1 Then
                    StrSql = StrSql & " ,posvehiculo = " & rs_Consult!tervehiculo
                Else
                    StrSql = StrSql & " ,posvehiculo = 0 "
                End If
                'StrSql = StrSql & " ,posvehiculo = " & CtrlNuloBOOL(rs_Consult!tervehiculo)
                
                StrSql = StrSql & " ,mednro = " & CtrlNuloNUM(rs_Consult!mednro)
                StrSql = StrSql & "  WHERE ternro = " & TernroTemp
                ConnLE.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 1) & "Postulante Modificado en la BD LE"
                
                'EAM - Borro el estruc_aplica de puesto
                StrSql = "DELETE FROM estruc_aplica"
                StrSql = StrSql & " WHERE ternro = " & TernroTemp
                StrSql = StrSql & " AND tenro = 4"
                ConnLE.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Borra estruc_aplica de puesto en la BD LE"
                
                If CtrlNuloNUM(rs_Consult!aspID) <> 0 Then
                
                    'EAM - Verifico que la estructura a insertar exista
                    StrSql = "SELECT estrnro, tenro FROM estructura WHERE tenro = 4 AND estrnro = " & rs_Consult!aspID
                    OpenRecordset StrSql, rs_Estr
                    If Not rs_Estr.EOF Then
                    
                        StrSql = "INSERT INTO estruc_aplica "
                        StrSql = StrSql & "(tenro, ternro, estrnro) "
                        StrSql = StrSql & " VALUES "
                        StrSql = StrSql & " ( 4"
                        StrSql = StrSql & " , " & TernroTemp
                        StrSql = StrSql & " , " & rs_Consult!aspID
                        StrSql = StrSql & " )"
                        ConnLE.Execute StrSql, , adExecuteNoRecords
                        
                        Flog.writeline Espacios(Tabulador * 1) & "Puesto al que aspira creado en la BD LE."
                        
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "No se crea puesto al que aspira porque en RHPro no existe la estructura " & rs_Consult!aspID & "en la BD LE"
                    End If
                End If
            End If
        End If
       
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 2"
    
    End If
    rs_Consult.Close
    
    'Marco la seccion como migrada sin error
    Call MarcarSeccion(Mail, 2, -1)
    

If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

MyCommitTrans

Exit Sub
E_Importar2:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar2"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 2, 2)
    
End Sub


Public Sub Importar3(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 3.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'    04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe

' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim DomNro As Long
Dim TernroTemp As Long

Dim rs_Consult As New ADODB.Recordset
Dim rs_DomNro As New ADODB.Recordset

On Error GoTo E_Importar3

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 3 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_3_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Domicilios"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 3, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Creacion del domicilio y telefono
    '---------------------------------------------------------------------------------------------------------------
    OpenRecordsetExt "EXEC REC_MIGRA_SP_3 " & ternroER, rs_Consult, ExtConn
    
    If Not rs_Consult.EOF Then
        
        DomNro = 0
        'Busco si tiene algun domicilio
        StrSql = "SELECT domnro"
        StrSql = StrSql & " FROM cabdom"
        StrSql = StrSql & " WHERE ternro = " & TernroRHPro
        OpenRecordset StrSql, rs_DomNro
        If Not rs_DomNro.EOF Then
            
            'Busco si tiene domicilio principal del tipo de domicilio a insertar para modificarlo
            StrSql = "SELECT domnro"
            StrSql = StrSql & " FROM cabdom"
            StrSql = StrSql & " WHERE ternro = " & TernroRHPro
            StrSql = StrSql & " AND domdefault = -1"
            'FGZ - 27/08/2013 ------------------------------
            'StrSql = StrSql & " AND tidonro = " & CtrlNuloNUM(rs_Consult!tidonro)
            StrSql = StrSql & " AND tidonro = " & NroTipoDomi
            'FGZ - 27/08/2013 ------------------------------
            OpenRecordset StrSql, rs_DomNro
            If Not rs_DomNro.EOF Then
                
                'Este el domicilio a modificar
                DomNro = rs_DomNro!DomNro
                
            Else
                
                'Busco si tiene un domicilio del tipo a insertar
                StrSql = "SELECT domnro"
                StrSql = StrSql & " FROM cabdom"
                StrSql = StrSql & " WHERE ternro = " & TernroRHPro
                'FGZ - 27/08/2013 ------------------------------
                'StrSql = StrSql & " AND tidonro = " & CtrlNuloNUM(rs_Consult!tidonro)
                StrSql = StrSql & " AND tidonro = " & NroTipoDomi
                'FGZ - 27/08/2013 ------------------------------
                OpenRecordset StrSql, rs_DomNro
                If Not rs_DomNro.EOF Then
                
                    'Este el domicilio a modificar
                    DomNro = rs_DomNro!DomNro
                    
                End If
                
                'Como el domicilio a modificar va a ser el default marco TODOS lo domicilios como no default
                StrSql = "UPDATE cabdom"
                StrSql = StrSql & " SET domdefault = 0 "
                StrSql = StrSql & " WHERE ternro = " & TernroRHPro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'EAM - 04/01/2010
                If (Usa_LE) And (Not Misma_BD) Then
                    'Como el domicilio a modificar va a ser el que quede por default marco TODOS lo domicilios como no default
                    StrSql = "UPDATE cabdom" & _
                             " SET domdefault = 0 " & _
                             " WHERE ternro = " & TernroTemp
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                End If
                
            End If
            
        Else
        
            'No tiene domicilio entonces creo uno
            DomNro = 0
        
        End If
    
        
        If DomNro = 0 Then
                
            'Alta de un domicilio nuevo - creo la cabecera del demicilio
            'JPB - Antes de insertar el nuevo domicilio desactivo los FK de la tabla cabdom. Luego las activo nuevamente.
            StrSql = " ALTER TABLE cabdom "
            StrSql = StrSql & " nocheck constraint fk_cabtipodom "
            StrSql = StrSql & " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
            'FGZ - 27/08/2013 ------------------------------
            'StrSql = StrSql & " VALUES(1," & TernroRHPro & ",-1," & CtrlNuloNUM(rs_Consult!tidonro) & ")"
            StrSql = StrSql & " VALUES(1," & TernroRHPro & ",-1," & NroTipoDomi & ")"
            'FGZ - 27/08/2013 ------------------------------
            StrSql = StrSql & " ALTER TABLE cabdom "
            StrSql = StrSql & " check constraint fk_cabtipodom "
            
            objConn.Execute StrSql, , adExecuteNoRecords
        
            'Obtengo el numero de domicilio en la tabla
            DomNro = CLng(getLastIdentity(objConn, "cabdom"))
               
            'Inserto el domicilio
            StrSql = "INSERT INTO detdom"
            StrSql = StrSql & " (domnro,calle,nro"
            StrSql = StrSql & " ,torre,piso"
            StrSql = StrSql & " ,oficdepto,codigopostal"
            StrSql = StrSql & " ,locnro,provnro,paisnro"
            StrSql = StrSql & " ,zonanro,partnro,barrio"
            StrSql = StrSql & " ,entrecalles)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & DomNro
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!calle, 250)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!nro, 8)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!torre, 8)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!piso, 8)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!oficdepto, 8)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!codigopostal, 12)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!locnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!provnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!paisnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!zonanro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!partnro)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!barrio, 30)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!entrecalles, 80)
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto el Domicilio"
            
            'JPB - Inserto el telefono - Personal
            If Not EsNulo(rs_Consult!telnro) Then
                StrSql = "INSERT INTO telefono"
                StrSql = StrSql & " (domnro,telnro,tipotel)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & DomNro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!telnro, 20)
                StrSql = StrSql & " ,1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Personal " & rs_Consult!telnro
            End If
            
            'JPB - Inserto el telefono - Celular
            If Not EsNulo(rs_Consult!auxchr1) Then
                StrSql = "INSERT INTO telefono"
                StrSql = StrSql & " (domnro,telnro,tipotel)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & DomNro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!auxchr1, 20)
                StrSql = StrSql & " ,2)"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Celular " & rs_Consult!auxchr1
            End If
            
            'EAM - 04/01/2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Alta de un domicilio nuevo - creo la cabecera del demicilio
                
                
            'JPB - Antes de insertar el nuevo domicilio desactivo los FK de la tabla cabdom. Luego las activo nuevamente.
            StrSql = " ALTER TABLE cabdom "
            StrSql = StrSql & " nocheck constraint fk_cabtipodom "
            StrSql = StrSql & " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
            'FGZ - 27/08/2013 ------------------------------
            'StrSql = StrSql & " VALUES(1," & TernroTemp & ",-1," & CtrlNuloNUM(rs_Consult!tidonro) & ")"
            StrSql = StrSql & " VALUES(1," & TernroTemp & ",-1," & NroTipoDomi & ")"
            'FGZ - 27/08/2013 ------------------------------
            StrSql = StrSql & " ALTER TABLE cabdom "
            StrSql = StrSql & " check constraint fk_cabtipodom "
                
            ConnLE.Execute StrSql, , adExecuteNoRecords
            
                'Obtengo el numero de domicilio en la tabla
                DomNro = CLng(getLastIdentity(ConnLE, "cabdom"))
                   
                'Inserto el domicilio
                StrSql = "INSERT INTO detdom"
                StrSql = StrSql & " (domnro,calle,nro,torre,piso"
                StrSql = StrSql & " ,oficdepto,codigopostal"
                StrSql = StrSql & " ,locnro,provnro,paisnro"
                StrSql = StrSql & " ,zonanro,partnro,barrio"
                StrSql = StrSql & " ,entrecalles)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & DomNro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!calle, 250)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!nro, 8)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!torre, 8)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!piso, 8)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!oficdepto, 8)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!codigopostal, 12)
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!locnro)
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!provnro)
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!paisnro)
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!zonanro)
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!partnro)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!barrio, 30)
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!entrecalles, 80)
                StrSql = StrSql & ")"
                ConnLE.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Domicilio en la BD LE"
                
                'Inserto el telefono - Personal
                If Not EsNulo(rs_Consult!telnro) Then
                    StrSql = "INSERT INTO telefono"
                    StrSql = StrSql & " (domnro,telnro,tipotel)"
                    StrSql = StrSql & " VALUES"
                    StrSql = StrSql & " (" & DomNro
                    StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!telnro, 20)
                    StrSql = StrSql & " ,1)"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Personal en la BD LE"
                End If
            
             
                'JPB - Inserto el telefono - Celular
                If Not EsNulo(rs_Consult!auxchr1) Then
                    StrSql = "INSERT INTO telefono"
                    StrSql = StrSql & " (domnro,telnro,tipotel)"
                    StrSql = StrSql & " VALUES"
                    StrSql = StrSql & " (" & DomNro
                    StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!auxchr1, 20)
                    StrSql = StrSql & " ,2)"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Celular en la BD LE"
                End If
            
            
            End If
            
            
        Else
            
            'Seteo el domicilio como personal
            StrSql = "UPDATE cabdom" & _
                     " SET domdefault = -1 " & _
                     " WHERE domnro = " & DomNro
             ' ConnLE.Execute StrSql, , adExecuteNoRecords
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Modificar el detalle
            StrSql = "UPDATE detdom SET "
            StrSql = StrSql & "  calle = " & CtrlNuloTXT(rs_Consult!calle, 250)
            StrSql = StrSql & " ,nro = " & CtrlNuloTXT(rs_Consult!nro, 8)
            StrSql = StrSql & " ,torre = " & CtrlNuloTXT(rs_Consult!torre, 8)
            StrSql = StrSql & " ,piso = " & CtrlNuloTXT(rs_Consult!piso, 8)
            StrSql = StrSql & " ,oficdepto = " & CtrlNuloTXT(rs_Consult!oficdepto, 8)
            StrSql = StrSql & " ,codigopostal = " & CtrlNuloTXT(rs_Consult!codigopostal, 12)
            StrSql = StrSql & " ,locnro = " & CtrlNuloNUM(rs_Consult!locnro)
            StrSql = StrSql & " ,provnro = " & CtrlNuloNUM(rs_Consult!provnro)
            StrSql = StrSql & " ,paisnro = " & CtrlNuloNUM(rs_Consult!paisnro)
            StrSql = StrSql & " ,zonanro = " & CtrlNuloNUM(rs_Consult!zonanro)
            StrSql = StrSql & " ,partnro = " & CtrlNuloNUM(rs_Consult!partnro)
            StrSql = StrSql & " ,barrio = " & CtrlNuloTXT(rs_Consult!barrio, 30)
            StrSql = StrSql & " ,entrecalles = " & CtrlNuloTXT(rs_Consult!entrecalles, 80)
            StrSql = StrSql & "  WHERE domnro = " & DomNro
            'ConnLE.Execute StrSql, , adExecuteNoRecords
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Modifico el domicilio"
            
            'Borro el telefono personal
            StrSql = "DELETE FROM telefono WHERE tipotel = 1 AND domnro = " & DomNro
            'ConnLE.Execute StrSql, , adExecuteNoRecords
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borro Telefono personal"
            
             'Borro el telefono celular
            StrSql = "DELETE FROM telefono WHERE tipotel = 2 AND domnro = " & DomNro
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borro Telefono Celular"
           
            
            'JPB - Inserto el telefono - Personal
            If Not EsNulo(rs_Consult!telnro) Then
                StrSql = "INSERT INTO telefono"
                StrSql = StrSql & " (domnro,telnro,tipotel)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & DomNro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!telnro, 20)
                StrSql = StrSql & " ,1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Personal " & rs_Consult!telnro
            End If
            
            'JPB - Inserto el telefono - Celular
            If Not EsNulo(rs_Consult!auxchr1) Then
                StrSql = "INSERT INTO telefono"
                StrSql = StrSql & " (domnro,telnro,tipotel)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & DomNro
                StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!auxchr1, 20)
                StrSql = StrSql & " ,2)"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Celular  " & rs_Consult!auxchr1
            End If
            
            
            
            '--------------- LE -----------------------------------------
            'JPB - 03/06/2011 - Cobtrola si usa LE y actualiza dicha base
            If (Usa_LE) And (Not Misma_BD) Then
              
               'Seteo el domicilio como personal
                StrSql = "UPDATE cabdom" & _
                        " SET domdefault = -1 " & _
                        " WHERE domnro = " & DomNro
                ConnLE.Execute StrSql, , adExecuteNoRecords
            
               'Modificar el detalle
                StrSql = "UPDATE detdom SET "
                StrSql = StrSql & "  calle = " & CtrlNuloTXT(rs_Consult!calle, 250)
                StrSql = StrSql & " ,nro = " & CtrlNuloTXT(rs_Consult!nro, 8)
                StrSql = StrSql & " ,torre = " & CtrlNuloTXT(rs_Consult!torre, 8)
                StrSql = StrSql & " ,piso = " & CtrlNuloTXT(rs_Consult!piso, 8)
                StrSql = StrSql & " ,oficdepto = " & CtrlNuloTXT(rs_Consult!oficdepto, 8)
                StrSql = StrSql & " ,codigopostal = " & CtrlNuloTXT(rs_Consult!codigopostal, 12)
                StrSql = StrSql & " ,locnro = " & CtrlNuloNUM(rs_Consult!locnro)
                StrSql = StrSql & " ,provnro = " & CtrlNuloNUM(rs_Consult!provnro)
                StrSql = StrSql & " ,paisnro = " & CtrlNuloNUM(rs_Consult!paisnro)
                StrSql = StrSql & " ,zonanro = " & CtrlNuloNUM(rs_Consult!zonanro)
                StrSql = StrSql & " ,partnro = " & CtrlNuloNUM(rs_Consult!partnro)
                StrSql = StrSql & " ,barrio = " & CtrlNuloTXT(rs_Consult!barrio, 30)
                StrSql = StrSql & " ,entrecalles = " & CtrlNuloTXT(rs_Consult!entrecalles, 80)
                StrSql = StrSql & "  WHERE domnro = " & DomNro
                ConnLE.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Modifico el domicilio en la BD LE"
            
                'Borro el telefono personal
                StrSql = "DELETE FROM telefono WHERE tipotel = 1 AND domnro = " & DomNro
                ConnLE.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Borro Telefono personal en la BD LE"
                 
                'Borro el telefono celular
                StrSql = "DELETE FROM telefono WHERE tipotel = 2 AND domnro = " & DomNro
                ConnLE.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Borro Telefono Celular en la BD LE"
            
                'Inserto el telefono - Particular
                If Not EsNulo(rs_Consult!telnro) Then
                    StrSql = "INSERT INTO telefono"
                    StrSql = StrSql & " (domnro,telnro,tipotel)"
                    StrSql = StrSql & " VALUES"
                    StrSql = StrSql & " (" & DomNro
                    StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!telnro, 20)
                    StrSql = StrSql & " ,1)"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Particular en la BD LE"
                End If
                
                'JPB - Inserto el telefono - Celular
                If Not EsNulo(rs_Consult!auxchr1) Then
                    StrSql = "INSERT INTO telefono"
                    StrSql = StrSql & " (domnro,telnro,tipotel)"
                    StrSql = StrSql & " VALUES"
                    StrSql = StrSql & " (" & DomNro
                    StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!auxchr1, 20)
                    StrSql = StrSql & " ,2)"
                    ConnLE.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Inserto el Telefono Celular en la BD LE"
                End If
                
           
            
            
            End If
            
              
                        
        End If
        
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 3, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 3"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
If rs_DomNro.State = adStateOpen Then rs_DomNro.Close
Set rs_DomNro = Nothing



MyCommitTrans

Exit Sub
E_Importar3:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar3"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans
    
    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 11, 3)
End Sub


Public Sub Importar4(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 4.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'        04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim TernroTemp As Long
Dim rs_Consult As New ADODB.Recordset
Dim StrSqlLE As String

On Error GoTo E_Importar4

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 4 para Postulante de ER " & ternroER & " con mail " & Mail
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_4_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Estudios Formales"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 4, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
        
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
        
    '---------------------------------------------------------------------------------------------------------------
    'Estudios Formales
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM cap_estformal WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Estudio Formal"
    
    'EAM - Verifica si tiene el LE
    If (Usa_LE) And (Not Misma_BD) Then
        StrSql = "DELETE FROM cap_estformal WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Estudio Formal en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_4 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            StrSql = "INSERT INTO cap_estformal"
            StrSql = StrSql & " (ternro,nivnro,titnro"
            StrSql = StrSql & " ,instnro,carredunro,capfecdes"
            StrSql = StrSql & " ,capfechas,capcomp,capanocur"
            StrSql = StrSql & " ,capcantmat,capestact,caprango"
            StrSql = StrSql & " ,capprom,capfutdesc)"
            'StrSql = StrSql & " ,capprom,capfutdesc"
            'StrSql = StrSql & " ,capactual,materias_aprob,porcentaje_carrera)"
            StrSql = StrSql & " VALUES"
            'EAM - Copia la Cabecera de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " (" & TernroRHPro
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!nivnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!titnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!instnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!carredunro)
            StrSql = StrSql & " ," & cambiaFecha(rs_Consult!capfecdes)
            StrSql = StrSql & " ," & cambiaFecha(rs_Consult!capfechas)
            'StrSql = StrSql & " ,NULL "
            StrSql = StrSql & " ," & CtrlNuloBOOL(rs_Consult!capcomp)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!capanocur)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!capcantmat)
            StrSql = StrSql & " ," & CtrlNuloBOOL(rs_Consult!capestact)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!caprango, 60)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!capprom, 30)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!capfutdesc, 60)
            'StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!capactual)
            'StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!materias_aprob)
            'StrSql = StrSql & " ," & rs_Consult!porcentaje_carrera
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Estudio Formal"
            
            If (Usa_LE) And (Not Misma_BD) Then
                'EAM - Arma la misma consulta para copiar en LE
                StrSqlLE = StrSqlLE & " (" & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!nivnro)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!titnro)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!instnro)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!carredunro)
                StrSqlLE = StrSqlLE & " ," & cambiaFecha(rs_Consult!capfecdes)
                StrSqlLE = StrSqlLE & " ," & cambiaFecha(rs_Consult!capfechas)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloBOOL(rs_Consult!capcomp)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!capanocur)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!capcantmat)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloBOOL(rs_Consult!capestact)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!caprango, 60)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!capprom, 30)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!capfutdesc, 60)
                'StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!capactual)
                'StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!materias_aprob)
                'StrSqlLE = StrSqlLE & " ," & rs_Consult!porcentaje_carrera
            
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Estudio Formal en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop
        
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 4, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 4"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar4:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar4"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans
    
    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 4, 2)
End Sub


Public Sub Importar5(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 5.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar5

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 5 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_5_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Estudios Informales"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 5, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Estudios Informales
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM cap_estinformal WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Estudio Informal"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM cap_estinformal WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Estudio Informal"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_5 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            StrSql = "INSERT INTO cap_estinformal"
            StrSql = StrSql & " (estinfdesabr,estinfdesext,tipcurnro"
            StrSql = StrSql & " ,instnro,ternro,estinffecha"
            StrSql = StrSql & " ,estinffechahasta)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloTXT(rs_Consult!estinfdesabr, 50)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!estinfdesext, 200)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!tipcurnro)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!instnro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " ," & cambiaFecha(rs_Consult!estinffecha)
            StrSql = StrSql & " ," & cambiaFecha(rs_Consult!estinffechahasta)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Estudio Informal"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para copiar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & cambiaFecha(rs_Consult!estinffecha)
                StrSqlLE = StrSqlLE & " ," & cambiaFecha(rs_Consult!estinffechahasta)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Estudio Informal en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 5, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 5"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar5:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar5"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 5, 2)
End Sub


Public Sub Importar6(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 6.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar6

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 6 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_6_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Especialidades (especemp)"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 6, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
        
    '---------------------------------------------------------------------------------------------------------------
    'Creacion de Especialidad
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM especemp WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Especialidad"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM especemp WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Especialidad en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_6 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            StrSql = "INSERT INTO especemp"
            StrSql = StrSql & " (eltananro,ternro,espnivnro)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!eltananro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!espnivnro)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Especialidad"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para copiar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!espnivnro)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Especialidad en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 6, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 6"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar6:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar6"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 6, 2)
    
End Sub

Public Sub Importar7(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 7.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar7

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 7 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_7_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Informacion (pos_infter)"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 7, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Informacion
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM pos_infter WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra informacion"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM pos_infter WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra informacion en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_7 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            StrSql = "INSERT INTO pos_infter"
            StrSql = StrSql & " (infnro,ternro)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!infnro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Informacion"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para copiar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Informacion en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada sin error
        Call MarcarSeccion(Mail, 7, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 7"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar7:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar7"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans
    
    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 7, 2)
    
End Sub


Public Sub Importar8(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 8.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar8

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 8 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_8_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Idiomas (emp_idi)"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 8, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
                
    '---------------------------------------------------------------------------------------------------------------
    'Idiomas
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM emp_idi WHERE empleado = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Idioma"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM emp_idi WHERE empleado = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Idiomas en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_8 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            StrSql = "INSERT INTO emp_idi"
            StrSql = StrSql & " (idinro,empleado,empidlee"
            StrSql = StrSql & " ,empidhabla,empidescr)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!idinro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!empidlee)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!empidhabla)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!empidescr)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Idioma"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para copiar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!empidlee)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!empidhabla)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!empidescr)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Idiomas en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 8, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 8"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing



MyCommitTrans

Exit Sub
E_Importar8:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar8"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 8, 2)
    
End Sub


Public Sub Importar9(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 9.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar9

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 9 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_9_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Habilidades (hab_ter)"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 9, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Estudios Informales
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM hab_ter WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Habilidades"
        
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM hab_ter WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Habilidades en la BD LE"
    End If
        
        
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_9 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            StrSql = "INSERT INTO hab_ter"
            StrSql = StrSql & " (habnro,ternro,sennro)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!habnro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!sennro)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Habilidad"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para copiar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!sennro)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Habilidad en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 9, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 9"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar9:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar9"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 9, 2)
    
End Sub


Public Sub Importar10(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 10.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------
 

Dim TernroRHPro As Long
Dim EmpAntRHPro As Long
Dim Area As Long
Dim Industria As Long
Dim TernroTemp As Long
Dim StrSqlLE As String
Dim EmpAntLE As Long

Dim rs_Consult As New ADODB.Recordset

On Error GoTo E_Importar10

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 10 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_10_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Empleos Anteriores (empant)"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 10, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
        
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
        
    '---------------------------------------------------------------------------------------------------------------
    'Empleos Anteriores
    '---------------------------------------------------------------------------------------------------------------
    
    'Recorro todos los empleos anteriores para borrar las estructuras asociadas
    StrSql = "SELECT empantnro"
    StrSql = StrSql & " FROM empant"
    StrSql = StrSql & " WHERE empleado = " & TernroRHPro
    OpenRecordset StrSql, rs_Consult
    Do While Not rs_Consult.EOF
        StrSql = "DELETE FROM empantestruc WHERE empantnro = " & rs_Consult!empantnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs_Consult.MoveNext
    Loop
    rs_Consult.Close
    Flog.writeline Espacios(Tabulador * 1) & "Borra Estruc de empleos anteriores"
    
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM empant WHERE empleado = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Empleos Anteriores"
    
    'Busco cuales son los alcances de la auditoria para industria y area
    Area = 0
    Industria = 0
    StrSql = "SELECT alcance_testr.tenro, alcance_testr.alteorden, tipoestructura.tedabr"
    StrSql = StrSql & " FROM alcance_testr"
    StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = alcance_testr.tenro"
    'StrSql = StrSql & " WHERE tanro = 15"
    StrSql = StrSql & " WHERE tanro = 18"
    StrSql = StrSql & " ORDER BY alteorden"
    OpenRecordset StrSql, rs_Consult
    Do While Not rs_Consult.EOF
        Select Case CLng(rs_Consult!alteorden)
            Case 1:
                Industria = rs_Consult!tenro
            Case 2:
                Area = rs_Consult!tenro
        End Select
        rs_Consult.MoveNext
    Loop
    
    If Area = 0 Then
        Flog.writeline Espacios(Tabulador * 2) & "No se encontro el alcance de la Estruct para Area (orden 1) de Empleos Anteriores."
    End If
    
    If Industria = 0 Then
        Flog.writeline Espacios(Tabulador * 2) & "No se encontro el alcance de la Estruct para Industria (orden 2) de Empleos Anteriores."
    End If
    
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Recorro todos los empleos anteriores para borrar las estructuras asociadas
        StrSql = "SELECT empantnro FROM empant" & _
                 " WHERE empleado = " & TernroTemp
        OpenRecordsetWithConn StrSql, rs_Consult, ConnLE
        
        Do While Not rs_Consult.EOF
            StrSql = "DELETE FROM empantestruc WHERE empantnro = " & rs_Consult!empantnro
            ConnLE.Execute StrSql, , adExecuteNoRecords
            
            rs_Consult.MoveNext
        Loop
        rs_Consult.Close
        Flog.writeline Espacios(Tabulador * 1) & "Borra Estruc de empleos anteriores en la BD LE"
        
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM empant WHERE empleado = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Empleos Anteriores en la BD LE"
    
    
        'Busco cuales son los alcances de la auditoria para industria y area
        Area = 0
        Industria = 0
        StrSql = "SELECT alcance_testr.tenro, alcance_testr.alteorden, tipoestructura.tedabr" & _
                 " FROM alcance_testr" & _
                 " INNER JOIN tipoestructura ON tipoestructura.tenro = alcance_testr.tenro" & _
                 " WHERE tanro = 18" & _
                 " ORDER BY alteorden"
        OpenRecordsetWithConn StrSql, rs_Consult, ConnLE
        
        Do While Not rs_Consult.EOF
            Select Case CLng(rs_Consult!alteorden)
                Case 1:
                    Industria = rs_Consult!tenro
                Case 2:
                    Area = rs_Consult!tenro
            End Select
            rs_Consult.MoveNext
        Loop
    
        If Area = 0 Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontro el alcance de la Estruct para Area (orden 1) de Empleos Anteriores en la BD LE."
        End If
    
        If Industria = 0 Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontro el alcance de la Estruct para Industria (orden 2) de Empleos Anteriores en la BD LE."
        End If
                
    End If
    
    
    'Creo todo los estudios que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_10 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            'Creo el empleo anterior
            StrSql = "INSERT INTO empant"
            StrSql = StrSql & " (empleado,empaleg,caunro"
            StrSql = StrSql & " ,empatel,empaini,empafin"
            StrSql = StrSql & " ,empatareas,emparemu,carnro"
            StrSql = StrSql & " ,empresa,empresaubica,emprefmail"
            StrSql = StrSql & " ,emprpuesto,empref,emprefpuesto,emparant)"
            StrSql = StrSql & " VALUES"
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
            
            StrSql = StrSql & " (" & TernroRHPro
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!nroempleado)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!caunro)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!telefonoref, 30)
            StrSql = StrSql & " ," & cambiaFecha(rs_Consult!Desde)
            StrSql = StrSql & " ," & cambiaFecha(rs_Consult!Hasta)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!tareasdese, 1000)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!remunaracion)
            StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!estrnro)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!empresa, 100)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!empresaubica, 100)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!mailref, 100)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!puesto, 100)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!referencia, 100)
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!puestoref, 100)
            StrSql = StrSql & " , 0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Empleos Anteriores"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para Insertar en LE
                StrSqlLE = StrSqlLE & " (" & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!nroempleado)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!caunro)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!telefonoref, 30)
                StrSqlLE = StrSqlLE & " ," & cambiaFecha(rs_Consult!Desde)
                StrSqlLE = StrSqlLE & " ," & cambiaFecha(rs_Consult!Hasta)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!tareasdese, 1000)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!remunaracion)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloNUM(rs_Consult!estrnro)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!empresa, 100)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!empresaubica, 100)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!mailref, 100)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!puesto, 100)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!referencia, 100)
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!puestoref, 100)
                StrSqlLE = StrSqlLE & " , 0)"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Empleos Anteriores en la BD LE"
                
                'Recupero el ultimo insertado en LE
                EmpAntLE = getLastIdentity(ConnLE, "empantnro")
                                                
            End If
            
            
            'Recupero el ultimo insertado
            EmpAntRHPro = getLastIdentity(objConn, "empantnro")
            
            'Creo Area (estructuras del alcance)
            If (CtrlNuloNUM(rs_Consult!areID) <> 0) And (Area <> 0) Then
                'Inserto el area
                StrSql = "INSERT INTO empantestruc"
                StrSql = StrSql & " (tenro,estrnro,empantnro)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & Area
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!areID)
                'EAM - Copia la parte de la consulta que es igual
                StrSqlLE = StrSql
                                
                StrSql = StrSql & " ," & EmpAntRHPro
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 2) & "Inserto Area"
                
                'EAM - 04-01-2010
                If (Usa_LE) And (Not Misma_BD) Then
                    'Arma la misma consulta para Insertar en LE
                    StrSqlLE = StrSqlLE & " ," & EmpAntLE
                    StrSqlLE = StrSqlLE & " )"
                    ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Inserto Area en la BD LE"
                End If
            End If
            
            'Creo Industria (estructuras del alcance)
            If (CtrlNuloNUM(rs_Consult!actividadnro) <> 0) And (Industria <> 0) Then
                'Inserto el area
                StrSql = "INSERT INTO empantestruc"
                StrSql = StrSql & " (tenro,estrnro,empantnro)"
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & " (" & Industria
                StrSql = StrSql & " ," & CtrlNuloNUM(rs_Consult!actividadnro)
                'EAM - Copia la parte de la consulta que es igual
                StrSqlLE = StrSql
                
                StrSql = StrSql & " ," & EmpAntRHPro
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 2) & "Inserto Industria"
                
                'EAM - 04-01-2010
                If (Usa_LE) And (Not Misma_BD) Then
                    StrSqlLE = StrSqlLE & " ," & EmpAntLE
                    StrSqlLE = StrSqlLE & " )"
                    ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "Inserto Industria en la BD LE"
                End If
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada sin error
        Call MarcarSeccion(Mail, 10, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 10"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar10:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar10"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans
    
    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 10, 2)
End Sub


Public Sub Importar11(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 11.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar11

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 11 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_11_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Notas (notas_ter)"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 11, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Notas
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM notas_ter WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Notas"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM notas_ter WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Notas en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_11 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
        
        Do While Not rs_Consult.EOF
            
            StrSql = "INSERT INTO notas_ter"
            StrSql = StrSql & " (tnonro ,ternro,notatxt)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!tnonro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
                
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " ," & CtrlNuloTXT(rs_Consult!notatxt, 1000)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Notas"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para Insertar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!notatxt, 1000)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Notas en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 11, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 11"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar11:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar11"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 11, 2)

End Sub


Public Sub Importar12(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 12.
' Autor      : Brzozowski Juan Pablo
' Fecha      : 26/05/2011
' Ultima Mod.:
'
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String
Dim Descripcion_SykesAcademy As String

On Error GoTo E_Importar12

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 11 para Postulante de ER " & ternroER & " con mail " & Mail
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_12_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Notas (notas_ter). Custom"


    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 12, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Notas
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM notas_ter WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Notas"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM notas_ter WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Notas en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_12 " & ternroER, rs_Consult, ExtConn
 
    If Not rs_Consult.EOF Then
       
        Do While Not rs_Consult.EOF
             
            'JPB - Armo la descripcion completa de la seccion Sykes Academy con todos los campos de la misma
            
            Descripcion_SykesAcademy = ""
            Descripcion_SykesAcademy = "Overall English Skills: " & rs_Consult!Overall & Chr(13)
            Descripcion_SykesAcademy = Descripcion_SykesAcademy & "How long have you studied English?: " & rs_Consult!howlongdesabr & " " & rs_Consult!How & Chr(13)
            Descripcion_SykesAcademy = Descripcion_SykesAcademy & "When and where did you study English?:  " & rs_Consult!wwstudy & Chr(13)
            Descripcion_SykesAcademy = Descripcion_SykesAcademy & "Questions: " & rs_Consult!questions
            
            
            
            StrSql = "INSERT INTO notas_ter"
            StrSql = StrSql & " (tnonro ,ternro,notatxt)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!tnonro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
                
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " , '" & Descripcion_SykesAcademy & "'"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Seccion Sykes Academy"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para Insertar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!notatxt, 1000)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Notas en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 12, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 12"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar12:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar12"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 12, 2)

End Sub




Public Sub Importar13(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de la seccion 13.
' Autor      : Brzozowski Juan Pablo
' Fecha      : 26/05/2011
' Ultima Mod.:
'
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim rs_Consult As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String
Dim Descripcion_SykesAcademy As String

On Error GoTo E_Importar13

MyBeginTrans


    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 1) & "Importacion Seccion 11 para Postulante de ER " & ternroER & " con mail " & Mail

    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_13_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Notas (notas_ter). Custom"

    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 13, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Notas
    '---------------------------------------------------------------------------------------------------------------
    
    'Borro todo lo que existe en RHPro
    StrSql = "DELETE FROM notas_ter WHERE ternro = " & TernroRHPro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Borra Notas"
    
    'EAM - Verifica si tiene instalado LE
    If (Usa_LE) And (Not Misma_BD) Then
        'Borro todo lo que existe en LE
        StrSql = "DELETE FROM notas_ter WHERE ternro = " & TernroTemp
        ConnLE.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Borra Notas en la BD LE"
    End If
    
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_13 " & ternroER, rs_Consult, ExtConn
 
    If Not rs_Consult.EOF Then
       
        Do While Not rs_Consult.EOF
             
            'JPB - Armo la descripcion completa de la seccion Sykes Academy con todos los campos de la misma
            
            Descripcion_SykesAcademy = ""
            Descripcion_SykesAcademy = "Overall English Skills: " & rs_Consult!Overall & Chr(13)
            Descripcion_SykesAcademy = Descripcion_SykesAcademy & "How long have you studied English?: " & rs_Consult!howlongdesabr & " " & rs_Consult!How & Chr(13)
            Descripcion_SykesAcademy = Descripcion_SykesAcademy & "When and where did you study English?:  " & rs_Consult!wwstudy & Chr(13)
            Descripcion_SykesAcademy = Descripcion_SykesAcademy & "Questions: " & rs_Consult!questions
            
            
            
            StrSql = "INSERT INTO notas_ter"
            StrSql = StrSql & " (tnonro ,ternro,notatxt)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & " (" & CtrlNuloNUM(rs_Consult!tnonro)
            'EAM - Copia la parte de la consulta que es igual
            StrSqlLE = StrSql
                
            StrSql = StrSql & " ," & TernroRHPro
            StrSql = StrSql & " , '" & Descripcion_SykesAcademy & "'"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto Seccion Sykes Academy"
            
            'EAM - 04-01-2010
            If (Usa_LE) And (Not Misma_BD) Then
                'Arma la misma consulta para Insertar en LE
                StrSqlLE = StrSqlLE & " ," & TernroTemp
                StrSqlLE = StrSqlLE & " ," & CtrlNuloTXT(rs_Consult!notatxt, 1000)
                StrSqlLE = StrSqlLE & " )"
                ConnLE.Execute StrSqlLE, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Notas en la BD LE"
            End If
            
            rs_Consult.MoveNext
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 13, -1)
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos para la seccion 13"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing


MyCommitTrans

Exit Sub
E_Importar13:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar13"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 13, 2)

End Sub



Public Sub Importar100(ByVal Mail As String, ByVal ternroER As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento encargado de la migracion de las postulaciones.
' Autor      : Martin Ferraro
' Fecha      : 27/11/2009
' Ultima Mod.:
'       04/01/2010 - Margiotta Emanuel - Verifica si esta instalado el LE y Copia los datos en la BD Temp si existe
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim TernroRHPro As Long
Dim ReqBusNro As Long
Dim BusqFormal As Boolean
Dim rs_Consult As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim TernroTemp As Long
Dim StrSqlLE As String

On Error GoTo E_Importar100

MyBeginTrans


    Flog.writeline
 
    FSinc.writeline
    FSinc.writeline Espacios(Tabulador * 1) & "SP de ER: REC_MIGRA_SP_100_POS"
    FSinc.writeline Espacios(Tabulador * 2) & "Busca todo los nuevo en ER que no existe en RHPRO y lo crea"
    FSinc.writeline Espacios(Tabulador * 3) & "Busquedas (pos_busqueda)"
    FSinc.writeline Espacios(Tabulador * 3) & "Requerimientos (pos_reqbus)"
    FSinc.writeline Espacios(Tabulador * 3) & "Psotulantes asoiados a Requerimientos (pos_terreqbusER)"
 
    '---------------------------------------------------------------------------------------------------------------
    'Busco el postulante en Rhpro
    '---------------------------------------------------------------------------------------------------------------
    TernroRHPro = BuscarTerceroXMail(Mail)
    If TernroRHPro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro en Rhpro un postulante con el Mail " & Mail
        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 100, 2)
        MyCommitTrans
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Modificando datos del tercero " & TernroRHPro
    
    'EAM- Obtiene el ternro de la BD LE
    TernroTemp = BuscarTerceroTempXMail(Mail)
    
    '---------------------------------------------------------------------------------------------------------------
    'Postulaciones a busquedas
    '---------------------------------------------------------------------------------------------------------------
    
    'Creo todo lo que exista en ER sea nuevo o no
    OpenRecordsetExt "EXEC REC_MIGRA_SP_100 " & ternroER, rs_Consult, ExtConn
    If Not rs_Consult.EOF Then
  
        Do While Not rs_Consult.EOF
 
            
            ReqBusNro = 0
            
            'Verifico que exista la busqueda
            StrSql = "select * from pos_busqueda where busnro = " & CtrlNuloNUM(rs_Consult!id_rhpro)
            OpenRecordset StrSql, rs_Datos
            If rs_Datos.EOF Then
                Flog.writeline Espacios(Tabulador * 1) & "No existe una busqueda con codigo " & CtrlNuloNUM(rs_Consult!id_rhpro) & " en RHPRO"
                GoTo SgtDato
            End If
            
            If rs_Datos!busformal = "-1" Then
                Flog.writeline Espacios(Tabulador * 1) & "Migrando Postulacion de Busqueda FORMAL " & CtrlNuloNUM(rs_Consult!id_rhpro)
                BusqFormal = True
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Migrando Postulacion de Busqueda FORMAL " & CtrlNuloNUM(rs_Consult!id_rhpro)
                BusqFormal = False
            End If
            
            
            'Verifico la relacion pos_reqbus
            StrSql = "select reqbusnro, reqpernro from pos_reqbus where busnro = " & CtrlNuloNUM(rs_Consult!id_rhpro) & " ORDER BY reqpernro"
            OpenRecordset StrSql, rs_Datos
            If Not rs_Datos.EOF Then
                'Recupero el existente
                ReqBusNro = rs_Datos!ReqBusNro
            Else
                If BusqFormal Then
                    Flog.writeline Espacios(Tabulador * 1) & "Busqueda FORMAL Sin requerimiento. No se migra la postulacion"
                    GoTo SgtDato
                Else
                    'Creo uno con requerimiento en nulo
                    StrSql = "INSERT INTO pos_reqbus "
                    StrSql = StrSql & " (busnro)"
                    StrSql = StrSql & " VALUES("
                    StrSql = StrSql & " " & CtrlNuloNUM(rs_Consult!id_rhpro)
                    StrSql = StrSql & " )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'Recupero el insertado
                    ReqBusNro = getLastIdentity(objConn, "pos_reqbus")
                End If
            End If
            
            If ReqBusNro <> 0 Then
                'Busco si ya esta postulado
                StrSql = "select *"
                StrSql = StrSql & " From pos_terreqbusER"
                StrSql = StrSql & " Where ternro = " & TernroRHPro
                StrSql = StrSql & " AND reqbusnro = " & ReqBusNro
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    Flog.writeline Espacios(Tabulador * 1) & "El tercero ya se encuentra postulado. No se modifican datos."
                    rs_Datos.Close
                Else
                    rs_Datos.Close
                    StrSql = "INSERT INTO pos_terreqbusER"
                    StrSql = StrSql & " (ternro,reqbusnro,conf,migrado)"
                    StrSql = StrSql & " VALUES("
                    StrSql = StrSql & " " & TernroRHPro
                    StrSql = StrSql & " , " & ReqBusNro
                    StrSql = StrSql & " ,0,0"
                    StrSql = StrSql & " )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Se creo la postulacion del tercero."
                End If
            End If
 


SgtDato:   'JPB - Marco la busqueda/postulante como Activo.
            StrSql = " UPDATE estado_postulaciones SET activo = 1 WHERE reqpernro =  " & rs_Consult!reqpernro
            ExtConn.Execute StrSql, , adExecuteNoRecords
            
            rs_Consult.MoveNext
        
        Loop

        'Marco la seccion como migrada con error
        Call MarcarSeccion(Mail, 100, -1)
               
        'Si LE esta instalado y esta en otra BD copia los datos
        If (Usa_LE) And (Not Misma_BD) Then
            rs_Consult.MoveFirst
            Do While Not rs_Consult.EOF
            
                ReqBusNro = 0
                
                'Verifico que exista la busqueda
                StrSql = "select * from pos_busqueda where busnro = " & CtrlNuloNUM(rs_Consult!id_rhpro)
                OpenRecordsetWithConn StrSql, rs_Datos, ConnLE
                
                If rs_Datos.EOF Then
                    Flog.writeline Espacios(Tabulador * 1) & "No existe una busqueda con codigo " & CtrlNuloNUM(rs_Consult!id_rhpro) & " en el LE"
                    GoTo SgtDato2
                End If
                
                If rs_Datos!busformal = "-1" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Migrando Postulacion de Busqueda FORMAL " & CtrlNuloNUM(rs_Consult!id_rhpro) & " en la BD LE"
                    BusqFormal = True
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Migrando Postulacion de Busqueda FORMAL " & CtrlNuloNUM(rs_Consult!id_rhpro) & " en la BD LE"
                    BusqFormal = False
                End If
                
                
                'Verifico la relacion pos_reqbus
                StrSql = "select reqbusnro, reqpernro from pos_reqbus where busnro = " & CtrlNuloNUM(rs_Consult!id_rhpro) & " ORDER BY reqpernro"
                OpenRecordsetWithConn StrSql, rs_Datos, ConnLE
                If Not rs_Datos.EOF Then
                    'Recupero el existente
                    ReqBusNro = rs_Datos!ReqBusNro
                Else
                    If BusqFormal Then
                        Flog.writeline Espacios(Tabulador * 1) & "Busqueda FORMAL Sin requerimiento. No se migra la postulacion en la BD LE"
                        GoTo SgtDato2
                    Else
                        'Creo uno con requerimiento en nulo
                        StrSql = "INSERT INTO pos_reqbus"
                        StrSql = StrSql & " (busnro)"
                        StrSql = StrSql & " VALUES("
                        StrSql = StrSql & " " & CtrlNuloNUM(rs_Consult!id_rhpro)
                        StrSql = StrSql & " )"
                        ConnLE.Execute StrSql, , adExecuteNoRecords
                        
                        'Recupero el insertado
                        ReqBusNro = getLastIdentity(ConnLE, "pos_reqbus")
                    End If
                End If
                
                If ReqBusNro <> 0 Then
                    'Busco si ya esta postulado
                    StrSql = "select *"
                    StrSql = StrSql & " From pos_terreqbusER"
                    StrSql = StrSql & " Where ternro = " & TernroTemp
                    StrSql = StrSql & " AND reqbusnro = " & ReqBusNro
                    OpenRecordsetWithConn StrSql, rs_Datos, ConnLE
                    If Not rs_Datos.EOF Then
                        Flog.writeline Espacios(Tabulador * 1) & "El tercero ya se encuentra postulado. No se modifican datos en la BD LE."
                        rs_Datos.Close
                    Else
                        rs_Datos.Close
                        StrSql = "INSERT INTO pos_terreqbusER"
                        StrSql = StrSql & " (ternro,reqbusnro,conf,migrado)"
                        StrSql = StrSql & " VALUES("
                        StrSql = StrSql & " " & TernroTemp
                        StrSql = StrSql & " , " & ReqBusNro
                        StrSql = StrSql & " ,0,0"
                        StrSql = StrSql & " )"
                        ConnLE.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Se creo la postulacion del tercero en la BD LE."
                    End If
                End If
                
SgtDato2:        rs_Consult.MoveNext
            
            Loop
            
        End If
        
    Else
    
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos de postulaciones"
    
    End If
    rs_Consult.Close
    
    
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing


MyCommitTrans

Exit Sub
E_Importar100:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Importar100"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    MyRollbackTrans

    'Marco la seccion como migrada con error
    Call MarcarSeccion(Mail, 11, 2)

End Sub




