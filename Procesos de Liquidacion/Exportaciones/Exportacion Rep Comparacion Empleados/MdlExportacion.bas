Attribute VB_Name = "MdlExportacion"
Option Explicit

'*************************************************************************************

Global Const Version = "1.01"
Global Const FechaModificacion = "21/10/2009"
Global Const UltimaModificacion = " " 'ML - Encriptacion de string connection

'*************************************************************************************

Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser As String
Global Fecha As Date
Global Hora As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa As New ADODB.Recordset
Global rs_tipocod As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String
Global SeparadorDecimales As String
Global totalImporte As Double
Global Total As Single
Global UltimaLeyenda As String

Dim fExport

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : Fernando Favre
' Fecha      : 26/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

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

    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    Nombre_Arch = PathFLog & "ExpReporteCompEmpleados" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
        
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    On Error GoTo ME_Local
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call Generacion(bprcparam)
'        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
Exit Sub
ME_Local:
    HuboError = True
        
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        If InStr(1, Err.Description, "ODBC") > 0 Then
            'Fue error de Consulta de SQL
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
            Flog.writeline
        End If
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprcestado = 'Error General' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin:
End Sub

Public Sub Generacion(ByVal Proceso As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo con distribucion por legajo
' Autor      : Fernando Favre
' Fecha      : 26/07/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0


Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim strLinea As String
Dim Aux_Linea As String
Dim Aux_Encabezado As String
Dim cadena As String
Dim Aux_Str As String
Dim Nro As Long
Dim SeparadorCampos
Dim I As Integer

Dim Encabezado As Boolean
Dim Corte As Boolean
Dim pliqmesanio1
Dim pliqmesanio2
Dim pliqdesc1
Dim pliqdesc2
Dim pronro1
Dim prodesc1
Dim pronro2
Dim prodesc2

'Auxiliares
Dim Apnom As String

Dim Vacio As String
Dim FechayHora As String
Dim Aux_ArchivoGenerado As String
Dim MaxLongitud As Long

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset


    StrSql = "SELECT * "
    StrSql = StrSql & "FROM rep_comp_empl WHERE bpronro = " & Proceso
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        
        FechayHora = Format(rs2!Fecha, "dd/mm/yyyy") & " - " & rs2!Hora
        
        pliqmesanio1 = rs2!pliqmesanio1
        pliqmesanio2 = rs2!pliqmesanio2
    
        pliqdesc1 = rs2!pliqdesc1
        pliqdesc2 = rs2!pliqdesc2
    
        pronro1 = rs2!pronro1
        pronro2 = rs2!pronro2
        
        StrSql = "SELECT prodesc FROM proceso WHERE pronro IN (" & pronro1 & ")"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        Do Until rs.EOF
            prodesc1 = prodesc1 & rs!prodesc & " - "
            rs.MoveNext
        Loop
        If rs.State = adStateOpen Then rs.Close
        
        If prodesc1 <> "" Then
            prodesc1 = Left(prodesc1, Len(prodesc1) - 3)
        End If
        
        StrSql = "SELECT prodesc FROM proceso WHERE pronro IN (" & pronro2 & ")"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        Do Until rs.EOF
            prodesc2 = prodesc2 & rs!prodesc & " - "
            rs.MoveNext
        Loop
        If rs.State = adStateOpen Then rs.Close
        
        If prodesc2 <> "" Then
            prodesc2 = Left(prodesc2, Len(prodesc2) - 3)
        End If
    
    End If
    
    If rs2.State = adStateOpen Then rs2.Close


    Vacio = ""
    
    'Archivo de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas)
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 254"
    OpenRecordset StrSql, rs_Modelo
    If Not rs_Modelo.EOF Then
        If Not IsNull(rs_Modelo!modarchdefault) Then
            Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
        End If
        SeparadorDecimales = rs_Modelo!modsepdec
        SeparadorCampos = rs_Modelo!modseparador
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
    End If
    
    
    'Busco las lineas a exportar
    StrSql = "SELECT * "
    StrSql = StrSql & "FROM rep_comp_empl_det WHERE bpronro = " & Proceso & " "
    StrSql = StrSql & "ORDER BY tconcepto, conccod, empleg "
    
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    CConceptosAProc = rs.RecordCount
    If CConceptosAProc = 0 Then
        CConceptosAProc = 1
        Flog.writeline Espacios(Tabulador * 1) & " No hay lineas para procesar "
    Else
        Flog.writeline Espacios(Tabulador * 1) & " Lineas a procesar " & CConceptosAProc
    End If
    IncPorc = (99 / CConceptosAProc)
    
    'Procesamiento
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
    Else
        'Seteo el nombre del archivo generado
        Aux_ArchivoGenerado = "rep_comp_empleados_" & Left(Trim(pliqdesc1), 10) & "_" & Left(Trim(pliqdesc2), 10) & ".csv"
        
        Archivo = Directorio & "\rep_comp_empleados_" & Left(Trim(pliqdesc1), 10) & "_" & Left(Trim(pliqdesc2), 10) & ".csv"
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        'Activo el manejador de errores
        On Error Resume Next
        Set fExport = fs.CreateTextFile(Archivo, True)
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
            Set carpeta = fs.CreateFolder(Directorio)
            Set fExport = fs.CreateTextFile(Archivo, True)
        End If
        'desactivo el manejador de errores
        On Error GoTo 0
    End If
    
    
    ' Comienzo la transaccion
    MyBeginTrans
    
    On Error GoTo ME_Local
    
    '------------------------------------------------------------------------
    ' Genero el detalle de la exportacion
    '------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del reporte"
    Flog.writeline

'Imprimo el encabezado
'calculo el acho de la fila del encabezado
    Aux_Linea = "Tipo Concepto"
    Aux_Linea = Aux_Linea & "Código"
    Aux_Linea = Aux_Linea & "Concepto"
    Aux_Linea = Aux_Linea & "Empleado"
    Aux_Linea = Aux_Linea & "Apellido y Nombre"
    Aux_Linea = Aux_Linea & "Monto " & pliqmesanio1
    Aux_Linea = Aux_Linea & "Monto " & pliqmesanio2
    Aux_Linea = Aux_Linea & "Diferencia Monto"
    Aux_Linea = Aux_Linea & "Porc. Monto"
    Aux_Linea = Aux_Linea & "Cantidad " & pliqmesanio1
    Aux_Linea = Aux_Linea & "Cantidad " & pliqmesanio2
    Aux_Linea = Aux_Linea & "Dif. Cant."
    Aux_Linea = Aux_Linea & "Porc. Cant."
    Aux_Encabezado = Aux_Linea
    MaxLongitud = Len(Aux_Encabezado) + 12

    'Aux_Linea = Espacios(MaxLongitud - Len(Aux_ArchivoGenerado)) & Aux_ArchivoGenerado
    Aux_Linea = Aux_ArchivoGenerado
    fExport.writeline Aux_Linea
    'Aux_Linea = Espacios(MaxLongitud - Len(FechayHora)) & FechayHora
    Aux_Linea = FechayHora
    fExport.writeline Aux_Linea

'Imprimo el titulo
    Aux_Linea = "COMPARATIVO"
    fExport.writeline Aux_Linea
    Aux_Linea = "Totales de Liquidación detallado por Empleados"
    fExport.writeline Aux_Linea


'Subtitulo
    Aux_Linea = pliqdesc1 & ": " & prodesc1
    fExport.writeline Aux_Linea
    Aux_Linea = pliqdesc2 & ": " & prodesc2
    fExport.writeline Aux_Linea


'Encabezado
    Aux_Linea = "Tipo Concepto"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Código"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Concepto"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Empleado"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Apellido y Nombre"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Monto " & pliqmesanio1
    Aux_Linea = Aux_Linea & SeparadorCampos & "Monto " & pliqmesanio2
    Aux_Linea = Aux_Linea & SeparadorCampos & "Diferencia Monto"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Porc. Monto"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Cantidad " & pliqmesanio1
    Aux_Linea = Aux_Linea & SeparadorCampos & "Cantidad " & pliqmesanio2
    Aux_Linea = Aux_Linea & SeparadorCampos & "Dif. Cant."
    Aux_Linea = Aux_Linea & SeparadorCampos & "Porc. Cant."
    Aux_Encabezado = Aux_Linea
    fExport.writeline Aux_Encabezado
    Flog.writeline Espacios(Tabulador * 2) & Aux_Encabezado


'Comienzo ciclo principal
    Do While Not rs.EOF
        Apnom = rs!terape
        If rs!terape2 <> "" Then
            Apnom = Apnom & " " & rs!terape2
        End If
        'FGZ - 03/08/2005
        'Apnom = Apnom & ", " & rs!ternom
        Apnom = Apnom & " " & rs!ternom
        If rs!ternom2 <> "" Then
            Apnom = Apnom & " " & rs!ternom2
        End If
  
        Aux_Linea = rs!TConcepto
        Aux_Linea = Aux_Linea & SeparadorCampos & rs!Conccod
        Aux_Linea = Aux_Linea & SeparadorCampos & rs!concabr
        Aux_Linea = Aux_Linea & SeparadorCampos & rs!empleg
        Aux_Linea = Aux_Linea & SeparadorCampos & Apnom
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!emplmonto1 <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!emplmonto1, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!emplmonto2 <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!emplmonto2, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!difmontoempl <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!difmontoempl, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!porcmontoempl <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!porcmontoempl, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!emplcant1 <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!emplcant1, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!emplcant2 <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!emplcant2, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!difcantempl <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!difcantempl, "###0.00")
        End If
        Aux_Linea = Aux_Linea & SeparadorCampos
        If rs!porccantempl <> 0 Then
            Aux_Linea = Aux_Linea & Format(rs!porccantempl, "###0.00")
        End If
        'FGZ - 03/08/2005
        'Aux_Linea = Replace(Aux_Linea, ",", "")
        'Aux_Linea = Replace(Aux_Linea, ".", ",")
        fExport.writeline Aux_Linea
        
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                
        'Siguiente proceso
        rs.MoveNext
        
    Loop
    
    'Fin de la transaccion
    MyCommitTrans
    
    
Fin:
    'Cierro el archivo creado
    fExport.Close
    
    If rs.State = adStateOpen Then rs.Close
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    Set rs = Nothing
    Set rs_Modelo = Nothing
    
    Exit Sub
ME_Local:
    HuboError = True
    MyRollbackTrans

        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        If InStr(1, Err.Description, "ODBC") > 0 Then
            'Fue error de Consulta de SQL
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
            Flog.writeline
        End If
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprcestado = 'Error General' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin:
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
'Orden de los parametros


'Call Generacion(Proceso, Detalle)
End Sub





