Attribute VB_Name = "MdlExportacion"
Option Explicit

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
' Autor      : FGZ
' Fecha      : 07/09/2004
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Exp_Reporte_Comparativo" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline Espacios(Tabulador * 0) & "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    On Error GoTo ME_Local
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 83 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
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

Public Sub Generacion(ByVal Proceso As Long, ByVal tipo As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo con distribucion por legajo
' Autor      : FGZ
' Fecha      : 22/04/2005
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
Dim i As Integer

Dim ProNro1 As String
Dim ProNro2 As String
Dim ProDesc1 As String
Dim ProDesc2 As String
Dim Cambios As Boolean

Dim Encabezado As Boolean
Dim Corte As Boolean
Dim Acunro_Ant As Long
Dim Concnro_Ant As Long
Dim Desc1 As String
Dim Desc2 As String

'Auxiliares
Dim Acudesabrant
Dim Acumonto1ant
Dim Acumonto2ant
Dim Difmontoacuant
Dim Porcmontoacuant
Dim Acucant1ant
Dim Acucant2ant
Dim Difcantacuant
Dim Porccantacuant

Dim Conccodant
Dim Concabrant
Dim Concmonto1ant
Dim Concmonto2ant
Dim Difmontoconcant
Dim Porcmontoconcant
Dim Conccant1ant
Dim Conccant2ant
Dim Difcantconcant
Dim Porccantconcant

Dim Empmonto1ant
Dim Empmonto2ant
Dim Difmontoempant
Dim Porcmontoempant
Dim Empcant1ant
Dim Empcant2ant
Dim Difcantempant
Dim Porccantempant
Dim Apnom As String

Dim cambiaacum As Boolean
Dim cambiaConc As Boolean
Dim Vacio As String
Dim Nro_Col As Integer

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim aux As String


Dim TipoHoja As String
Dim OrientacionVertical As Boolean
Dim LineasPorHoja As Integer
Dim LineaActual As Long

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
    Select Case tipo
    Case 0:
        aux = "Acumulador"
        StrSql = "SELECT DISTINCT acunro, acudesc, pliqmesanio1, pronro1, acumonto1, acucant1, pliqmesanio2, pliqdesc1, "
        StrSql = StrSql & "pronro2, acumonto2, acucant2, difmontoacu, porcmontoacu, difcantacu, porccantacu, pliqdesc2 "
        StrSql = StrSql & "FROM rep_comp_acum WHERE bpronro = " & Proceso & " "
        StrSql = StrSql & "ORDER BY acudesc"
    Case 1:
        aux = "Concepto"
        StrSql = "SELECT DISTINCT acunro, acudesc, pliqmesanio1, pronro1, concnro, conccod, concabr, acumonto1, acucant1, pliqdesc1, "
        StrSql = StrSql & "concmonto1, conccant1, pliqmesanio2, pronro2, acumonto2, acucant2, concmonto2, conccant2, difmontoacu, "
        StrSql = StrSql & "porcmontoacu, difcantacu, porccantacu, difmontoconc, porcmontoconc, difcantconc, porccantconc, pliqdesc2 "
        StrSql = StrSql & "FROM rep_comp_acum WHERE bpronro = " & Proceso & " "
        StrSql = StrSql & "ORDER BY acudesc, conccod"
    Case 2:
        aux = "Empleado"
        StrSql = "SELECT * FROM  rep_comp_acum "
        StrSql = StrSql & " WHERE bpronro =" & Proceso
        StrSql = StrSql & " ORDER BY acudesc, conccod, empleg "
    Case Else
        aux = "default"
        StrSql = "SELECT * FROM  rep_comp_acum "
        StrSql = StrSql & " WHERE bpronro =" & Proceso
        StrSql = StrSql & " ORDER BY acudesc, conccod, empleg "
    End Select
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
    IncPorc = (100 / CConceptosAProc)
    
    'Procesamiento
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
    Else
        'Seteo el nombre del archivo generado
        
        Archivo = Directorio & "\rep_comparativo_x_" & aux & "_" & Left(Trim(rs!pliqdesc1), 10) & "_" & Left(Trim(rs!pliqdesc2), 10) & ".csv"
        
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
    
    
'FGZ - seteo de las dimensiones de pagina
TipoHoja = "LEGAL"
OrientacionVertical = False
If OrientacionVertical Then
    Select Case TipoHoja
    Case "A4":
        LineasPorHoja = 63
    Case "CARTA":
        LineasPorHoja = 59
    Case "LEGAL", "OFICIO":
        LineasPorHoja = 76
    Case Else
        LineasPorHoja = 63
    End Select
Else
    Select Case TipoHoja
    Case "A4", "OFICIO", "CARTA":
        LineasPorHoja = 46
    Case "LEGAL"
        LineasPorHoja = 47
    End Select
End If
LineaActual = 0

    
    
    ' Comienzo la transaccion
    MyBeginTrans
    
    On Error GoTo ME_Local
    
    '------------------------------------------------------------------------
    ' Genero el detalle de la exportacion
    '------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del reporte"
    Flog.writeline

'Imprimo el titulo
    Aux_Linea = "COMPARATIVO"
    fExport.writeline Aux_Linea
    LineaActual = LineaActual + 1
    Aux_Linea = "Totales de Liquidación detallado por Empleado"
    fExport.writeline Aux_Linea
    LineaActual = LineaActual + 1

'Subtitulo
    ProNro1 = rs!ProNro1
    ProNro2 = rs!ProNro2
    StrSql = "SELECT pronro, prodesc FROM proceso WHERE pronro IN (" & ProNro1 & ") ORDER BY prodesc"
    If rs2.State = adStateOpen Then rs2.Close
    OpenRecordset StrSql, rs2
    ProDesc1 = ""
    Do Until rs2.EOF
        ProDesc1 = ProDesc1 & rs2!prodesc & " (" & rs2!pronro & ") - "
        rs2.MoveNext
    Loop
    ProDesc1 = Left(ProDesc1, Len(ProDesc1) - 3)
    StrSql = "SELECT pronro, prodesc FROM proceso WHERE pronro IN (" & ProNro2 & ") ORDER BY prodesc"
    If rs2.State = adStateOpen Then rs2.Close
    OpenRecordset StrSql, rs2
    ProDesc2 = ""
    Do Until rs2.EOF
        ProDesc2 = ProDesc2 & rs2!prodesc & " (" & rs2!pronro & ") - "
        rs2.MoveNext
    Loop
    ProDesc2 = Left(ProDesc2, Len(ProDesc2) - 3)

    Aux_Linea = Trim(rs!pliqdesc1) & ":" & ProDesc1
    fExport.writeline Aux_Linea
    LineaActual = LineaActual + 1
    Aux_Linea = Trim(rs!pliqdesc2) & ":" & ProDesc2
    fExport.writeline Aux_Linea
    LineaActual = LineaActual + 1

'Encabezado
    Aux_Linea = "Acumulador"
    Select Case tipo
    Case 0:
    Case 1:
        Aux_Linea = Aux_Linea & SeparadorCampos & "Código"
        Aux_Linea = Aux_Linea & SeparadorCampos & "Concepto"
    Case "2":
        Aux_Linea = Aux_Linea & SeparadorCampos & "Legajo"
        Aux_Linea = Aux_Linea & SeparadorCampos & "Apellido y Nombre"
    End Select
    Aux_Linea = Aux_Linea & SeparadorCampos & "Monto " & rs!pliqmesanio1
    Aux_Linea = Aux_Linea & SeparadorCampos & "Monto " & rs!pliqmesanio2
    Aux_Linea = Aux_Linea & SeparadorCampos & "Deferencia Monto"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Porc. Monto"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Cantidad " & rs!pliqmesanio1
    Aux_Linea = Aux_Linea & SeparadorCampos & "Cantidad " & rs!pliqmesanio2
    Aux_Linea = Aux_Linea & SeparadorCampos & "Deferencia Cantidad"
    Aux_Linea = Aux_Linea & SeparadorCampos & "Porc. Cantidad"
    Aux_Encabezado = Aux_Linea
    fExport.writeline Aux_Encabezado
    Flog.writeline Espacios(Tabulador * 2) & Aux_Encabezado
    LineaActual = LineaActual + 1
    
    'Seteo los valores de corte
    Encabezado = False
    Cambios = True
    Corte = False
    Acunro_Ant = rs!acunro
    If tipo <> 0 Then
        Concnro_Ant = rs!concnro
    End If

    Select Case tipo
        Case "0":
            Nro_Col = 9
            'l_titulo = "Acumulador"
        Case "1":
            cambiaacum = True
            cambiaConc = True
            Nro_Col = 11
            'l_titulo = "Concepto"
        Case "2":
            Nro_Col = 11
            'l_titulo = "Empleado"
     End Select



'Comienzo ciclo principal
    Do While Not rs.EOF
        'Guarda los valores
        Acudesabrant = rs!acudesc
        Acumonto1ant = rs!acumonto1
        Acumonto2ant = rs!acumonto2
        Difmontoacuant = rs!difmontoacu
        Porcmontoacuant = rs!porcmontoacu
        Acucant1ant = rs!acucant1
        Acucant2ant = rs!acucant2
        Difcantacuant = rs!difcantacu
        Porccantacuant = rs!porccantacu
        
        If tipo = 1 Or tipo = 2 Then
            Conccodant = rs!Conccod
            Concabrant = rs!concabr
            Concmonto1ant = rs!concmonto1
            Concmonto2ant = rs!concmonto2
            Difmontoconcant = rs!difmontoconc
            Porcmontoconcant = rs!porcmontoconc
            Conccant1ant = rs!conccant1
            Conccant2ant = rs!conccant2
            Difcantconcant = rs!difcantconc
            Porccantconcant = rs!porccantconc
        End If
        
        If tipo = 2 Then
            Empmonto1ant = rs!empmonto1
            Empmonto2ant = rs!empmonto2
            Difmontoempant = rs!difmontoemp
            Porcmontoempant = rs!porcmontoemp
            Empcant1ant = rs!empcant1
            Empcant2ant = rs!empcant2
            Difcantempant = rs!difcantemp
            Porccantempant = rs!porccantemp
        End If
        
        Desc1 = rs!acudesc
        If tipo <> 0 Then
            Desc2 = "--- " & rs!Conccod & " " & rs!concabr
        End If
        
        If Cambios Then
            'titulos
            If cambiaacum And tipo <> 0 Then
                fExport.writeline Desc1
                Flog.writeline Espacios(Tabulador * 2) & "Acumulador : " & Desc1
                LineaActual = LineaActual + 1
            End If
            If cambiaConc And tipo = 2 Then
                fExport.writeline Desc2
                Flog.writeline Espacios(Tabulador * 3) & "Concepto : " & Desc2
                LineaActual = LineaActual + 1
            End If
            Acunro_Ant = rs!acunro
            If tipo <> 0 Then
                Concnro_Ant = rs!concnro
            End If
        End If
        
        'Detalles
        Select Case CStr(tipo)
            Case "0":
                Aux_Linea = rs!acudesc
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!acumonto1, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!acumonto2, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!difmontoacu, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!porcmontoacu, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!acucant1, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!acucant2, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!difcantacu, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!porccantacu, "###0.00")
                fExport.writeline Aux_Linea
                LineaActual = LineaActual + 1
            Case "1":
                Aux_Linea = Vacio & SeparadorCampos & rs!Conccod & SeparadorCampos & rs!concabr
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!concmonto1, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!concmonto2, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!difmontoconc, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!porcmontoconc, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!conccant1, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!conccant2, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!difcantconc, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!porccantconc, "###0.00")
                fExport.writeline Aux_Linea
                LineaActual = LineaActual + 1
            Case "2":
                Apnom = rs!terape
                If Not EsNulo(rs!terape2) Then
                    Apnom = Apnom & " " & rs!terape2
                End If
                Apnom = Apnom & " " & rs!ternom
                If Not EsNulo(rs!ternom2) Then
                    Apnom = Apnom & " " & rs!ternom2
                End If
                Flog.writeline Espacios(Tabulador * 4) & "Legajo: " & Apnom
                Aux_Linea = Vacio & SeparadorCampos & rs!empleg & SeparadorCampos & Apnom
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!empmonto1, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!empmonto2, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!difmontoemp, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!porcmontoemp, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!empcant1, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!empcant2, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!difcantemp, "###0.00")
                Aux_Linea = Aux_Linea & SeparadorCampos & Format(rs!porccantemp, "###0.00")
                fExport.writeline Aux_Linea
                LineaActual = LineaActual + 1
        End Select
        
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    'Siguiente proceso
    rs.MoveNext
    
        
    'Si hubo algun cambio ==> muestro los totales
    cambiaConc = False
    cambiaacum = False

    If Not rs.EOF Then
        If Acunro_Ant <> rs!acunro Then
            cambiaacum = True
            If tipo <> 0 Then
                cambiaConc = True
            End If
        Else
            If Concnro_Ant <> rs!concnro Then
                cambiaConc = True
                cambiaacum = False
            End If
        End If
    Else
        cambiaacum = True
        If tipo <> 0 Then
            cambiaConc = True
        End If
    End If
    Cambios = (cambiaConc Or cambiaacum)
    
    If Cambios And tipo <> 0 Then
        'Muestro los totales
        If cambiaConc And tipo = 2 Then
            Aux_Linea = "--- Resultado " & Conccodant & " " & Concabrant & SeparadorCampos & Vacio & SeparadorCampos & Vacio
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Concmonto1ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Concmonto2ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Difmontoconcant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Porcmontoconcant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Conccant1ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Conccant2ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Difcantconcant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Porccantconcant, "###0.00")
            fExport.writeline Aux_Linea
            LineaActual = LineaActual + 1
            
            'aca deberia ir un salto de pagina
            If LineaActual Mod LineasPorHoja <> 0 Then
                If Not cambiaacum Then
                    Do
                        fExport.writeline
                        LineaActual = LineaActual + 1
                    Loop While LineaActual Mod LineasPorHoja <> 0
                End If
            End If
            'fExport.writeline Chr$(13) & Chr$(10)
            'fExport.writeline vbCrLf
            'Chr$(13) & Chr$(10)
            
        End If
        
        If cambiaacum Then
            Aux_Linea = "Resultado " & Acudesabrant
            For i = 1 To Nro_Col - 9
                Aux_Linea = Aux_Linea & SeparadorCampos & Vacio
            Next i
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Acumonto1ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Acumonto2ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Difmontoacuant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Porcmontoacuant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Acucant1ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Acucant2ant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Difcantacuant, "###0.00")
            Aux_Linea = Aux_Linea & SeparadorCampos & Format(Porccantacuant, "###0.00")
            fExport.writeline Aux_Linea
            LineaActual = LineaActual + 1
            
            'aca deberia ir un salto de pagina
            If LineaActual Mod LineasPorHoja <> 0 Then
                Do
                    fExport.writeline
                    LineaActual = LineaActual + 1
                Loop While LineaActual Mod LineasPorHoja <> 0
            End If
            
        End If
    End If
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
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim Proceso As Long
Dim Detalle As Integer

'Orden de los parametros
'bpronro
'tipo de detalle    '0 - Acumulador
                    '1 - Concepto
                    '2 - Empleado

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Proceso = CLng(Trim(Mid(parametros, pos1, pos2 - pos1 + 1)))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Detalle = CInt(Trim(Mid(parametros, pos1, pos2 - pos1 + 1)))
    End If
End If
Call Generacion(Proceso, Detalle)
'Select Case Detalle
'Case 0:
'    'Call Generacion_Por_Acumulador(proceso)
'    Flog.writeline Espacios(Tabulador * 1) & "Tipo de detalle No implementado " & Detalle
'Case 1:
'    'Call Generacion_Por_Concepto(proceso)
'    Flog.writeline Espacios(Tabulador * 1) & "Tipo de detalle No implementado " & Detalle
'Case 2:
'    Call Generacion_por_Empleado(Proceso)
'Case Else
'    Flog.writeline Espacios(Tabulador * 1) & "Tipo de detalle desconocido"
'End Select
End Sub





