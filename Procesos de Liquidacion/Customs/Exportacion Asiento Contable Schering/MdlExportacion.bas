Attribute VB_Name = "MdlExportacion"
Option Explicit

'Version: 2.01
'Integracion con las politicas customizadas de Halliburton
Const Version = 1.01
Const FechaVersion = "27/09/2005"


Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser               As String
Global Fecha                As Date
Global Hora                 As String

'Adri�n - Declaraci�n de dos nuevos registros.
Global rs_Empresa           As New ADODB.Recordset
Global rs_tipocod           As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo    As Date
Global StrSql2              As String
Global SeparadorDecimales   As String
Global totalImporte         As Double
Global Total                As Double
Global TotalABS             As Double
Global UltimaLeyenda        As String
Global EsUltimoItem         As Boolean
Global EsUltimoProceso      As Boolean



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

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Exp_Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 27 AND bpronro =" & NroProcesoBatch
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
    
    Flog.writeline Espacios(Tabulador * 1) & "FIN " & Now
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
    Exit Sub
    
ME_Main:
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    HuboError = True

    If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
    Set rs_batch_proceso = Nothing
End Sub

Public Sub Generacion(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Asinro As String, ByVal Empresa As Long, ByVal TipoArchivo As Long, ByVal ProcVol As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Asiento Contable
' Autor      : FGZ
' Fecha      : 25/10/2004
' Modificado : 09/08/2005 - Fapitalle N. - Se agrega la generacion de encabezado
'                                        - Se agregan casos en Programa: TAB y CUENTAZ
'                                        - Se agregan casos especiales en Programa para Shering
'                                        - Se agregan casos especiales en Programa para Halliburton
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim strLinea As String
Dim Aux_Linea As String
Dim cadena As String
Dim Aux_str As String
Dim tipo As String
Dim Cantidad As String
Dim Posicion As String
Dim Formato  As String
Dim Nro As Long
Dim NroL As Long
Dim Programa As String
Dim POS As Integer
Dim debeCod As String
Dim haberCod As String
Dim tmpStr As String
Dim separadorCampos
Dim Completa As Boolean
Dim asi_cod_ant
Dim Enter As String
Dim Fecha_Proc As Date
Dim esprimero As Boolean

Dim Cadena_Anterior As String
Dim Cadena_Anterior2 As String
Dim Desc_Item As String
Dim Producto As String
Dim Cuentas_Diferentes As String
Dim Cuenta_Contable As String
Dim Codigo1 As String
Dim Texto2 As String
Dim Ant_Monto As String
Dim CCosto As String
Dim Fec_Asig As String
Dim Texto As String

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Items As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 234"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo ser� generado en el directorio default"
    End If
    SeparadorDecimales = rs_Modelo!modsepdec
    separadorCampos = rs_Modelo!modseparador
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontr� el modelo. El archivo ser� generado en el directorio default"
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontr� el Periodo"
    Exit Sub
End If

'Seteo el nombre del archivo generado
Select Case TipoArchivo
    Case 1
        Archivo = Directorio & "\asi_" & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
    Case 2
        tmpStr = "int_cont_AR"
        tmpStr = tmpStr & "_" & Format(CStr(Year(Date)), "0000")
        tmpStr = tmpStr & "_" & Format(CStr(Month(Date)), "00")
        tmpStr = tmpStr & "_" & Format(CStr(Day(Date)), "00")
        tmpStr = tmpStr & "_" & Format(CStr(Hour(Now)), "00")
        tmpStr = tmpStr & "" & Format(CStr(Minute(Now)), "00")
        tmpStr = tmpStr & "" & Format(CStr(Second(Now)), "00")
        tmpStr = tmpStr & "_01.txt"
        Archivo = Directorio & "\" & tmpStr
    Case 3
        Archivo = Directorio & "\SAP" & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
    Case 4
        Archivo = Directorio & "\" & Format(Right(CStr(Asinro), 4), "0000") & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
    Case Else
        Archivo = Directorio & "\asi_" & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
   
End Select

Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se crear�."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

On Error GoTo CE

' Comienzo la transaccion
MyBeginTrans

'Busco los procesos a evaluar
StrSql = "SELECT * FROM  proc_vol "
StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
StrSql = StrSql & " WHERE proc_vol.pliqnro =" & Nroliq
If ProcVol <> 0 Then 'si no son todos
    StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
End If
StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
StrSql = StrSql & " ORDER BY linea_asi.masinro, linea_asi.linea,linea_asi.cuenta, linea_asi.dh"
OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay Proceso de Volcados para ese asiento en ese periodo"
Else
    Flog.writeline Espacios(Tabulador * 1) & " Lineas de Procesos de Volcados para ese asiento en ese periodo " & CConceptosAProc
End If
IncPorc = (100 / CConceptosAProc)

'Procesamiento
If rs_Procesos.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
End If

'------------------------------------------------------------------------
' Genero el encabezado de la exportacion
'------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del encabezado del proceso de volcado "
Flog.writeline

'Cantidad_Warnings = 0
'Nro = Nro + 1 'Contador de Lineas

StrSql = "SELECT * FROM confitemicenc "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicenc.itemicnro = itemintcont.itemicnro "
StrSql = StrSql & " ORDER BY confitemicenc.confitemicorden "
OpenRecordset StrSql, rs_Items
            
Enter = Chr(13) + Chr(10)
Fecha_Proc = Date
Aux_Linea = ""
Do While Not rs_Items.EOF
    cadena = ""
    If rs_Items!itemicfijo Then
        If rs_Items!itemicvalorfijo = "" Then
            cadena = String(256, " ")
        Else
            cadena = rs_Items!itemicvalorfijo
        End If
    Else
        Programa = UCase(rs_Items!itemicprog)
        Select Case Programa
        Case "HEADHALLISAP":
            Call Archivo_ASTO_SAP(Directorio, rs_Periodo!pliqhasta)
            cadena = ";Company Code;6055;;Control Totals" + Enter
            cadena = cadena + ";Posting Date;" + Format(Fecha_Proc, "MM/DD/YY") + Enter
            cadena = cadena + ";Document Date;" + Format(Fecha_Proc, "MM/DD/YY") + Enter
            cadena = cadena + ";Reversal Entry Date" + Enter
            cadena = cadena + ";Document Type;SA" + Enter
            cadena = cadena + ";Currency;ARS" + Enter
            cadena = cadena + ";Reference Document;Sueldos" + Enter
            cadena = cadena + ";Document Header;Sueldos" + Enter
            cadena = cadena + ";Calculate Tax (Put X)" + Enter
        Case "LINEHALLISAP":
            cadena = "Line # ; SAP G/L Account ; Amount ; Tax Code ; Cost Center ; Internal Order ; Profit Center ; Personnel Number ; Intercompany ; Allocation ; Line Item Text ; Quantity ; UoM ; WBS Element ; Network ; Activity ; TP Profit Center ; Trading Partner ; Settlement Period ; Tax Jur code ; Asset Trans Type ; Tax Tran Type"
        Case "FECHA MYYYY":
            Call Fecha4(rs_Procesos!vol_fec_asiento, cadena)
        Case "ESPACIOS":
            cadena = String(rs_Items!itemiclong, " ")
        Case "FECHAACTUAL" To "FECHAACTUAL YYYYYYYY"
             If Len(Programa) >= 13 Then
                 Formato = Mid(Programa, 13, Len(Programa) - 6)
             Else
                 Formato = "DDMMYYYY"
             End If
             Select Case Formato
                 Case "YYYDDD":
                      Call Fecha1(Date, cadena)
                 Case Else
                      Call Fecha_Estandar(Date, Formato, True, rs_Items!itemiclong, cadena)
             End Select

        Case Else
            cadena = " ERROR "
            Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
        End Select
    End If
        
    If Mid(cadena, 1, 2) <> "RR" Then
        Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
    Else
        Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
    End If
    
    rs_Items.MoveNext
Loop
   
'Escribo en el archivo de texto
If Trim(Aux_Linea) <> "" Then
   fExport.writeline Aux_Linea '& Aux_Relleno
End If


'------------------------------------------------------------------------
' Genero el detalle de la exportacion
'------------------------------------------------------------------------
totalImporte = 0
Total = 0
NroL = 1
UltimaLeyenda = ""
EsUltimoItem = False
EsUltimoProceso = False

asi_cod_ant = -1
Do While Not rs_Procesos.EOF
        If EsUltimoRegistro(rs_Procesos) Then
            EsUltimoProceso = True
        End If
        Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
        Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del proceso de volcado " & rs_Procesos!vol_cod & " Linea " & rs_Procesos!masinro & " cuenta: " & rs_Procesos!Cuenta
        Flog.writeline
        
        Cantidad_Warnings = 0
        Nro = Nro + 1 'Contador de Lineas
        
        If UCase(UltimaLeyenda) <> UCase(rs_Procesos!desclinea) Then
            NroL = NroL + 1
        End If
        UltimaLeyenda = rs_Procesos!desclinea
        
        StrSql = "SELECT * FROM confitemic "
        StrSql = StrSql & " INNER JOIN itemintcont ON confitemic.itemicnro = itemintcont.itemicnro "
        StrSql = StrSql & " ORDER BY confitemic.confitemicorden "
        OpenRecordset StrSql, rs_Items
                    
        Aux_Linea = ""
        EsUltimoItem = False
        
        
        Do While Not rs_Items.EOF
            'caso particular de Schering
            Select Case Len(rs_Procesos!Cuenta)
                Case 19:
                    Cuenta_Contable = "00" + Mid(rs_Procesos!Cuenta, 11, 9) + "      "
                    Flog.writeline Espacios(Tabulador * 3) & "Cuenta = " & Cuenta_Contable
                    Desc_Item = "ITEMA"
                    CCosto = "    "
                    Producto = "        "
                    Cuentas_Diferentes = "       "
                    Fec_Asig = Format(rs_Procesos!vol_fec_asiento, "DDMMYYYY") + Mid(rs_Procesos!Cuenta, 14, 5) + Espacios(13)
                Case 10:
                    Cuenta_Contable = rs_Procesos!Cuenta
                    Flog.writeline Espacios(Tabulador * 3) & "Cuenta = " & Cuenta_Contable
                    Desc_Item = "ITEMS"
                    CCosto = "|   "
                    Producto = "|       "
                    Cuentas_Diferentes = " | | | "
                    Fec_Asig = "|       " + Espacios(18)
                Case 14:
                    Desc_Item = "ITEMS"
                    Cuenta_Contable = Mid(rs_Procesos!Cuenta, 1, 10)
                    Flog.writeline Espacios(Tabulador * 3) & "Cuenta = " & Cuenta_Contable
                    CCosto = Mid(rs_Procesos!Cuenta, 11, 4)
                    Flog.writeline Espacios(Tabulador * 3) & "CCosto = " & CCosto
                    Producto = "|       "
                    Cuentas_Diferentes = " | | | "
                    Fec_Asig = "|       " + Espacios(18)
                Case Else: 'no deberia darse
                    Flog.writeline Espacios(Tabulador * 3) & "no deberia darse"
                    Desc_Item = "ITEMX"
                    Cuenta_Contable = Mid(rs_Procesos!Cuenta, 1, 10)
                    Flog.writeline Espacios(Tabulador * 3) & "Cuenta Contable " & Cuenta_Contable
                    Cuentas_Diferentes = "       "
            End Select
            
            Select Case rs_Procesos!masinro
                Case 1:
                    Texto = "HABERES Y RETENCIONES    "
                    Texto2 = Texto
                    If Desc_Item = "ITEMA" Then
                        Codigo1 = "39"
                    Else
                        If rs_Procesos!Dh Then
                            Codigo1 = "40"
                        Else
                            Codigo1 = "50"
                        End If
                    End If
                Case 2:
                    Texto = "APORTES PATRONALES       "
                    Texto2 = Texto
                    If rs_Procesos!Dh Then
                        Codigo1 = "40"
                    Else
                        Codigo1 = "50"
                    End If
                Case 3:
                    Texto = "PREVISIONES              "
                    Texto2 = Texto
                    'codigo1 = IF LAST-OF (linea-asi.asi_cod) THEN "  " ELSE "40".
                    If rs_Procesos!Dh Then
                        Codigo1 = "40"
                    Else
                        Codigo1 = "50"
                    End If

                Case 4:
                    Texto = "INTERES S/PRESTAMO" + Format(rs_Procesos!vol_fec_asiento, "MM/YYYY")
                    Texto2 = "INT.S/PREST.-ANTIC.AL PERSONAL"
                    If Desc_Item = "ITEMA" Then
                        Codigo1 = "29"
                    Else
                        If rs_Procesos!Dh Then
                            Codigo1 = "40"
                        Else
                            Codigo1 = "50"
                        End If
                    End If
                    
                    rs_Procesos.MoveNext
                    If rs_Procesos.EOF Then
                       CCosto = "9999"
                       Cuentas_Diferentes = "      "
                    End If
                    rs_Procesos.MovePrevious
                    
                Case Else: 'no deberia darse
                    Texto = "<< masinro > 4 >>"
            End Select
            
            Ant_Monto = "0"
            Call ImporteABS_3(rs_Procesos!Monto, rs_Procesos!Dh, True, 15, Ant_Monto)
            
            'Flog.writeline Espacios(Tabulador * 2) & "Item: " & rs_Items!itemicdesabr
            cadena = ""
            If rs_Items!itemicfijo Then
                If rs_Items!itemicvalorfijo = "" Then
                    cadena = String(256, " ")
                Else
                    cadena = rs_Items!itemicvalorfijo
                    If Len(cadena) < rs_Items!itemiclong Then
                        cadena = cadena & String(rs_Items!itemiclong - Len(cadena), " ")
                    End If
                End If
            Else
                Programa = UCase(rs_Items!itemicprog)
                Flog.writeline Espacios(Tabulador * 3) & "Programa: " & Programa
                Select Case Programa
                '/////casos especiales de shering - start//////////
                Case "HEAD":
                    If rs_Procesos!masinro <> asi_cod_ant Then
                        Call Hacer_Header(rs_Procesos!Dh, rs_Procesos!Cuenta, rs_Procesos!masinro, rs_Procesos!vol_fec_asiento, cadena)
                        asi_cod_ant = rs_Procesos!masinro
                        fExport.Write cadena & Texto
                        Flog.writeline Espacios(Tabulador * 0) & cadena
                        cadena = ""
                        esprimero = True
                    End If
                Case "DETALLESHERING":
                    If esprimero Then
                       fExport.writeline IIf(Not EsNulo(Codigo1), Codigo1, "") & IIf(Not EsNulo(Cuenta_Contable), Cuenta_Contable, "")
                    Else
                       fExport.writeline IIf(Not EsNulo(Codigo1), Codigo1, "") & IIf(Not EsNulo(Cuenta_Contable), Cuenta_Contable, "") & IIf(Not EsNulo(Cuentas_Diferentes), Cuentas_Diferentes, "")
                    End If
                    esprimero = False
                    fExport.Write IIf(Not EsNulo(Cadena_Anterior), Cadena_Anterior, "")
                    'Flog.writeline Espacios(Tabulador * 0) & cadena
                Case "PIESHERING":
                    If Hacer_Pie(rs_Procesos) Then
                        'cadena = Mid(Aux_Linea, 1, 110) + String(13, " ") + Mid(Aux_Linea, 124, 6)
                        'If rs_Procesos!masinro = 4 Then
                        '    cadena = Mid(cadena, 1, 49) + "INT.S/PREST.-ANTIC.AL PERSONAL                   " + Mid(cadena, 99, 30)
                        'End If
                        'fExport.writeline cadena
                        cadena = ""
                        fExport.writeline Espacios(12) & Cuentas_Diferentes
                        fExport.writeline "FINAL"
                    End If
               
                Case Else
                    cadena = " ERROR "
                    Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
                End Select
            End If
                
            If Mid(cadena, 1, 2) <> "RR" Then
                Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
            Else
                Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
            End If
            
            'Schering
            If Len(Texto) < 50 Then
                Texto = Texto & String(50 - Len(Texto), " ")
            End If
            If Len(Texto2) < 50 Then
                Texto2 = Texto2 & String(50 - Len(Texto2), " ")
            End If
            If Hacer_Pie(rs_Procesos) Then
               Cadena_Anterior = IIf(Not EsNulo(Desc_Item), Desc_Item, "") + IIf(Not EsNulo(Ant_Monto), Ant_Monto, "") + "| " + IIf(Not EsNulo(Fec_Asig), Fec_Asig, "") + IIf(Not EsNulo(Texto2), Texto2, "") + IIf(Not EsNulo(CCosto), CCosto, "") + IIf(Not EsNulo(Producto), Producto, "")
            Else
               Cadena_Anterior = IIf(Not EsNulo(Desc_Item), Desc_Item, "") + IIf(Not EsNulo(Ant_Monto), Ant_Monto, "") + "| " + IIf(Not EsNulo(Fec_Asig), Fec_Asig, "") + IIf(Not EsNulo(Texto), Texto, "") + IIf(Not EsNulo(CCosto), CCosto, "") + IIf(Not EsNulo(Producto), Producto, "")
            End If
            Flog.writeline Espacios(Tabulador * 2) & "guardo la linea anterior " & Cadena_Anterior
            
            rs_Items.MoveNext
            
        Loop
            
        ' ------------------------------------------------------------------------
        'Escribo en el archivo de texto
        'Aux_Relleno = Space(256 - Len(Aux_Linea))
        'fExport.writeline Aux_Linea '& Aux_Relleno
            
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    'Siguiente proceso
    rs_Procesos.MoveNext
    
Loop

'------------------------------------------------------------------------
' Genero el pie de la exportacion
'------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del pie del proceso de volcado "
Flog.writeline

Cantidad_Warnings = 0
Nro = Nro + 1 'Contador de Lineas

StrSql = "SELECT * FROM confitemicpie "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicpie.itemicnro = itemintcont.itemicnro "
StrSql = StrSql & " ORDER BY confitemicpie.confitemicorden "
OpenRecordset StrSql, rs_Items
            
Aux_Linea = ""
Do While Not rs_Items.EOF
    cadena = ""
    If rs_Items!itemicfijo Then
        If rs_Items!itemicvalorfijo = "" Then
            cadena = String(256, " ")
        Else
            cadena = rs_Items!itemicvalorfijo
        End If
    Else
        Programa = UCase(rs_Items!itemicprog)
        Select Case Programa
        Case "PIESAP":
            cadena = "*****;;0.00"
        
        Case "ESPACIOS":
            cadena = String(rs_Items!itemiclong, " ")
        
        Case "PERIODO" To "PERIODO 99":
            If Len(Programa) > 8 Then
                Posicion = Mid(Programa, 9, 2)
                Call NroPeriodo(rs_Periodo!pliqmes, CLng(Posicion), True, rs_Items!itemiclong, cadena)
            Else
                Posicion = "0"
                Call NroPeriodo(rs_Periodo!pliqmes, CLng(Posicion), True, rs_Items!itemiclong, cadena)
            End If
        
        Case "FECHAACTUAL" To "FECHAACTUAL YYYYYYYY"
            If Len(Programa) >= 13 Then
                Formato = Mid(Programa, 13, Len(Programa) - 6)
            Else
                Formato = "DDMMYYYY"
            End If
        
            Select Case Formato
            Case "YYYDDD":
                Call Fecha1(Date, cadena)
            Case Else
                Call Fecha_Estandar(Date, Formato, True, rs_Items!itemiclong, cadena)
            End Select

        Case "IMPORTETOTAL":
            Call ImporteTotal(True, rs_Items!itemiclong, cadena)

        Case "TOTALREG":
            Call totalRegistros(Nro - 1, True, rs_Items!itemiclong, cadena)

        Case Else
            cadena = " ERROR "
            Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
        End Select
    End If
        
    If Mid(cadena, 1, 2) <> "RR" Then
        Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
    Else
        Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
    End If
    
    rs_Items.MoveNext
Loop
   

' ------------------------------------------------------------------------
'Escribo en el archivo de texto
'Aux_Relleno = Space(256 - Len(Aux_Linea))
If Trim(Aux_Linea) <> "" Then
   fExport.writeline Aux_Linea '& Aux_Relleno
End If

'Cierro el archivo creado
fExport.Close

'Fin de la transaccion
MyCommitTrans


If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_Items.State = adStateOpen Then rs_Items.Close

Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Modelo = Nothing
Set rs_Items = Nothing

Exit Sub
CE:
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    HuboError = True
    MyRollbackTrans

    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    If rs_Items.State = adStateOpen Then rs_Items.Close
    
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_Modelo = Nothing
    Set rs_Items = Nothing
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

Dim Periodo As Long
Dim Asiento As String
Dim Empresa As Long
Dim TipoArchivo As Long
Dim ProcVol As Long

'Orden de los parametros
'pliqnro
'Asinro, lista separada por comas
'tipoarchivo
'proceso de volcado, 0=todos

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Periodo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Asiento = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        TipoArchivo = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        ProcVol = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If
Call Generacion(bpronro, Periodo, Asiento, Empresa, TipoArchivo, ProcVol)
End Sub

Private Sub NroCuenta(ByVal Cuenta As String, ByVal POS As Integer, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               desde la posicion Pos por una cantidad Cant
'               OBS: Completa con ESPACIOS al FINAL,
'               ej: si la cuenta es: 11000003.521110.01
'                   pos = 1, Cant = 8, Completar = True , Longitud = 12
'                   debera salir: '11000003    '
'
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < Cant Then
        cadena = Mid(Cuenta, POS, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, POS, Cant)
    End If

    If Completar Then
        If Len(cadena) < Longitud Then
            'cadena = String(Longitud - Len(cadena), "0") & cadena
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena
End Sub



Private Sub NroCuenta_1(ByVal Cuenta As String, ByVal POS As Integer, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               desde la posicion Pos por una cantidad Cant y completa con ceros hasta Longitud
'               ej: si la cuenta es: 11000003.521110.01
'                   pos = 1, Cant = 8, Completar = True , Longitud = 12
'                   debera salir: 000011000003
'  Autor: Fapitalle N.
'  Fecha: 09/08/2005
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < Cant Then
        cadena = Mid(Cuenta, POS, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, POS, Cant)
    End If

    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
            'cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena
End Sub

Private Sub Hacer_Header(ByVal Dh As Boolean, ByVal Cuenta As String, ByVal Asi_Cod As String, ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci�n: Devuelve el encabezado por asi_cod para la exportacion de shering
'  Autor: Fapitalle N.
'  Fecha: 12/08/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim Cuenta_Contable As String
Dim Codigo1 As String
Dim Texto As String

    cadena = "HEADR" + Format(Fecha, "DDMMYYYY") + "SA" + Format(Fecha, "MM")
    
    If Dh Then
        Codigo1 = "40"
    Else
        Codigo1 = "50"
    End If
    
    Select Case Asi_Cod
        Case 1:
            Texto = "HABERES Y RETENCIONES    "
            If Len(Cuenta) = 19 Then
                Codigo1 = "39"
            End If
        Case 2:
            Texto = "APORTES PATRONALES       "
        Case 3:
            Texto = "PREVISIONES              "
        Case 4:
            Texto = "INTERES S/PRESTAMO" + Format(Fecha, "MMYYYY")
            If Len(Cuenta) = 19 Then
                Codigo1 = "29"
            End If
        Case Else:  'no deberia darse
            Texto = "<<<..ASICOD.WRONG.....>>>"
    End Select
    
    Select Case Len(Cuenta)
        Case 19:
            Cuenta_Contable = "00" + Mid(Cuenta, 11, 9)
        Case 10:
            Cuenta_Contable = Cuenta
        Case 14:
            Cuenta_Contable = Mid(Cuenta, 1, 10)
        Case Else:  'no deberia darse
            Cuenta_Contable = "<<LENWRG>>"
    End Select
    
    'cadena = cadena + Texto + Codigo1 + Cuenta_Contable
    Cuenta_Contable = ""
    Str_Salida = cadena
End Sub

Private Function Hacer_Pie(ByRef Reg As ADODB.Recordset)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci�n: Devuelve verdadero si se necesita el pie por asi_cod para la exportacion de shering
'  Autor: Fapitalle N.
'  Fecha: 16/08/2005
'-------------------------------------------------------------------------------
Dim Asi_Cod_Actual
Dim hacer As Boolean

    Asi_Cod_Actual = Reg!masinro
    Reg.MoveNext
    If Reg.EOF Then
        hacer = True
    Else
        If Asi_Cod_Actual <> Reg!masinro Then
            hacer = True
        Else
            hacer = False
        End If
    End If
    Reg.MovePrevious
    Hacer_Pie = hacer
End Function

Public Sub Archivo_ASTO_SAP(ByVal Dir As String, ByVal Fecha As Date)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci�n: Genera el archivo ASTOmmaa.txt para el volcado SAP de Halliburton
'  Autor: Fapitalle N.
'  Fecha: 18/08/2005
'-------------------------------------------------------------------------------
Dim fAstoSAP
Dim fs
Dim Archivo
Dim carpeta
Dim cadena As String

Archivo = Dir & "\ASTO" & Format(Fecha, "MMYY") & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fAstoSAP = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Set carpeta = fs.CreateFolder(Dir)
    Set fAstoSAP = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

cadena = "Constante" + ";" + "Blancos" + ";" + "Cuenta" + ";" + "Blancos" + ";" + "Entidad" + ";" + _
         "Blancos" + ";" + "Constante" + ";" + "Blanco" + ";" + "Moneda" + ";" + "Blanco" + ";" + _
         "Constante" + ";" + "Descripcion" + ";" + "Blanco" + ";" + "Fecha" + ";" + "Blanco" + ";" + _
         "Constante" + ";" + "Blanco" + ";" + "Importe" + ";" + "Debe/Haber"

fAstoSAP.writeline cadena
fAstoSAP.Close
    
End Sub

Private Sub Cuenta(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci�n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = cadena & IIf(Len(cadena) = 10, ".1000", "")
    Str_Salida = cadena

End Sub


Private Sub Cuenta_1(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci�n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  OBS  : TRAER LA UNIDAD DE NEGOCIO AL FRENTE DE LA CUENTA
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim subcad As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = cadena & IIf(Len(cadena) = 10, ".1000", "")
    
    'PARA TRAER LA UNIDAD DE NEGOCIO AL FRENTE DE LA CUENTA
    subcad = cadena
    cadena = ""
    cadena = Mid(subcad, 12, 4) & "." & Mid(subcad, 1, 10)
    
    Str_Salida = cadena
End Sub


Private Sub Cuenta_2(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta2.p
'  Descripci�n: devuelve los primeros numeros de la cuenta, hasta el primer punto
'               en un formato de 12 digitos.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    
    cadena = IIf(Mid(cadena, 12, 4) = "", "1000", Mid(cadena, 12, 4))
    Str_Salida = cadena

End Sub

Private Sub Cuenta_3(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta2.p
'  Descripci�n: devuelve los primeros numeros de la cuenta, hasta el primer punto
'               en un formato de 6 digitos.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = IIf(Mid(cadena, 1, 6) = "", "000000", Mid(cadena, 1, 6))
    Str_Salida = cadena

End Sub

Private Sub Cuenta_4(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci�n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = IIf(Mid(cadena, 8, 3) = "", "000", Mid(cadena, 8, 3))
    Str_Salida = cadena

End Sub


Private Sub Importe(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve el importe de la linea en el siguiente formato:
'               Con dos decimales seguidos con el separador definido en el modelo
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If Debe Then
        If Completar Then
            cadena = Format(Numero(0), String(Longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(Longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & SeparadorDecimales & "00"
    End If
    
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(cadena)), 2))
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(Total + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * Total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If Debe Then
                Total = Round(Total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                Total = Round(Total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            Balancea = True
        End If
    Else
        If Debe Then
            Total = Round(Total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            Total = Round(Total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
Loop

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < Longitud Then
        cadena = String(Longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > Longitud Then
            cadena = Right(cadena, Longitud)
        End If
    End If
End If
Str_Salida = cadena
 
End Sub

Private Sub Fecha1(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a�o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del a�o (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Year(CDate(Fecha)) Mod 1900, "000") & Format(CDate(Fecha) - CDate("01/01/" & Year(CDate(Fecha))), "000")
        
    Str_Salida = cadena

End Sub

Private Sub Fecha2(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a�o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del a�o (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, "ddmmyy")
        
    Str_Salida = cadena

End Sub

Private Sub Fecha3(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a�o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del a�o (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, "ddmm")
        
    Str_Salida = cadena

End Sub

Private Sub Fecha4(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve la fecha de emision en el siguiente formato:
'               MYYYY donde los primeros 2 corresponden al mes de 1 a 12
'               y los otros 4 digitos son para el a�o
'  Autor: Fapitalle N.
'  Fecha: 09/08/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, "MYYYY")
        
    Str_Salida = cadena

End Sub


Private Sub Fecha_Estandar(ByVal Fecha As Date, ByVal Formato As String, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a�o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del a�o (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, Formato)
    'Cadena = Format(Fecha, "ddmmyy")
        
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
        
    Str_Salida = cadena

End Sub

Private Sub NroLinea(ByVal Linea As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/nrolinea.p
'  Descripci�n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Linea, String(Longitud, "0"))
        
'    If Completar Then
'        If Len(Cadena) < Longitud Then
'            Cadena = String(Longitud - Len(Cadena), "0") & Cadena
'        End If
'    End If
    Str_Salida = cadena
End Sub

Private Sub NroAsiento(ByVal Asiento As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Asiento, String(Longitud, "0"))
        
    Str_Salida = cadena
End Sub


Private Sub Leyenda(ByVal Descripcion As String, ByVal POS As Integer, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci�n: devuelve el la descripcion.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Descripcion) < Cant Then
        cadena = Mid(Descripcion, POS, Len(Descripcion))
    Else
        cadena = Mid(Descripcion, POS, Cant)
    End If

    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena

End Sub



Private Sub Leyenda1(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal POS As Long, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci�n: devuelve el la descripcion.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Linea As New ADODB.Recordset
Dim Encontro As Boolean

    If Len(Descripcion) < Cant Then
        cadena = Mid(Descripcion, POS, Len(Descripcion))
    Else
        cadena = Mid(Descripcion, POS, Cant)
    End If

Encontro = False
StrSql = "SELECT * FROM mod_linea "
StrSql = StrSql & " WHERE mod_linea.masinro =" & Asi_Cod
StrSql = StrSql & " AND mod_linea.linaorden =" & Linea
OpenRecordset StrSql, rs_Mod_Linea
If Not rs_Mod_Linea.EOF Then
    '1er nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro1) Then
        If rs_Mod_Linea!lineanivternro1 = 32 Then
            Encontro = True
        End If
    End If
    '2do nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro2) Then
        If rs_Mod_Linea!lineanivternro2 = 32 Then
            Encontro = True
        End If
    End If
    '3er nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro3) Then
        If rs_Mod_Linea!lineanivternro3 = 32 Then
            Encontro = True
        End If
    End If
    
    If Encontro Then
        cadena = "Jornales " & cadena
    Else
        cadena = "Sueldos " & cadena
    End If
End If
    
If Completar Then
    If Len(cadena) < Longitud Then
        cadena = cadena & String(Longitud - Len(cadena), " ")
    End If
End If
Str_Salida = cadena

End Sub

Private Sub Leyenda2(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal POS As Long, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci�n: devuelve la descripcion del modelo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset

    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & Asi_Cod
    
    OpenRecordset StrSql, rs_Mod_Asiento
    
    cadena = ""
    
    If Not rs_Mod_Asiento.EOF Then
       cadena = rs_Mod_Asiento!masidesc
    End If
    
    rs_Mod_Asiento.Close
        
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    
    Str_Salida = cadena

End Sub


Private Sub Leyenda3(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal POS As Long, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByVal periodoMes As Integer, ByVal periodoAnio As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci�n: devuelve la descripcion del modelo y el periodo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset

    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & Asi_Cod
    
    OpenRecordset StrSql, rs_Mod_Asiento
    
    cadena = ""
    
    If Not rs_Mod_Asiento.EOF Then
       cadena = Left(rs_Mod_Asiento!masidesc, 7)
    End If
    
    rs_Mod_Asiento.Close
    
    If periodoMes < 10 Then
       cadena = cadena & " 0" & periodoMes
    Else
       cadena = cadena & " " & periodoMes
    End If
    
    cadena = cadena & " " & periodoAnio
        
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    
    Str_Salida = cadena

End Sub


Private Sub NroPeriodo(ByVal Periodo As Long, ByVal Inicial As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Periodo + Inicial, String(Longitud, "0"))
        
    Str_Salida = cadena
End Sub


Private Sub ImporteABS(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = Format(Numero(0), String(Longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    Total = Round(Total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            Total = Round(Total - CDbl(cadena), 2)
            If Diferencia < 0 Then
                'Monto = CSng(Aux_Cadena) - Diferencia
                Monto = TotalABS * -1
            Else
                'Monto = CSng(Aux_Cadena) + Diferencia
                Monto = -1 * TotalABS
            End If
        Else
            If Debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If Debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub


Private Sub ImporteABS_2(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el ".", el relleno es con espacios al final
'               La funcion es una ligera modificacion de ImporteABS
'  Autor: Fapitalle N.
'  Fecha: 10/08/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & "."
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
    
    If Completar Then
        cadena = cadena & String(Longitud - Len(cadena), " ")
    End If
    
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    Total = Round(Total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            Total = Round(Total - CDbl(cadena), 2)
            If Diferencia < 0 Then
                'Monto = CSng(Aux_Cadena) - Diferencia
                Monto = TotalABS * -1
            Else
                'Monto = CSng(Aux_Cadena) + Diferencia
                Monto = -1 * TotalABS
            End If
        Else
            If Debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If Debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    Str_Salida = cadena

End Sub

Private Sub ImporteABS_3(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el ",", el relleno es con ceros adelante
'               La funcion es una ligera modificacion de ImporteABS
'  Autor: Fapitalle N.
'  Fecha: 10/08/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

    Balancea = False


    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & ","
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, "-", "")
    
    If Completar Then
        cadena = String(Longitud - Len(cadena), "0") + cadena
    End If
    
    Str_Salida = cadena
End Sub


Private Sub ImporteABS_old(ByVal Monto As Single, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci�n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    'Parte_Entera = Fix(Monto)
    'Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)

'    If Debe Then
'       totalImporte = totalImporte + Monto
'    Else
'       totalImporte = totalImporte - Monto
'    End If
    
    If Completar Then
        cadena = Format(Numero(0), String(Longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    'totalImporte = totalImporte + Abs(CSng(cadena))
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Private Sub DebeHaber(ByVal Debe As Boolean, ByVal debeCod As String, ByVal haberCod As String, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: devuelve debeCod o haberCod dependiendo si es debe o haber.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Debe Then
        cadena = debeCod
    Else
        cadena = haberCod
    End If
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Private Sub ImporteTotal(ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: devuelve el importe total, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Parte_Entera As Long
Dim Parte_Decimal As Integer
Dim Numero

    Numero = Split(CStr(totalImporte), ".")
    Parte_Entera = Fix(totalImporte)
    Parte_Decimal = IIf(Round((totalImporte - Parte_Entera) * 100, 0) > 0, Round((totalImporte - Parte_Entera) * 100, 0), 0)

    Numero(0) = Parte_Entera
    'cadena = Format(Parte_Entera, String(Longitud - 3, "0")) & SeparadorDecimales & Format(Parte_Decimal, "00")
    cadena = Format(Parte_Entera, String(Longitud - 3, "0")) & SeparadorDecimales
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    cadena = Replace(cadena, ",", ".")
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Private Sub totalRegistros(ByVal Total As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci�n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Total, String(Longitud, "0"))
        
    Str_Salida = cadena
End Sub



Private Sub Importe_Format(ByVal Monto As Single, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String, ByVal Signo As String, ByVal Separador As String)
'--------------------------------------------------------------------------------
'Descripci�n: devuelve el importe de la linea en el siguiente formato:
'               Si va al debe es + y - sino, con dos decimales seguidos con el
'               Separador
'Autor: FGZ
'Fecha: 25/04/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)
    
    Numero(0) = Parte_Entera
    If Completar Then
        cadena = Format(Numero(0), String(Longitud - 3, "0"))
    Else
        cadena = Numero(0)
    End If
    If Debe Then
        cadena = " " & cadena
    Else
        cadena = "-" & cadena
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    If Debe Then
        Aux_Cadena = " " & Format(Numero(0), String(Longitud - 3, "0")) & Separador
    Else
        Aux_Cadena = "-" & Format(Numero(0), String(Longitud - 3, "0")) & Separador
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub


Public Function EsUltimoRegistroItem(ByRef Reg As ADODB.Recordset) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve TRUE si es el ultimo registro del recordset del tipo de item
' Autor      : FGZ
' Fecha      : 17/06/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Hay As Boolean
Dim Aux_Pos As Long

    Hay = False
    Aux_Pos = Reg.AbsolutePosition
    If Not Reg!itemicfijo Then
        Reg.MoveNext
        Do While Not Reg.EOF And Not Hay
            If UCase(Reg!itemicprog) = "IMPORTE" Then
                Hay = True
            End If
            Reg.MoveNext
        Loop
        'Reposiciono
        Reg.MoveFirst
        Do While Not Reg.AbsolutePosition = Aux_Pos
            Reg.MoveNext
        Loop
        If Not Hay Then
            EsUltimoRegistroItem = True
        Else
            EsUltimoRegistroItem = False
        End If
    Else
        EsUltimoRegistroItem = False
    End If
End Function

Public Function EsUltimoRegistroItemABS(ByRef Reg As ADODB.Recordset) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve TRUE si es el ultimo registro del recordset del tipo de item
' Autor      : FGZ
' Fecha      : 17/06/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Hay As Boolean
Dim Aux_Pos As Long

    Hay = False
    Aux_Pos = Reg.AbsolutePosition
    If Not Reg!itemicfijo Then
        Reg.MoveNext
        Do While Not Reg.EOF And Not Hay
            If UCase(Reg!itemicprog) = "IMPORTEABS" Then
                Hay = True
            End If
            Reg.MoveNext
        Loop
        'Reposiciono
        Reg.MoveFirst
        Do While Not Reg.AbsolutePosition = Aux_Pos
            Reg.MoveNext
        Loop
        If Not Hay Then
            EsUltimoRegistroItemABS = True
        Else
            EsUltimoRegistroItemABS = False
        End If
    Else
        EsUltimoRegistroItemABS = False
    End If
End Function

