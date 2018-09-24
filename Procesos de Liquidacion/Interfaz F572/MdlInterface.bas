Attribute VB_Name = "MdlInterface"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "17/07/2014"
'Global Const UltimaModificacion = "Version inicial - Dimatz Rafael - Se creo Interface para levantar Formulario 572 - XML "

'Global Const Version = "1.01"
'Global Const FechaModificacion = "06/08/2014"
'Global Const UltimaModificacion = "Version inicial - Dimatz Rafael - Se modifico para que cuando haga el backup del archivo le concatene el Nro de Proceso. Se modifico para que traiga el CUIL y busque con y sin '-'. Se verifica si existen los tag en el XML "

'Global Const Version = "1.02"
'Global Const FechaModificacion = "22/09/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico Deducciones y Retencion de Pagos para que muestre Sub Totales cuando es necesario. Se agrego el Item ganLiqOtrosEmpEnt "
            
'Global Const Version = "1.03"
'Global Const FechaModificacion = "17/12/2014"
'Global Const UltimaModificacion = " " ' FGZ - Dimatz Rafael - Mejoras varias de manejos de errores, logs, controles, transacciones, etc
''               CAS-20374 - H&A - Adecuaciones LIQ - Importacion F 572 Web

'Global Const Version = "1.04"
'Global Const FechaModificacion = "19/12/2014"
'Global Const UltimaModificacion = " " ' FGZ - Mejoras en validación del CUIL
'               CAS-20374 - H&A - Adecuaciones LIQ - Importacion F 572 Web

'Global Const Version = "1.05"
'Global Const FechaModificacion = "12/01/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se modifico el calculo de las deducciones cuando agrupa las mismas cuotas
'               CAS-20374 - H&A - Adecuaciones LIQ - Importacion F 572 Web

'Global Const Version = "1.06"
'Global Const FechaModificacion = "13/01/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se da formato de 2 decimales al Monto en las deducciones
'               CAS-20374 - H&A - Adecuaciones LIQ - Importacion F 572 Web

'Global Const Version = "1.07"
'Global Const FechaModificacion = "13/01/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se saco el FormatNumber del Monto
'               CAS-20374 - H&A - Adecuaciones LIQ - Importacion F 572 Web

'Global Const Version = "1.08"
'Global Const FechaModificacion = "13/01/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se el replace en el Monto
'               CAS-20374 - H&A - Adecuaciones LIQ - Importacion F 572 Web

'Global Const Version = "1.09"
'Global Const FechaModificacion = "02/02/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se agrego que guarde montos negativos. Se agrego que no guarde si los montos son 0. Se agrego Cuit y Denominacion en GanLiqOtrosEmpEnt
'               CAS-29273 - H&A - LIQ - Interfaz 572 Web - Items con valores negativos

'Global Const Version = "1.10"
'Global Const FechaModificacion = "12/03/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se agrego Items 99 y se modifico para que prorratee multiplicando por los meses que sean necesarios el total
'               CAS-29916 - H&A - LIQ - Interfaz 572 Web - Items 99

'Global Const Version = "1.11"
'Global Const FechaModificacion = "17/03/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se modifico para que ganLiqOtrosEmpEnt no inserte en desmen si solo informa CUIT y Denominacion
'               CAS-29916 - H&A - LIQ - Interfaz 572 Web - ganLiqOtrosEmpEnt

'Global Const Version = "1.12"
'Global Const FechaModificacion = "04/05/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se modifico para que ganLiqOtrosEmpEnt inserte en desmen el CUIT de cada pluriempleo definido pro empleado
'               CAS-30680 - H&A - LIQ - Interfaz 572 Web - ganLiqOtrosEmpEnt

'Global Const Version = "1.13"
'Global Const FechaModificacion = "21/05/2015"
'Global Const UltimaModificacion = " " ' Dimatz Rafael - Se modifico que inserte en desmen el campo desmondec como valor numerico
'               CAS-30825 - G.BAPRO - Bug en interfaz 572 web - Insertar en desmen el campo desmondec como valor numerico

'Global Const Version = "1.14"
'Global Const FechaModificacion = "28/07/2015"
'Global Const UltimaModificacion = " " ' Fernandez,Matias - CAS- 32346 - Prudential - Bug Interfaz 292 -  se contempla que la deduccion
                                      'no tenga periodo

'Global Const Version = "1.15"
'Global Const FechaModificacion = "25/09/2015"
'Global Const UltimaModificacion = " " ' Borrelli Facundo - CAS-32946 - PRAXAIR - Bug en Interfaz 292
                                      ' Para el tipo de duduccion 10, se agregaron los campos Fecha desde/hasta, MontoTotal, Marca de Prorrateado
                                      ' y dentro de detalles el valor que corresponde a "desc", ademas se agrego el campo periodo_anio al T_RegPeriodo.

'Global Const Version = "1.16"
'Global Const FechaModificacion = "29/09/2015"
'Global Const UltimaModificacion = " " ' Miriam Ruiz - CAS-33306 - ACARA - LIQ - Interfaz F572 Web - Bug otros empleadores
                                      ' Se inicializan todas las variables
                                  
'Global Const Version = "1.17"
'Global Const FechaModificacion = "01/10/2015"
'Global Const UltimaModificacion = " " ' Borrelli Facundo - CAS-32946 - PRAXAIR - Bug en Interfaz 292 [Entrega 2]
                                      ' Se corrige la funcion Grabar_DeduccionesTipo10() al grabar el periodo
                                      
'Global Const Version = "1.18"
'Global Const FechaModificacion = "09/10/2015"
'Global Const UltimaModificacion = " " ' Borrelli Facundo - CAS-33306 - ACARA - LIQ - Interfaz F572 Web - Bug Gcias otros empleadores no graba razon social
                                      ' Se inserta la Razon Social (Denominacion) en desmen cuando el monto es <> 0, cuando se informan ganancias de otros empleadores
                                      
Global Const Version = "1.19"
Global Const FechaModificacion = "25/11/2015"
Global Const UltimaModificacion = " " ' Dimatz Rafael - CAS 33582 - BAPRO - Mejora en proceso de Interfaz F572 web
                                      ' Se Guarda tantas veces los datos como ternro figure por CUIT
                                      
'------------------------------------------------------------------------------------

Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
Global HuboErrorLocal As Boolean
Global Path
Global NArchivo
Global NroLinea As Long

Global separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean
Global NroModelo As Integer
Global NombreArchivo As String

Global cuit
Global NroTernro
Global Arr_NroTernro(5) As Integer

Global Deducciones() As T_RegDeducciones
Global retPerPagos() As T_RegRetPerPago
Global CargaFamilia() As T_RegCargaFamilia
Global GanLiq() As T_RegGanLiq
Global DescripcionModelo As String


Dim directorio As String
Dim CArchivos
Dim Archivo
Dim Folder
Global cantArch As Integer
Public Sub Escribir_Log(ByVal TipoLog As String, ByVal Lin As Long, ByVal col As Long, ByVal msg As String, ByVal CantTab As Long, ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Escribe un mensage determinado en uno de 3 archivos de log
' Autor      : FGZ
' Fecha      : 18/04/2005
' Ultima Mod.:
' Descripcion: 13/02/2012 - Gonzalez Nicolás - Se agregó Multilenguaje
' ---------------------------------------------------------------------------------------------
Dim Texto
'Texto = EscribeLogMI("Linea") & " " & Lin
'Texto = Texto & " " & EscribeLogMI("Columna") & " " & col
Select Case UCase(TipoLog)

    Case "FLOG" 'Archivo de Informacion de resumen
            Flog.writeline Espacios(Tabulador * CantTab) & msg
    Case "FLOGE" 'Archivo de Errores
            'FlogE.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & col & ": " & msg
            FlogE.writeline Espacios(Tabulador * CantTab) & Texto & ": " & msg
            FlogE.writeline Espacios(Tabulador * CantTab) & strLinea
    Case "FLOGP" 'Archivo de lineas procesadas
            'FlogP.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & col & ": " & msg
            FlogP.writeline Espacios(Tabulador * CantTab) & Texto & ": " & msg
    Case Else
        'Flog.writeline Espacios(Tabulador * CantTab) & "Nombre de archivo de log incorrecto " & TipoLog
        'Flog.writeline Espacios(Tabulador * CantTab) & Replace(EscribeLogMI("Nombre de archivo de log incorrecto"), "@@TXT@@", TipoLog)
End Select

End Sub
Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial de Interface.
    ' Autor      : Lisandro Moro
    ' Fecha      : 12/07/2011
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim objconnMain As New ADODB.Connection
    Dim strCmdLine
    Dim Nombre_Arch As String
    Dim rs_batch_proceso As New ADODB.Recordset
    Dim bprcparam As String
    Dim PID As String
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
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Interface_F572 " & "-" & NroProcesoBatch & ".log"
    'Archivo de log
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"

    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 304 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    'Primera_Vez = True
    
    If Not rs_batch_proceso.EOF Then
        'bprcparam = rs_batch_proceso!bprcparam
        usuario = rs_batch_proceso!iduser
        'rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        'Call LevantarParamteros(bprcparam)
        Call ComenzarTransferencia
    End If
    'FGZ - 17/12/2014 -------------------------------------------------------------------
    NroModelo = 292
    'FGZ - 17/12/2014 -------------------------------------------------------------------
    Call ComenzarTransferencia
    If Not HuboError Then
        If ErroresNov Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        End If
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    End
End Sub

Private Sub LeeArchivo(ByVal NombreArchivo As String)
    'Espero hasta que se crea el archivo
    'On Error Resume Next
    On Error GoTo CE
    Err.Number = 1
    Do Until Err.Number = 0
        Flog.writeline "Marca:" & Err.Number
        Err.Number = 0
        Set f = fs.GetFile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    
    Dim doc As DOMDocument30
    Dim nodes As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim node2 As IXMLDOMNode
    Dim node3 As IXMLDOMNode
    Dim s As String
    Dim I As Integer
    Dim b As Integer
    Dim T As Integer
    Dim periodo As Integer
    Dim NroDoc As String
    Dim Denominacion As String
    Dim desano As Integer
    Dim StrSql As String
    Dim rsConsult As New ADODB.Recordset
    Dim desmondec As String
    Dim desmenprorra As Integer
    Dim itenro As Integer
    Dim desfecdes As Date
    Dim desfechas As Date
    Dim CargaFamilia_Vacio
    Dim retPerPagos_Vacio
    Dim Deducciones_Vacio
    Dim GanLiq_Vacio
    Dim Meses
    Dim razonsocial
    Dim Fam_Cuit
    
    Set doc = New DOMDocument30
    doc.Load NombreArchivo


    'FGZ - 17/12/2014 ---------------------------------------------------------------------------
    StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                  NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
    objConn.Execute StrSql, , adExecuteNoRecords
    crpNro = getLastIdentity(objConn, "inter_pin")
    'FGZ - 17/12/2014 ---------------------------------------------------------------------------

'------------------------------------- Genera Cuit --------------------------------------------------------
    Set nodes = doc.selectNodes("presentacion/periodo")
    periodo = nodes.Item(0).Text

    'FGZ - 17/12/2014 ---------------------------------------------------------------------------
    For I = 0 To 5
        Arr_NroTernro(I) = 0
    Next
    'FGZ - 17/12/2014 ---------------------------------------------------------------------------
    Set nodes = doc.selectNodes("presentacion/empleado")
        For Each node In nodes
            Flog.writeline "Generar Empleados."
            Call Generar_Empleado(node)
        Next
        For T = 0 To UBound(Arr_NroTernro())
            StrSql = "SELECT * FROM desmen WHERE desano = " & periodo & " AND empleado = " & Arr_NroTernro(T)
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
        'FGZ - 17/12/2014 ---------------------------------------------------------------------------
                If Arr_NroTernro(T) <> 0 Then
                    MyBeginTrans
                    StrSql = "DELETE desmen WHERE desano = " & periodo & "AND empleado = " & Arr_NroTernro(T)
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            rsConsult.Close
        Next
    '------------------------------------- Genera Cargas Familiares -------------------------------------------
        Set nodes = doc.selectNodes("presentacion/cargasFamilia")
            CargaFamilia_Vacio = True
            For Each node In nodes
                Flog.writeline "Cargas Familia."
                Call Generar_CargasFamilia(node, doc)
                CargaFamilia_Vacio = False
            Next
    '------------------------------------- Genera Ganancia Liquidacion -------------------------------------------
        Set nodes = doc.selectNodes("presentacion/ganLiqOtrosEmpEnt")
            GanLiq_Vacio = True
            For Each node In nodes
                Flog.writeline "Ganancia Liquidacion."
                Call Generar_GanLiq(node, doc)
                GanLiq_Vacio = False
            Next
    '------------------------------------- Genera Deducciones -------------------------------------------------
        Set nodes = doc.selectNodes("presentacion/deducciones")
            Deducciones_Vacio = True
            For Each node In nodes
                Flog.writeline "Generar Deducciones."
                Call Generar_Deducciones(node, doc)
                Deducciones_Vacio = False
            Next
    
    '------------------------------------- Genera Retenciones -------------------------------------------------
        Set nodes = doc.selectNodes("presentacion/retPerPagos")
            retPerPagos_Vacio = True
            For Each node In nodes
                Flog.writeline "Generar Retenciones."
                Call Generar_Retenciones(node, doc)
                retPerPagos_Vacio = False
            Next
    '----------------------------------------------------------------------------------------------------------
    
    '------------------------------------ Inserta Carga Familia en Desmen -------------------------------------
    If Not CargaFamilia_Vacio Then
        For I = 0 To (UBound(CargaFamilia) - 1)
            desano = periodo
            desmondec = "1"
            desmenprorra = -1
            desfecdes = "01/" & CargaFamilia(I).MesDesde & "/" & periodo
            desfechas = UltimoDiaMes(periodo, CargaFamilia(I).MesHasta)
            razonsocial = CargaFamilia(I).Apellido & " " & CargaFamilia(I).Nombre
           
            Fam_Cuit = CargaFamilia(I).NroDoc
            StrSql = "SELECT itenro FROM parentesco_572 WHERE itenro572 = " & CargaFamilia(I).Parentesco
            OpenRecordset StrSql, rsConsult
        
            If Not rsConsult.EOF Then
                itenro = rsConsult!itenro
            Else
                itenro = 0
            End If
            rsConsult.Close
            
            If (itenro = 8) Or (itenro = 13) Or (itenro = 15) Or (itenro = 21) Or (itenro = 9) Or (itenro = 20) Or (itenro = 22) Or (itenro = 28) Or (itenro = 30) Or (itenro = 24) Or (itenro = 27) Or (itenro = 26) Or (itenro = 25) Then
                desmondec = (-1) * desmondec
            End If
            
            If itenro <> 0 Then
                For T = 0 To UBound(Arr_NroTernro())
                    NroTernro = Arr_NroTernro(T)
                    If NroTernro <> 0 Then
                        StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,desrazsoc,descuit) "
                        StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                        StrSql = StrSql & "'" & desfechas & "','" & razonsocial & "','" & Fam_Cuit & "' )"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline ("Inserto Carga Familia Correctamente")
                    End If
                Next
            Else
                Flog.writeline ("No existe el tipo Parentesco ") & CargaFamilia(I).Parentesco
            End If

        Next I
    End If
    CargaFamilia_Vacio = True
    '----------------------------------------------------------------------------------------------------------
    
    '------------------------------------ Inserta Ganancias en Desmen -------------------------------------
    If Not GanLiq_Vacio Then
        For I = 0 To (UBound(GanLiq) - 1)
            For b = 0 To (UBound(GanLiq(I).ingAp) - 1)
                desano = periodo
                desmenprorra = 0
                If GanLiq(I).ingAp(b).Mes <> 0 Then
                    desfecdes = "01/" & GanLiq(I).ingAp(b).Mes & "/" & periodo
                    desfechas = UltimoDiaMes(periodo, GanLiq(I).ingAp(b).Mes)
                       
                    desmondec = GanLiq(I).ingAp(b).obrasoc
                    cuit = GanLiq(I).cuit
                    'FB - Se guarda Denominacion para cuando desmondec <> "0"
                    Denominacion = GanLiq(I).Denominacion
                    itenro = 6
                    
                    If desmondec <> "0" Then
                        desmondec = (-1) * desmondec
                         For T = 0 To UBound(Arr_NroTernro())
                            NroTernro = Arr_NroTernro(T)
                            If NroTernro <> 0 Then
                                StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                                StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                                StrSql = StrSql & "'" & desfechas & "','" & cuit & "', '" & Denominacion & "')"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline ("Ganancias: ") & StrSql
                            End If
                        Next
                    End If
            '-----------------------------------------------------------------------------------------------------------------------
                    desmondec = GanLiq(I).ingAp(b).segsocial
                    itenro = 5
                    If desmondec <> "0" Then
                        desmondec = (-1) * desmondec
                         For T = 0 To UBound(Arr_NroTernro())
                            NroTernro = Arr_NroTernro(T)
                            If NroTernro <> 0 Then
                                StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                                StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                                StrSql = StrSql & "'" & desfechas & "','" & cuit & "', '" & Denominacion & "')"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline ("Ganancias2: ") & StrSql
                            End If
                        Next
                    End If
            '-----------------------------------------------------------------------------------------------------------------------
                    desmondec = GanLiq(I).ingAp(b).sind
                    itenro = 7
                    If desmondec <> "0" Then
                        desmondec = (-1) * desmondec
                         For T = 0 To UBound(Arr_NroTernro())
                            NroTernro = Arr_NroTernro(T)
                            If NroTernro <> 0 Then
                                StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                                StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                                StrSql = StrSql & "'" & desfechas & "','" & cuit & "', '" & Denominacion & "')"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline ("Ganancias3: ") & StrSql
                            End If
                         Next
                    End If
            '-----------------------------------------------------------------------------------------------------------------------
                    desmondec = GanLiq(I).ingAp(b).ganbrut
                    itenro = 1
                    If desmondec <> "0" Then
                         For T = 0 To UBound(Arr_NroTernro())
                            NroTernro = Arr_NroTernro(T)
                            If NroTernro <> 0 Then
                                StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                                StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                                StrSql = StrSql & "'" & desfechas & "','" & cuit & "', '" & Denominacion & "')"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline ("Ganancias4: ") & StrSql
                            End If
                         Next
                    End If
            '-----------------------------------------------------------------------------------------------------------------------
                    desmondec = GanLiq(I).ingAp(b).ajuste
                    itenro = 1
                    If desmondec <> "0" Then
                     For T = 0 To UBound(Arr_NroTernro())
                        NroTernro = Arr_NroTernro(T)
                        If NroTernro <> 0 Then
                            StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                            StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                            StrSql = StrSql & "'" & desfechas & "','" & cuit & "', '" & Denominacion & "')"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline ("Ganancias5: ") & StrSql
                        End If
                     Next
                    End If
            '-----------------------------------------------------------------------------------------------------------------------
                    desmondec = GanLiq(I).ingAp(b).retgan
                    desfechas = UltimoDiaMes(periodo, GanLiq(I).ingAp(b).Mes)
                    For T = 0 To UBound(Arr_NroTernro)
                        StrSql = "SELECT * FROM ficharet WHERE fecha = '" & desfechas & "' "
                        StrSql = StrSql & " AND empleado = " & Arr_NroTernro(T) & " AND liqsistema = 0 "
                        OpenRecordset StrSql, rsConsult
                        If Not rsConsult.EOF Then
                            If Arr_NroTernro(T) <> 0 Then
                                StrSql = "DELETE ficharet WHERE fecha = '" & desfechas & "' "
                                StrSql = StrSql & " AND empleado = " & Arr_NroTernro(T) & " AND liqsistema = 0 "
                                objConn.Execute StrSql, , adExecuteNoRecords
                            End If
                        End If
                    rsConsult.Close
                    Next
                    If desmondec <> "0" Then
                    For T = 0 To UBound(Arr_NroTernro())
                        NroTernro = Arr_NroTernro(T)
                        If NroTernro <> 0 Then
                            StrSql = "INSERT INTO ficharet(importe,fecha, empleado,pronro,liqsistema) VALUES(" & desmondec & ",'" & desfechas & "'," & Arr_NroTernro(T) & ",0,0) "
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Next
                    End If
            '-----------------------------------------------------------------------------------------------------------------------
                    desmondec = GanLiq(I).ingAp(b).retribNoHab
                    itenro = 1

                    If desmondec <> "0" Then
                         For T = 0 To UBound(Arr_NroTernro())
                            NroTernro = Arr_NroTernro(T)
                            If NroTernro <> 0 Then
                                StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                                StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                                StrSql = StrSql & "'" & desfechas & "','" & cuit & "', '" & Denominacion & "')"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline ("Ganancias7: ") & StrSql
                            End If
                         Next
                    End If
             End If
                      GanLiq(I).ingAp(b).Mes = 0
                      GanLiq(I).ingAp(b).obrasoc = ""
                      'FB - Se comenta, para no blanquear la variable denominacion, y guardarla en demen
                      'GanLiq(I).Denominacion = ""
                      GanLiq(I).ingAp(b).segsocial = ""
                      GanLiq(I).ingAp(b).sind = ""
                      GanLiq(I).ingAp(b).ganbrut = ""
                      GanLiq(I).ingAp(b).ajuste = ""
                      GanLiq(I).ingAp(b).retgan = ""
                      GanLiq(I).ingAp(b).retribNoHab = ""
             Next b

        Next I
    End If
   ' GanLiq_Vacio = True
    Dim MMensual
    '--------------------------------- Inserta Deducciones en Desmen ------------------------------------------
    If Not Deducciones_Vacio Then
        For I = 0 To (UBound(Deducciones) - 1)
        desano = periodo
         For b = 0 To (UBound(Deducciones(I).periodosdeduc) - 1)
                If Deducciones(I).periodosdeduc(b).periodo_mesDesde <> Deducciones(I).periodosdeduc(b).periodo_meshasta Then
                    Meses = (Deducciones(I).periodosdeduc(b).periodo_meshasta - Deducciones(I).periodosdeduc(b).periodo_mesDesde) + 1
                    desmondec = CDbl(Deducciones(I).periodosdeduc(b).periodo_montoMensual) * CDbl(Meses)
                    'desmondec = Replace(desmondec, ",", ".")
                    desmenprorra = -1
                Else
                     If Deducciones(I).periodosdeduc(b).periodo_montoMensual <> "" Then 'MDF 28/07/2015
                        desmondec = Deducciones(I).periodosdeduc(b).periodo_montoMensual
                     Else
                        desmondec = Deducciones(I).MontoTotal
                     End If
                     desmenprorra = 0
                End If
         If Deducciones(I).tipo = 99 Then
                desmondec = Deducciones(I).MontoTotal
                desfecdes = "01/" & Deducciones(I).Mes & "/" & periodo
                desfechas = UltimoDiaMes(periodo, Deducciones(I).Mes)
         Else
         'FB - Si la deduccion no es de tipo 10
            If Deducciones(I).tipo <> 10 Then
              If Deducciones(I).periodosdeduc(b).periodo_mesDesde <> "" Then 'MDF 28/07/2015
                desfecdes = "01/" & Deducciones(I).periodosdeduc(b).periodo_mesDesde & "/" & periodo
                desfechas = UltimoDiaMes(periodo, Deducciones(I).periodosdeduc(b).periodo_meshasta)
              Else 'MDF 28/07/2015
                 desfecdes = "01/" & Deducciones(I).Mes_periodo & "/" & periodo
                 desfechas = UltimoDiaMes(periodo, Deducciones(I).Mes_periodo)
              End If
            End If
         'FB
         End If
         
         NroDoc = Deducciones(I).NroDoc
         Denominacion = Deducciones(I).Denominacion
        
        'FB - Si la deduccion es de tipo 10
         If Deducciones(I).tipo = 10 Then
                desmondec = Deducciones(I).MontoTotal
                desfecdes = "01/01/" & periodo
                desfechas = "31/12/" & periodo
                desmenprorra = -1
                NroDoc = ""
                'Denominacion = Deducciones(I).detallesdeduc(b).detalle_valor
                Denominacion = Deducciones(I).detallesdeduc(1).detalle_valor
         End If
         'FB
            StrSql = "SELECT itenro FROM item_572 WHERE itenro572 = " & Deducciones(I).tipo
            OpenRecordset StrSql, rsConsult
        
            If Not rsConsult.EOF Then
                itenro = rsConsult!itenro
            Else
                itenro = 0
            End If
            rsConsult.Close
            'FB - Como el item 23 queda configurado como de tipo Ganancia Neta (haber), debe grabar el monto en negativo en la tabla desmen
            If (itenro = 8) Or (itenro = 13) Or (itenro = 15) Or (itenro = 21) Or (itenro = 9) Or (itenro = 20) Or (itenro = 22) Or (itenro = 28) Or (itenro = 30) Or (itenro = 24) Or (itenro = 27) Or (itenro = 26) Or (itenro = 25) Or (itenro = 23) Then
                desmondec = (-1) * desmondec
            End If
            
            If itenro <> 0 Then
             For T = 0 To UBound(Arr_NroTernro())
                    NroTernro = Arr_NroTernro(T)
                    If NroTernro <> 0 Then
                        StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                        StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                        StrSql = StrSql & "'" & desfechas & "','" & NroDoc & "','" & Denominacion & "')"
                        Flog.writeline ("Deduccion1: ") & StrSql
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline ("Inserto Tipo Deduccion Correctamente")
                    End If
             Next
            Else
                    Flog.writeline ("No existe el tipo Deduccion ") & Deducciones(I).tipo
            End If
            Deducciones(I).periodosdeduc(b).periodo_mesDesde = ""
            Deducciones(I).periodosdeduc(b).periodo_meshasta = ""
            Deducciones(I).periodosdeduc(b).periodo_montoMensual = ""

          Next b

            Deducciones(I).MontoTotal = ""
            Deducciones(I).Mes = ""
            Deducciones(I).tipo = 0
            Deducciones(I).NroDoc = ""
            Deducciones(I).Denominacion = ""
            Deducciones(I).detallesdeduc(1).detalle_valor = ""
        Next I
    End If
    Deducciones_Vacio = True
    '----------------------------------------------------------------------------------------------------------
    
    '--------------------------------- Inserta Retencion de Pagos en Desmen -----------------------------------
    If Not retPerPagos_Vacio Then
        For I = 0 To (UBound(retPerPagos) - 1)
        desano = periodo
         For b = 0 To (UBound(retPerPagos(I).periodosretpagos) - 1)
                If retPerPagos(I).periodosretpagos(b).periodo_mesDesde <> retPerPagos(I).periodosretpagos(b).periodo_meshasta Then
                    Meses = (retPerPagos(I).periodosretpagos(b).periodo_meshasta - retPerPagos(I).periodosretpagos(b).periodo_mesDesde) + 1
                    desmondec = retPerPagos(I).periodosretpagos(b).periodo_montoMensual * CDbl(Meses)
                    desmenprorra = -1
                Else
                    desmondec = retPerPagos(I).periodosretpagos(b).periodo_montoMensual
                    desmenprorra = 0
                End If
                desfecdes = "01/" & retPerPagos(I).periodosretpagos(b).periodo_mesDesde & "/" & periodo
                desfechas = UltimoDiaMes(periodo, retPerPagos(I).periodosretpagos(b).periodo_meshasta)
            NroDoc = retPerPagos(I).NroDoc
            Denominacion = retPerPagos(I).Denominacion
        
            StrSql = "SELECT itenro FROM item_572 WHERE itenro572 = " & retPerPagos(I).tipo
            OpenRecordset StrSql, rsConsult
        
            If Not rsConsult.EOF Then
                itenro = rsConsult!itenro
            Else
                itenro = 0
            End If
            rsConsult.Close
            
            If (itenro = 8) Or (itenro = 13) Or (itenro = 15) Or (itenro = 21) Or (itenro = 9) Or (itenro = 20) Or (itenro = 22) Or (itenro = 28) Or (itenro = 30) Or (itenro = 24) Or (itenro = 27) Or (itenro = 26) Or (itenro = 25) Then
                desmondec = (-1) * desmondec
            End If
            
            If itenro <> 0 And desmondec <> "0" Then
                 For T = 0 To UBound(Arr_NroTernro())
                    NroTernro = Arr_NroTernro(T)
                    If NroTernro <> 0 Then
                        StrSql = "INSERT INTO desmen(itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc) "
                        StrSql = StrSql & "VALUES(" & itenro & "," & NroTernro & "," & desmondec & "," & desmenprorra & "," & desano & ",'" & desfecdes & "',"
                        StrSql = StrSql & "'" & desfechas & "','" & NroDoc & "','" & Denominacion & "')"
                        Flog.writeline ("Retencion: ") & StrSql
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline ("Inserto Ret Per Pago Correctamente")
                    End If
                 Next
            Else
                Flog.writeline ("No existe el tipo Retencion Per Pago ") & retPerPagos(I).tipo
            End If
          Next b
        Next I
    End If
    retPerPagos_Vacio = True
    MyCommitTrans
'Else
        'Texto = ": " & "El legajo no es numerico "
'        Flog.writeline ("No se reconoce ningun empleado con el CUIL informado ")
        'Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
'        InsertaError 1, 8
'End If
'FGZ - 17/12/2014 ---------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------
Fin:
    Set doc = Nothing
    Set nodes = Nothing
    Set node = Nothing
    Exit Sub
'
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error CE - Leer Archivo"
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
End Sub

Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer

separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        If pos2 > 0 Then
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
    End If
End If
End Sub

Public Sub ComenzarTransferencia()
    
    'Leo los datos del Sistema
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    'FGZ - 17/12/2014 -------------------------------------------------------------------
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = directorio & Trim(objRs!modarchdefault)
        'Directorio = "\\rhdesa\Fuentes\4000_rhprox2_r4\In-Out"
        separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & NroModelo & " " & objRs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de importación:" & directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    'If Right(directorio, 1) = "\" Then
    '    directorio = directorio & "Form572"
    'Else
    '    directorio = directorio & "\Form572"
    'End If
    'FGZ - 17/12/2014 -------------------------------------------------------------------
    
        Flog.writeline "Directorio a buscar :  " & directorio
        Dim fc, F1, s2
        
        Set Folder = fs.GetFolder(directorio)
        Set CArchivos = Folder.Files

        'Determino la proporcion de progreso
        Progreso = 0
        
        HuboError = False
        HuboErrorLocal = False
        cantArch = CArchivos.Count
        
        If CLng(cantArch) < 1 Then
            Flog.writeline "No se encontraron archivos a procesar."
        End If
        
        For Each Archivo In CArchivos
            If UBound(Split(CStr(Archivo.Name), ".")) > 0 Then
                If Split(CStr(Archivo.Name), ".")(1) = "xml" Then
                    NombreArchivo = CStr(Archivo.Name)
                    NArchivo = directorio & "\" & NombreArchivo
                    Flog.writeline "----------------------------------------------------------"
                    Flog.writeline "Archivo Procesado: " & NombreArchivo
                    Flog.writeline "----------------------------------------------------------"
                    HuboErrorLocal = False
                    
                    Call LeeArchivo(NArchivo)
                    
                    Flog.writeline "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
                    
                    'Borrar el archivo
                    If HuboErrorLocal Then
                        'Mantengo el archivo
                        Flog.writeline "Error: Se encontro un error al procesar el archivo." & NArchivo
                        Flog.writeline "Error: El archivo no se movera a la carpeta BK."
                    Else
                        If fs.FileExists(NArchivo) Then
                            
                            On Error Resume Next
                            If InStr(Folder, "/") > 0 Then
                                'fs.MoveFile NArchivo, Folder & "/bk/"
                                fs.MoveFile directorio & "/" & NombreArchivo, directorio & "/bk/" & Split(CStr(NombreArchivo), ".")(0) & "_" & NroProcesoBatch & "." & Split(CStr(NombreArchivo), ".")(1)
                            Else
                                fs.MoveFile directorio & "\" & NombreArchivo, directorio & "\bk\" & Split(CStr(NombreArchivo), ".")(0) & "_" & NroProcesoBatch & "." & Split(CStr(NombreArchivo), ".")(1)
                            End If
                            If Err.Number <> 0 Then
                                Flog.writeline "Error: Se produjo un error al querer mover el archivo." & NArchivo
                                Flog.writeline "Error: El archivo no se movera a la carpeta BK."
                                Flog.writeline "Error: " & Err.Description
                            Else
                                Flog.writeline "Moviendo archivo:" & NArchivo
                            End If
                            Err.Clear
                        End If
                    End If
                    HuboErrorLocal = False
                    
                End If
            End If
            On Error GoTo ErroroTransferencia
        Next
    Exit Sub
ErroroTransferencia:
    HuboError = True
    Flog.writeline "Error Transferencia."
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    Err.Clear
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub
