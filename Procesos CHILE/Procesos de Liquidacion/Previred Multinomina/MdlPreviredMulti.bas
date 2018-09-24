Attribute VB_Name = "MdlPreviredMulti"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "10/03/2010"
'Autor = Martin Ferraro - Se Genero la version del previred multinomina a partir del previred normal

'Const Version = "1.01"
'Const FechaVersion = "07/05/2010"
'Correccion en RUT

'Const Version = "1.02"
'Const FechaVersion = "13/05/2010"
'Martin - Corrcciones para santander que no tiene los procesos para la empresa
'A las modificaciones que realice le puse 13/05/2010 - SANT, buscar eso para ver que se modifico
'Tb se corrigio problema de apert por CCAF y no imprimir cabeceras de que no tienen nominas

'Const Version = "1.03"
'Const FechaVersion = "20/05/2010"
'Martin - En REGINO se reemplazaron dos espacios en blanco despues del RUT por dos 00

'Const Version = "1.04"
'Const FechaVersion = "31/05/2010"
'Martin - Columna 85 y 87 conceptos no sumar. Poner solo el que coincide con la CCAF del empleado
'                                        Montos de la ccaf en valor absoluto

'Const Version = "1.05"
'Const FechaVersion = "07/06/2010"
'Martin - Cuando abre por CCAF mostrar el dlcant como codigo de CCAF y Tipo pago, periodo desde/hasta en vacio

'Const Version = "1.06"
'Const FechaVersion = "08/06/2010"
'Martin - Cuando abre por CCAF la columna 14 va 02

'Const Version = "1.07"
'Const FechaVersion = "09/06/2010"
'Martin - se volvio para atras cambios de version 1.05
'         En lineas que abre CCAF tipo pago Fijo 1

'Const Version = "1.08"
'Const FechaVersion = "21/10/2010"
'Martin - Se agrego Manejador de errores en imprimirarchivo y previmultinomina
'         Ahora para hacer apertura por CCAF busco el codigo previred del empleado
'         y lo comparo con el parametro imprimible de los conceptos que liquidan CCAF.
'         Si son distintos hace apertura

'Const Version = "1.09"
'Const FechaVersion = "28/10/2010"
'Martin - arregloEstruc(62) - Ahora verifica que si no esta cargado arreglo(62) antes

'Const Version = "1.10"
'Const FechaVersion = "09/08/2010"
'Martin - aux85 y aux87

'Const Version = "1.11"
'Const FechaVersion = "11/08/2010"
'Martin - Correccion de sql de busqueda de licencias

'Const Version = "1.12"
'Const FechaVersion = "18/08/2010"
'Martin - Problema que imprimia fechas de apertura de apv nulas

'Const Version = "1.13"
'Const FechaVersion = "06/10/2010"
'Martin Ferraro - Cuando es apertura por APVI o APVC. codigo Movimiento de personal Fijo 0

'Const Version = "1.14"
'Const FechaVersion = "01/12/2010"
'MB - Correcion de gratificaciones

'Const Version = "1.15"
'Const FechaVersion = "28/12/2010"
'Martin Ferraro - En gratificaciones campo 84 renta impo ccaf estaba en cero, ahora arreglo(84)

'Const Version = "1.16"
'Const FechaVersion = "18/01/2012"
'           FGZ - Se le agregó el codigo para string de conexion encriptado y ademas se reordenó la creacion de archivo de log.

'Const Version = "1.17"
'Const FechaVersion = "06/09/2012"
'           Sebastian Stremel - modificaciones en ccaf - CAS-16634 - BANCO SANTANDER CHILE - Error de CCAF Previred multi y mononomina

'Const Version = "1.18"
'Const FechaVersion = "19/09/2012"
'           Sebastian Stremel - modificacion en la empresa cuando es gratificacion

'Const Version = "1.19"
'Const FechaVersion = "03/10/2012"
''           Sebastian Stremel - modificaciones varias para el caso CAS-17155 -
''           BANCO SANTANDER CHILE- ERROR EN GENERACION DEL PREVIRED MULTINOMINA --

'Const Version = "1.20"
'Const FechaVersion = "31/05/2011"
'           FGZ - CAS-19856 - RHPro Consulting - ACS - Error Previred
'           Se agregó control por Licencia de codigo 11

Const Version = "1.21"
Const FechaVersion = "28/10/2014"
'           Carmen Quintero - CAS-26972 - H&A - Bugs detectados en R4 - Error en el Mapeo de Documentos en el reporte Previred
'           Se agregó relacion con la tabla tipodocu_pais al momento de buscar el nro de documento de un empleado


'---------------------------------------------------------------
Private Type TregAPV
    Cod As Long
    Contrato As String
    FPago As Long
    Cotiza As Double
    Depositos As Double
End Type


Global ProgresoxEmpresa As Double

Global CantEmplError '08-03-07 Diego Rosso
Global CantEmplSinError '08-03-07 Diego Rosso
Global Errores As Boolean '08-03-07 Diego Rosso
Global ListaConcCcaf As String
Global TipoEstrCcaf As Long
Global CodCCAF As Long
Global CantEmpr As Long
Global Num_linea As Long




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte Previred.
' Autor      : Martin Ferraro
' Fecha      : 10/03/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
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

    Nombre_Arch = PathFLog & "Previred_Multinomina" & "-" & NroProcesoBatch & ".log"
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
    
    'On Error Resume Next
    'Abro la conexion
    'OpenConnection strconexion, objConn
    'If Err.Number <> 0 Then
    '    Flog.writeline "Problemas en la conexion"
    '    Exit Sub
    'End If
    'OpenConnection strconexion, objconnProgreso
    'If Err.Number <> 0 Then
    '    Flog.writeline "Problemas en la conexion"
    '    Exit Sub
    'End If
    'On Error GoTo 0
    
    On Error GoTo ME_Main
    
    'Nombre_Arch = PathFLog & "Previred_Multinomina" & "-" & NroProcesoBatch & ".log"
    
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    'PID = GetCurrentProcessId
    'Flog.writeline "-------------------------------------------------"
    'Flog.writeline "Version                  : " & Version
    'Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    'Flog.writeline "PID                      : " & PID
    'Flog.writeline "-------------------------------------------------"
    'Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 263 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call PreviredMulti(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    'Flog.writeline
    'Flog.writeline "**********************************************************"
    'Flog.writeline
    'Flog.writeline "Cantidad de Empleados Insertados: " & CantEmplSinError
    'Flog.writeline "Cantidad de Empleados Con ERRORES: " & CantEmplError
    'Flog.writeline
    'Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    'Flog.writeline
    'Flog.writeline "**********************************************************"
    If Not Errores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
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
    'MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
End Sub



Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
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


Public Sub PreviredMulti(ByVal Bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del Previred
' Autor      : Diego Rosso
' Fecha      : 20/01/2007
' --------------------------------------------------------------------------------------------
Dim ListaEmpresas As String
Dim Lista_Pro As String
Dim PliqNro As Long
Dim ArrParametros
Dim Lote As String
Dim RUTPrevi As String
Dim DVPrevi As String
Dim MailPrevi As String
Dim ArrEmpresas
Dim TipoNomina As Long

Dim rs_Consult As New ADODB.Recordset

Dim Directorio As String
Dim Archivo As String
Dim carpeta
Dim fExport
Dim Linea As String
Dim PeriAnio As String
Dim PeriMes As String
Dim Indice As Long
Dim EmpNombre As String
Dim EMPternro As Long
Dim EmpEstrNro As Long
Dim RUT As String
Dim PMReginoNro As Long
Dim DomicilioEmp As String
Dim CantReg As Long
Dim Lista_ProEmpr As String
Dim Fechadesde As Date
Dim Fechahasta As Date
Dim DescFinNomina As String
Dim FechadesdeFiltro As Date
Dim FechahastaFiltro As Date

On Error GoTo ErrPreviredMulti

' El formato de los parametros pasados es
'l_parametros = l_periodo & "@*@" & l_empresa & "@*@" & l_pronro & "@*@" & l_lote & "@*@" & l_rut & "@*@" & l_dv & "@*@" & l_mail

Flog.writeline "Levantando Parametros  "
Flog.writeline Espacios(Tabulador * 1) & Parametros
Flog.writeline
If Not IsNull(Parametros) Then
    
    ArrParametros = Split(Parametros, "@*@")
    If UBound(ArrParametros) = 7 Then
        PliqNro = CLng(ArrParametros(0))
        ListaEmpresas = CStr(ArrParametros(1))
        Lista_Pro = CStr(ArrParametros(2))
        Lote = CStr(ArrParametros(3))
        RUTPrevi = CStr(ArrParametros(4))
        DVPrevi = CStr(ArrParametros(5))
        MailPrevi = CStr(ArrParametros(6))
        TipoNomina = CLng(ArrParametros(7))
        
        Select Case TipoNomina
            Case 1:
                Flog.writeline Espacios(Tabulador * 1) & "Tipo de Nomina Remuneracion del mes"
            Case 2:
                Flog.writeline Espacios(Tabulador * 1) & "Tipo de Nomina Gratificacion"
            Case 3:
                Flog.writeline Espacios(Tabulador * 1) & "Tipo de Nomina Bono Ley Modernización Emp. Públicas"
            Case Else:
                Flog.writeline Espacios(Tabulador * 1) & "Tipo de Nomina Desconocida"
        End Select
        Flog.writeline Espacios(Tabulador * 1) & "Periodo " & PliqNro
        Flog.writeline Espacios(Tabulador * 1) & "Lista de Empresas " & ListaEmpresas
        Flog.writeline Espacios(Tabulador * 1) & "Lista de Procesos " & Lista_Pro
        Flog.writeline Espacios(Tabulador * 1) & "RUT Usuario Previred " & RUTPrevi
        Flog.writeline Espacios(Tabulador * 1) & "Digito Verificador Usuario Previred " & DVPrevi
        Flog.writeline Espacios(Tabulador * 1) & "E-mail " & MailPrevi
        
        'Analizo la cantidad de empresas a procesar
        ArrEmpresas = Split(ListaEmpresas, ",")
        ProgresoxEmpresa = 99 / UBound(ArrEmpresas)
        
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Cantidad de empresas a procesar " & UBound(ArrEmpresas)
                
        '-------------------------------------------------------------------------------------
        'Busco Año y mes del periodo
        '-------------------------------------------------------------------------------------
        StrSql = "SELECT pliqnro, pliqdesc, pliqanio, pliqmes, pliqdesde, pliqhasta FROM periodo WHERE pliqnro = " & PliqNro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            PeriAnio = rs_Consult!pliqanio
            PeriMes = rs_Consult!pliqmes
            Fechadesde = rs_Consult!pliqdesde
            FechadesdeFiltro = rs_Consult!pliqdesde
            Fechahasta = rs_Consult!pliqhasta
            FechahastaFiltro = rs_Consult!pliqhasta
        End If
        rs_Consult.Close
        
        
        
        
        '-------------------------------------------------------------------------------------
        'Configuro el directorio de salida
        '-------------------------------------------------------------------------------------
        StrSql = "SELECT sis_dirsalidas FROM sistema"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            Directorio = Trim(rs_Consult!sis_dirsalidas)
        End If
        rs_Consult.Close
        
        Directorio = Directorio & "\Previred_Multi"
        
        Flog.writeline
        'Activo el manejador de errores
        On Error Resume Next
        'Archivo para la cabecera del Pedido de Pago
        Archivo = Directorio & "\Prev_Multi" & Bpronro & ".txt"
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set fExport = fs.CreateTextFile(Archivo, True)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
            Set carpeta = fs.CreateFolder(Directorio)
            Set fExport = fs.CreateTextFile(Archivo, True)
        End If
        
        Flog.writeline Espacios(Tabulador * 1) & "Se genero el archivo " & Directorio & "\Prev_Multi" & Bpronro & ".txt"
        
        'desactivo el manejador de errores
        On Error GoTo 0
        
        ListaConcCcaf = "'0'"
        TipoEstrCcaf = 0
        CodCCAF = 0
        StrSql = "SELECT * FROM confrep WHERE repnro = 280"
        OpenRecordset StrSql, rs_Consult
        Do While Not rs_Consult.EOF
            
            Select Case CLng(rs_Consult!confnrocol)
                Case 1:
                    If Not EsNulo(rs_Consult!confval2) Then
                        'Concecptos que liquidan CCaf
                        ListaConcCcaf = ListaConcCcaf & ",'" & rs_Consult!confval2 & "'"
                    End If
                Case 2:
                    TipoEstrCcaf = rs_Consult!confval
                Case 3:
                    CodCCAF = rs_Consult!confval
            End Select
            
            rs_Consult.MoveNext
        Loop
        rs_Consult.Close
        
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Configuracion del reporte 280:"
        If ListaConcCcaf = "0" Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontraron configurados conceptos de ccaf en la columna 1"
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Lista de conceptos de Ccaf " & ListaConcCcaf
        End If
        If TipoEstrCcaf = 0 Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encuentra configurado Tipo de estructura ccaf en la columna 2"
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Tipo de estructura Ccaf " & TipoEstrCcaf
        End If
        If CodCCAF = 0 Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encuentra configurado Tipo de codigo ccaf en la columna 3"
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Tipo de codigo Ccaf " & CodCCAF
        End If
        
        Progreso = 0
        CantEmpr = 0
        '-------------------------------------------------------------------------------------
        'REGISTRO ENCABEZADO DE ARCHIVO DE MULTINÓMINA
        '-------------------------------------------------------------------------------------
        StrSql = "INSERT INTO PMRegimn"
        StrSql = StrSql & "(bpronro,lote,RUTUP"
        StrSql = StrSql & ",DV,mail,nominas)"
        StrSql = StrSql & "VALUES("
        StrSql = StrSql & "  " & Bpronro
        StrSql = StrSql & ",'" & Left(Lote, 25) & "'"
        StrSql = StrSql & ",'" & Left(RUTPrevi, 11) & "'"
        StrSql = StrSql & ",'" & Left(DVPrevi, 1) & "'"
        StrSql = StrSql & ",'" & Left(MailPrevi, 50) & "'"
        StrSql = StrSql & ",0"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Por cada empresa imprimo los encabezado
        For Indice = 1 To UBound(ArrEmpresas)
            
            '-------------------------------------------------------------------------------------
            'REGISTRO ENCABEZADO DE NÓMINA
            '-------------------------------------------------------------------------------------
            EmpNombre = ""
            StrSql = "SELECT ternro, empnom, estrnro FROM empresa WHERE empnro = " & ArrEmpresas(Indice)
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                EmpNombre = rs_Consult!empnom
                EMPternro = rs_Consult!ternro
                EmpEstrNro = rs_Consult!estrnro
            End If
            rs_Consult.Close
            
            
            RUT = "00"
            StrSql = "SELECT nrodoc FROM ter_doc WHERE tidnro = 1 AND ternro = " & EMPternro
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                RUT = rs_Consult!nrodoc
                RUT = Replace(RUT, "-", "")
                RUT = Replace(RUT, "/", "")
                RUT = Replace(RUT, "\", "")
            End If
            rs_Consult.Close
            
            DomicilioEmp = ""
            StrSql = "SELECT detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto"
            StrSql = StrSql & " FROM cabdom"
            StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
            StrSql = StrSql & " WHERE domdefault = -1 AND ternro = " & EMPternro
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                DomicilioEmp = IIf(EsNulo(rs_Consult!calle), "", rs_Consult!calle)
                DomicilioEmp = DomicilioEmp & IIf(EsNulo(rs_Consult!nro), "", " " & rs_Consult!nro)
                DomicilioEmp = DomicilioEmp & IIf(EsNulo(rs_Consult!piso), "", " Piso " & rs_Consult!piso)
                DomicilioEmp = DomicilioEmp & IIf(EsNulo(rs_Consult!oficdepto), "", " Dpto. " & rs_Consult!oficdepto)
            End If
            rs_Consult.Close
                        
                        
            StrSql = "INSERT INTO PMRegino"
            StrSql = StrSql & "(bpronro,identificador,nomina"
            StrSql = StrSql & ",RUTPag,DVPag,TipoNom,CodForm"
            StrSql = StrSql & ",Periodo,CantReg,rol,mail)"
            StrSql = StrSql & "VALUES("
            StrSql = StrSql & "  " & Bpronro
            StrSql = StrSql & ",'" & Left(EmpNombre, 25) & "'"
            StrSql = StrSql & ",'" & Left("Periodo de Pago " & Format(PeriAnio, "0000") & Format(PeriMes, "00"), 50) & "'"
            'StrSql = StrSql & ",'" & Left(RUT, 11) & "'"
            If Len(RUT) > 0 Then
                StrSql = StrSql & ",'" & Mid(RUT, 1, Len(RUT) - 1) & "'"
                StrSql = StrSql & ",'" & Right(RUT, 1) & "'"
            Else
                StrSql = StrSql & ",'0'"
                StrSql = StrSql & ",'0'"
            End If
            StrSql = StrSql & ",'" & TipoNomina & "'"
            StrSql = StrSql & ",'1'"
            StrSql = StrSql & ",'" & Format(PeriAnio, "0000") & Format(PeriMes, "00") & "'"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",'TE'"
            StrSql = StrSql & ",'" & Left(MailPrevi, 50) & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            PMReginoNro = getLastIdentity(objConn, "PMRegino")
            
            'Actualizo la cantidad de nominas
            CantEmpr = CantEmpr + 1
            
            'Lista de procesos de la empresa en cuestion
            Lista_ProEmpr = "0"
            StrSql = "SELECT pronro "
            StrSql = StrSql & " FROM proceso "
            StrSql = StrSql & " WHERE pronro IN (" & Lista_Pro & ")"
            StrSql = StrSql & " AND empnro = " & ArrEmpresas(Indice)
            OpenRecordset StrSql, rs_Consult
            Do While Not rs_Consult.EOF
                Lista_ProEmpr = Lista_ProEmpr & "," & rs_Consult!pronro
                rs_Consult.MoveNext
            Loop
            rs_Consult.Close
            Lista_ProEmpr = Lista_Pro
            
            DescFinNomina = Mid(RUT & " " & EmpNombre & " " & DomicilioEmp, 1, 555)
            
            '-------------------------------------------------------------------------------------
            'PREVIRED ESTANDARD
            '-------------------------------------------------------------------------------------
            '13/05/2010 - SANT
            'Call Previred(PMReginoNro, "", Lista_ProEmpr, ArrEmpresas(Indice), Fechadesde, Fechahasta, TipoNomina, DescFinNomina)
            
            'seba 03/10/2012
            If TipoNomina = 2 Then
                StrSql = " select p.pliqnro, p.pliqdesc, p.pliqdesde, p.pliqhasta from impuni_peri ip "
                StrSql = StrSql & " inner join periodo p ON ip.pliqnro = p.pliqnro "
                StrSql = StrSql & " Where ip.pronro IN (" & Lista_Pro & ")"
                StrSql = StrSql & " ORDER BY p.pliqhasta desc "
                OpenRecordset StrSql, rs_Consult
                If Not rs_Consult.EOF Then
                    Fechahasta = rs_Consult!pliqhasta
                End If
                rs_Consult.Close
               
                
                StrSql = " select p.pliqnro, p.pliqdesc, p.pliqdesde, p.pliqhasta from impuni_peri ip "
                StrSql = StrSql & " inner join periodo p ON ip.pliqnro = p.pliqnro "
                StrSql = StrSql & " Where ip.pronro IN (" & Lista_Pro & ") "
                StrSql = StrSql & " ORDER BY p.pliqdesde "
                OpenRecordset StrSql, rs_Consult
                If Not rs_Consult.EOF Then
                    Fechadesde = rs_Consult!pliqdesde
                End If
                rs_Consult.Close
                
            End If
            'hasta aca
            
            '//sebastian stremel si nomina = 2 busco fecha de impuni_peri
'            If TipoNomina = 2 Then
'                StrSql = "SELECT * FROM proceso WHERE pliqnro = " & PliqNro
'                StrSql = StrSql & " AND pronro IN (" & Lista_Pro & ")"
'                OpenRecordset StrSql, rs_Consult
'                If Not rs_Consult.EOF Then
'                    StrSql = "select MAX(pliqnro) periodo from impuni_peri "
'                    StrSql = StrSql & " Where pronro = " & rs_Consult!pronro
'                    OpenRecordset StrSql, rs_Consult
'                    If Not rs_Consult.EOF Then
'                        StrSql = "SELECT pliqdesde,pliqhasta FROM periodo where pliqnro=" & rs_Consult!periodo
'                        OpenRecordset StrSql, rs_Consult
'                        If Not rs_Consult.EOF Then
'                            Fechahasta = rs_Consult!pliqhasta
'                            Fechadesde = rs_Consult!pliqdesde
'                        End If
'                    End If
'                End If
'            End If
            '//hasta aca
            
            Call Previred(PMReginoNro, "", Lista_Pro, ArrEmpresas(Indice), Fechadesde, Fechahasta, TipoNomina, DescFinNomina, EmpEstrNro, FechadesdeFiltro, FechahastaFiltro)
            
            
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Progreso " & Progreso & " %"

            '-------------------------------------------------------------------------------------
            'REGISTRO FIN DE NÓMINA
            '-------------------------------------------------------------------------------------
            If Num_linea > 1 Then
                StrSql = "INSERT INTO PMRegfno"
                StrSql = StrSql & " (bpronro,PMreginonro,descripcion)"
                StrSql = StrSql & " VALUES("
                StrSql = StrSql & "  " & Bpronro
                StrSql = StrSql & " ," & PMReginoNro
                StrSql = StrSql & " ,'" & DescFinNomina & "'"
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Next
        
        '-------------------------------------------------------------------------------------
        'REGISTRO FIN DE ARCHIVO MULTINÓMINA
        '-------------------------------------------------------------------------------------
        If CantEmpr > 0 Then
            StrSql = "INSERT INTO PMRegfmn"
            StrSql = StrSql & " (bpronro,descripcion)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & "  " & Bpronro
            StrSql = StrSql & " ,'" & Mid("Fin Proceso " & Format(PeriAnio, "0000") & Format(PeriMes, "00"), 1, 555) & "'"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Guardo la cantidad de nominas
            StrSql = "UPDATE PMRegimn"
            StrSql = StrSql & " SET nominas = " & CantEmpr
            StrSql = StrSql & " WHERE bpronro = " & Bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            '-------------------------------------------------------------------------------------
            'Impresion de Archivo
            '-------------------------------------------------------------------------------------
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Imprimiendo datos en archivo"
            Call ImprimirArchivo(fExport, Bpronro)
            
        Else
            StrSql = "DELETE PMRegimn"
            StrSql = StrSql & " WHERE bpronro = " & Bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
        
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "No se genereron registros para imprimir"
        End If
        
        
    Else
        Flog.writeline "La cantidad de parametros no es la esperada"
    End If
Else
    Flog.writeline "parametros nulos"
End If


If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

Exit Sub

ErrPreviredMulti:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en PreviredMulti"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
End Sub



Public Sub ImprimirArchivo(ByVal FileArchExp, ByVal Bpronro As Long)

Dim rs_encab As New ADODB.Recordset
Dim rs_reg As New ADODB.Recordset
Dim Linea As String
Dim Monto As String

On Error GoTo ErrImprimirArchivo

    '--------------------------------------------------------------------------------------
    'Encabezado de la multinomina
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM PMRegimn WHERE bpronro = " & Bpronro
    OpenRecordset StrSql, rs_encab
    If Not rs_encab.EOF Then
        
        'REGIMN X(6)
        Linea = "REGIMN"
        'Lote X(25)
        Linea = Linea & FormatTexto(rs_encab!Lote, 25, True, " ")
        'Código de Tipo de Entidad   9(3)
        Linea = Linea & String(3, "0")
        'Código de Entidad   X (3)
        Linea = Linea & String(3, "0")
        'Rut Entidad 9(11)
        Linea = Linea & String(11, "0")
        'Dígito Verfificador de la Entidad   X(1)
        Linea = Linea & String(1, " ")
        'Código de División de la Entidad    X(25)
        Linea = Linea & String(25, " ")
        'Rut Usuario Previred    9(11)
        Linea = Linea & FormatNumero(rs_encab!rutup, 11, True, "0")
        'Dígito Verificador Usuario Previred X(1)
        Linea = Linea & FormatTexto(rs_encab!DV, 1, True, " ")
        'Número de Nóminas   9(10)
        Linea = Linea & FormatNumero(rs_encab!nominas, 10, True, "0")
        'Total Remuneraciones    9(14)
        Linea = Linea & String(14, "0")
        'Total Cotizaciones  9(14)
        Linea = Linea & String(14, "0")
        'E-Mail  X(50)
        Linea = Linea & FormatTexto(rs_encab!mail, 50, True, " ")
        '    X (50)
        Linea = Linea & String(50, " ")
        '    X (25)
        Linea = Linea & String(25, " ")
        '    X (25)
        Linea = Linea & String(25, " ")
        '    X (12)
        Linea = Linea & String(12, " ")
        '    X (6)
        Linea = Linea & String(6, " ")
        '    X (3)
        Linea = Linea & String(3, " ")
        '    X (3)
        Linea = Linea & String(3, " ")
        '    X (1)
        Linea = Linea & String(1, " ")
        '    X (1)
        Linea = Linea & String(1, " ")
        '    9(14)
        Linea = Linea & String(14, "0")
        '    9(14)
        Linea = Linea & String(14, "0")
        '    9(10)
        Linea = Linea & String(10, "0")
        '    9(10)
        Linea = Linea & String(10, "0")
        '    9(5)
        Linea = Linea & String(5, "0")
        '    9(5)
        Linea = Linea & String(5, "0")
        '    999,9999(8)
        Linea = Linea & "000,0000"
        '    999,9999(8)
        Linea = Linea & "000,0000"
        '    X (50)
        Linea = Linea & String(50, " ")
        '    X (25)
        Linea = Linea & String(25, " ")
        '    X (12)
        Linea = Linea & String(12, " ")
        '    X  (12)
        Linea = Linea & String(12, " ")
        '    X (6)
        Linea = Linea & String(6, " ")
        '    X (3)
        Linea = Linea & String(3, " ")
        '    X (3)
        Linea = Linea & String(3, " ")
        '    X (2)
        Linea = Linea & String(2, " ")
        '    9(14)
        Linea = Linea & String(14, "0")
        '    9(14)
        Linea = Linea & String(14, "0")
        '    9(10)
        Linea = Linea & String(10, "0")
        '    9(10)
        Linea = Linea & String(10, "0")
        '    9(5)
        Linea = Linea & String(5, "0")
        '    9(5)
        Linea = Linea & String(5, "0")
        '    999,9999(8)
        Linea = Linea & "000,0000"
        '    999,9999(8)
        Linea = Linea & "000,0000"
        'FINRE X(5)
        Linea = Linea & "FINRE"
        
        FileArchExp.writeline Linea
        
    End If
    rs_encab.Close
    
    '--------------------------------------------------------------------------------------
    'Encabezado de la Nomina
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM PMRegino WHERE bpronro = " & Bpronro
    OpenRecordset StrSql, rs_encab
    Do While Not rs_encab.EOF
    
            'REGINO X(6)
            Linea = "REGINO"
            'Identificador X(25)
            Linea = Linea & FormatTexto(rs_encab!identificador, 25, True, " ")
            'Nombre de la Nómina X (50)
            Linea = Linea & FormatTexto(rs_encab!nomina, 50, True, " ")
            'Rut Pagador 9(11)
            Linea = Linea & FormatNumero(rs_encab!RUTPag, 11, True, "0")
            'Dígito Verificador del Pagador  X(1)
            Linea = Linea & FormatTexto(rs_encab!DVPag, 1, True, " ")
            'Código de División del Pagador  X(25)
            Linea = Linea & "00"
            Linea = Linea & String(23, " ")
            'Código Tipo de Nómina   9(5)
            Linea = Linea & FormatNumero(rs_encab!TipoNom, 5, True, "0")
            'Código de Formato   9(5)
            Linea = Linea & FormatNumero(rs_encab!CodForm, 5, True, "0")
            'Período X(6)
            Linea = Linea & FormatTexto(rs_encab!periodo, 6, True, " ")
            'Número de Registros 9(10)
            Linea = Linea & FormatNumero(rs_encab!CantReg, 10, True, "0")
            'Total Remuneraciones    9(14)
            Linea = Linea & String(14, "0")
            'Total Cotizaciones  9(14)
            Linea = Linea & String(14, "0")
            'Rol X(5)
            Linea = Linea & FormatTexto(rs_encab!rol, 5, True, " ")
            'E-Mail  X (50)
            Linea = Linea & FormatTexto(rs_encab!mail, 50, True, " ")
            '    X  (12)
            Linea = Linea & String(12, " ")
            '    X (6)
            Linea = Linea & String(6, " ")
            '    X (6)
            Linea = Linea & String(6, " ")
            '    X (3)
            Linea = Linea & String(3, " ")
            '    X (3)
            Linea = Linea & String(3, " ")
            '    X (2)
            Linea = Linea & String(2, " ")
            '    X (1)
            Linea = Linea & String(1, " ")
            '    9(14)
            Linea = Linea & String(14, "0")
            '    9(14)
            Linea = Linea & String(14, "0")
            '    9(14)
            Linea = Linea & String(14, "0")
            '    9(10)
            Linea = Linea & String(10, "0")
            '    9(10)
            Linea = Linea & String(10, "0")
            '    999,9999(8)
            Linea = Linea & "000,0000"
            '    999,9999(8)
            Linea = Linea & "000,0000"
            '    X (50)
            Linea = Linea & String(50, " ")
            '    X (25)
            Linea = Linea & String(25, " ")
            '    X (25)
            Linea = Linea & String(25, " ")
            '    X (12)
            Linea = Linea & String(12, " ")
            '    X  (12)
            Linea = Linea & String(12, " ")
            '    X (6)
            Linea = Linea & String(6, " ")
            '    X (6)
            Linea = Linea & String(6, " ")
            '    X (3)
            Linea = Linea & String(3, " ")
            '    X (3)
            Linea = Linea & String(3, " ")
            '    X (2)
            Linea = Linea & String(2, " ")
            '    X (1)
            Linea = Linea & String(1, " ")
            '    9(14)
            Linea = Linea & String(14, "0")
            '    9(14)
            Linea = Linea & String(14, "0")
            '    9(14)
            Linea = Linea & String(14, "0")
            '    9(10)
            Linea = Linea & String(10, "0")
            '    9(10)
            Linea = Linea & String(10, "0")
            '    999,9999(8)
            Linea = Linea & "000,0000"
            '    999,9999(8)
            Linea = Linea & "000,0000"
            'FINRE X(5)
            Linea = Linea & "FINRE"
        
            FileArchExp.writeline Linea
        '--------------------------------------------------------------------------------------
        'Nomina
        '--------------------------------------------------------------------------------------
        StrSql = "SELECT * FROM PMprevired "
        StrSql = StrSql & " WHERE pmreginonro = " & rs_encab!PMReginoNro
        StrSql = StrSql & " ORDER BY num_linea"
        OpenRecordset StrSql, rs_reg
        Do While Not rs_reg.EOF
        
            Linea = FormatNumero(rs_reg!RUT, 11, True, "0")
            Linea = Linea & FormatNumero(rs_reg!DV, 1, True, "0")
            Linea = Linea & FormatTexto(rs_reg!Apellido, 30, True, " ")
            Linea = Linea & FormatTexto(rs_reg!Apellido2, 30, True, " ")
            Linea = Linea & FormatTexto(rs_reg!Nombres, 30, True, " ")
            Linea = Linea & FormatTexto(rs_reg!Sexo, 1, True, " ")
            Linea = Linea & FormatNumero(rs_reg!Nacionalidad, 1, True, "0")
            Linea = Linea & FormatNumero(rs_reg!tipo_pago, 2, True, "0")
            Linea = Linea & FormatNumero(rs_reg!Periodo_desde, 6, True, "0")
            Linea = Linea & FormatNumero(rs_reg!Periodo_hasta, 6, True, "0")
            Linea = Linea & FormatTexto(rs_reg!reg_pre, 3, True, " ")
            Linea = Linea & FormatNumero(rs_reg!TipTrabajador, 1, True, "0")
            Linea = Linea & FormatNumero(rs_reg!DiasTrab, 2, True, "0")
            Linea = Linea & FormatTexto(rs_reg!TipoDeLinea, 2, True, " ")
            Linea = Linea & FormatNumero(rs_reg!CodmovPer, 2, True, "0")
            Linea = Linea & FormatTexto(rs_reg!Fechadesde, 10, True, " ")
            Linea = Linea & FormatTexto(rs_reg!Fechahasta, 10, True, " ")
            Linea = Linea & FormatTexto(rs_reg!TramoAsigFam, 1, True, " ")
            Linea = Linea & FormatNumero(rs_reg!NumCargasSim, 2, True, "0")
            Linea = Linea & FormatNumero(rs_reg!NumCargasMat, 1, True, "0")
            Linea = Linea & FormatNumero(rs_reg!NumCargasInv, 1, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!AsigFam, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 6, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!AsigFamRetro, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 6, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!ReintCarFam, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 6, True, "0")
            
            Linea = Linea & FormatTexto(rs_reg!SolicSubsidioTrabJoven, 1, True, " ")
            Linea = Linea & FormatNumero(rs_reg!CodAFP, 2, True, "0")
            Linea = Linea & FormatNumero(rs_reg!RentaImponibleAFP, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizObligAFP, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!AporteSIS, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!CAVoluntaAFP, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!RenImpSustAFP, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!TasaPact, 2))
            Monto = Replace(Monto, ".", ",")
            Linea = Linea & FormatNumero(Monto, 5, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!AportIndem, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 9, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!NumPeriodos, 2, True, "0")
            Linea = Linea & FormatTexto(rs_reg!PeriDesdeAFP, 10, True, " ")
            Linea = Linea & FormatTexto(rs_reg!PeriHastaAFP, 10, True, " ")
            Linea = Linea & FormatTexto(rs_reg!PuesTrabPesado, 40, True, " ")
            
            Monto = CStr(FormatNumber(rs_reg!PorcCotizTrabPesa, 2))
            Monto = Replace(Monto, ".", ",")
            Linea = Linea & FormatNumero(Monto, 5, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizTrabPesa, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 6, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!InstAutAPV, 3, True, "0")
            Linea = Linea & FormatTexto(rs_reg!NumContratoAPVI, 20, True, " ")
            Linea = Linea & FormatNumero(rs_reg!ForPagAPV, 1, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizAPV, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizDepConv, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!CodInstAutorizadaAPVC, 3, True, "0")
            Linea = Linea & FormatTexto(rs_reg!NumContratoAPVC, 20, True, " ")
            Linea = Linea & FormatNumero(rs_reg!FPagoAPVC, 1, True, "0")
            Linea = Linea & FormatNumero(rs_reg!CotizTrabajadorAPVC, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!CotizEmpleadorAPVC, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!RUTAfVolunt, 11, True, "0")
            Linea = Linea & FormatTexto(rs_reg!DVAfVolunt, 1, True, " ")
            Linea = Linea & FormatTexto(rs_reg!ApePatVolunt, 30, True, " ")
            Linea = Linea & FormatTexto(rs_reg!ApeMatVolunt, 30, True, " ")
            Linea = Linea & FormatTexto(rs_reg!NombVolunt, 30, True, " ")
            Linea = Linea & FormatNumero(rs_reg!CodMovPersVolunt, 2, True, "0")
            Linea = Linea & FormatTexto(rs_reg!FecDesdeVolunt, 10, True, " ")
            Linea = Linea & FormatTexto(rs_reg!FecHastaVolunt, 10, True, " ")
            Linea = Linea & FormatNumero(rs_reg!CodAFPVolunt, 2, True, "0")
            Linea = Linea & FormatNumero(rs_reg!MontoCapVolunt, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!MontoAhorroVolunt, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!NumPerVolunt, 2, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CodCaReg, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 4, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!TasaCotCajPrev, 2))
            Monto = Replace(Monto, ".", ",")
            Linea = Linea & FormatNumero(Monto, 5, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!RentaImpIPS, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizObligINP, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!RentImpoDesah, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!CodCaRegDesah, 4, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!TasaCotDesah, 2))
            Monto = Replace(Monto, ".", ",")
            Linea = Linea & FormatNumero(Monto, 5, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizDesah, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizFonasa, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizAccTrab, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!BonLeyInp, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!DescCargFam, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!BonosGobierno, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!CodInstSal, 2, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!NumFun, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatTexto(Monto, 16, True, " ")
            
            Linea = Linea & FormatNumero(rs_reg!RentaImpIsapre, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!MonPlanIsapre, 1, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizPact, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizObligIsapre, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizAdicVolun, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & String(8, "0")
            
            Linea = Linea & FormatNumero(rs_reg!CodCCAF, 2, True, "0")
            
            Monto = CStr(rs_reg!RentaImponibleCCAF)
            Monto = Replace(Monto, ".", ",")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CredPerCCAF, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!DescDentCCAF, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!DescLeasCCAF, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!DescVidaCCAF, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!OtrosDesCCAF, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotCCAFnoIsapre, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!DesCarFamCCAF, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!OtrosDesCCAF1, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!OtrosDesCCAF2, 8, True, "0")
            Linea = Linea & FormatNumero(rs_reg!BonosGobiernoCCAF, 8, True, "0")
            Linea = Linea & FormatTexto(rs_reg!CodigoSucursalCCAF, 20, True, " ")
            Linea = Linea & FormatNumero(rs_reg!CodMut, 2, True, "0")
            
            Monto = Replace(rs_reg!RentaimpMut, ".", ",")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!CotizAccTrabMut, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!SucPagMut, 3, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!RentTotImp, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!AporTrabSeg, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Monto = CStr(FormatNumber(rs_reg!AporEmpSeg, 0))
            Monto = Replace(Monto, ".", "")
            Monto = Replace(Monto, ",", "")
            Linea = Linea & FormatNumero(Monto, 8, True, "0")
            
            Linea = Linea & FormatNumero(rs_reg!RUTPag, 11, True, "0")
            Linea = Linea & FormatTexto(rs_reg!DVPag, 1, True, " ")
            
            If rs_reg!CentroCosto = "0" Then
                Linea = Linea & String(20, " ")
            Else
                Linea = Linea & FormatTexto(rs_reg!CentroCosto, 20, True, " ")
            End If
            
            FileArchExp.writeline Linea
        
            rs_reg.MoveNext
            
        Loop
        rs_reg.Close
        
        '--------------------------------------------------------------------------------------
        'Fin de la nomina
        '--------------------------------------------------------------------------------------
        StrSql = "SELECT * FROM PMRegfno WHERE PMreginonro = " & rs_encab!PMReginoNro
        OpenRecordset StrSql, rs_reg
        If Not rs_reg.EOF Then
            'REGFMN X(6)
            Linea = "REGFNO"
            '    X (555)
            Linea = Linea & FormatTexto(rs_reg!descripcion, 555, True, " ")
            'FINRE X(5)
            Linea = Linea & "FINRE"
            
            FileArchExp.writeline Linea
            
        End If
        rs_reg.Close
        
        
        rs_encab.MoveNext
    Loop
    rs_encab.Close
    
    
    '--------------------------------------------------------------------------------------
    'Fin de la multinomina
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM PMRegfmn WHERE bpronro = " & Bpronro
    OpenRecordset StrSql, rs_encab
    If Not rs_encab.EOF Then
        'REGFMN X(6)
        Linea = "REGFMN"
        '    X (555)
        Linea = Linea & FormatTexto(rs_encab!descripcion, 555, True, " ")
        'FINRE X(5)
        Linea = Linea & "FINRE"
        
        FileArchExp.writeline Linea
    
    End If
    rs_encab.Close

If rs_reg.State = adStateOpen Then rs_reg.Close
Set rs_reg = Nothing
If rs_encab.State = adStateOpen Then rs_encab.Close
Set rs_encab = Nothing


Exit Sub

ErrImprimirArchivo:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en ImprimirArchivo"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="

End Sub


Public Function FormatTexto(ByVal Str, ByVal Longitud As Long, ByVal Completar As Boolean, ByVal Str_Completar As String)
    If EsNulo(Str) Then
        Str = Str_Completar
    End If
    
    Str = Left(Str, Longitud)
    
    If Completar Then
        If Len(Str) < Longitud Then
            Str = Str & String(Longitud - Len(Str), Str_Completar)
        End If
    End If
    
    FormatTexto = Str
End Function


Public Function FormatNumero(ByVal Str, ByVal Longitud As Long, ByVal Completar As Boolean, ByVal Str_Completar As String)
    If EsNulo(Str) Then
        Str = Str_Completar
    End If
    
    Str = Left(Str, Longitud)
    If Completar Then
        If Len(Str) < Longitud Then
            Str = String(Longitud - Len(Str), Str_Completar) & Str
        End If
    End If
    
    FormatNumero = Str
End Function


Public Sub Previred(ByVal PMReginoNro As Long, ByVal Titulo As String, ByVal Lista_Pro As String, ByVal Empresa As Long, ByVal Fechadesde As Date, ByVal Fechahasta As Date, ByVal tNomina As Integer, ByVal DescPie As String, ByVal EmpEstrNro As Long, ByVal FechadesdeFiltro As Date, ByVal FechahastaFiltro As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del Previred
' Autor      : Diego Rosso
' Fecha      : 20/01/2007
' Modifiacdo : 26/06/2009 - Stankunas Cesar - Se adecuo para el formato de 105 campos
' --------------------------------------------------------------------------------------------
Dim topeArreglo As Integer   'USAR ESTA VARIABLE PARA EL TOPE
Dim arreglo(110) As Double
Dim arregloEstruc(110) As String

Dim arregloMov(30) As Integer
Dim arregloFecD(30) As Date
Dim arregloFecH(30) As Date
Dim total_mov As Integer
Dim aux As Integer

Dim arregloAPVI(30) As TregAPV
Dim arregloAPVC(30) As TregAPV
Dim total_APVI As Long
Dim total_APVC As Long

Dim I As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim UltimoEmpleado As Long
Dim Apellido
Dim Apellido2
Dim NombreEmp
Dim RUT
Dim DV
'Dim Num_linea
Dim Contador
Dim Sexo
Dim TipoPago
Dim EMPternro
Dim FUN
Dim EsFonasa
Dim SeguroCesantia As Boolean
Dim EstaIPS As Boolean

Dim Nacionalidad As Integer
Dim TipoLinea As String
Dim SolicSubsidioJoven As String

Dim CodCabCCAF As Long
Dim ConcCCAFEmp As String
Dim AperCCAF As Boolean
Dim NroLineaCCAF As Long
Dim Aux85 As Double
Dim Aux87 As Double
Dim FechaDesdePeri As Date
Dim FechaHastaPeri As Date
Dim Aux85_nroLinea() As Double
Dim Aux87_nroLinea() As Double
Dim nro_linea_Aux85 As Integer
Dim nro_linea_Aux87 As Integer


'recordsets
Dim rs_Empleados As New ADODB.Recordset
Dim rs_CantEmpleados As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Conceptos As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Rut As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset
Dim rs_Nacionalidad As New ADODB.Recordset
Dim rs_CCAF As New ADODB.Recordset
Dim rs_impuniperi As New ADODB.Recordset
Dim strsql1 As String

 ' Inicio codigo ejecutable
    On Error GoTo CE



' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 2) & "Procesando Previred 105 para"
Flog.writeline Espacios(Tabulador * 3) & "Titulo = " & Titulo
Flog.writeline Espacios(Tabulador * 3) & "Lista_Pro = " & Lista_Pro
Flog.writeline Espacios(Tabulador * 3) & "Empresa = " & Empresa
Flog.writeline Espacios(Tabulador * 3) & "Fecha Desde = " & Fechadesde
Flog.writeline Espacios(Tabulador * 3) & "Fecha Hasta = " & Fechahasta
Flog.writeline Espacios(Tabulador * 3) & "Tipo de Nomina = " & tNomina
Flog.writeline


HuboError = False


'Configuracion del Reporte
StrSql = "SELECT * FROM confrep"
Select Case tNomina
    Case 1:
        StrSql = StrSql & " WHERE repnro = 186 "
    Case 2:
        StrSql = StrSql & " WHERE repnro = 256 "
    Case 3:
        StrSql = StrSql & " WHERE repnro = 257 "
    Case Else:
        StrSql = StrSql & " WHERE repnro = 186 "
End Select
StrSql = StrSql & " AND confrep.confnrocol not in (44,49)"
StrSql = StrSql & " ORDER BY confrep.confnrocol "
        

OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Select Case tNomina
        Case 1:
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró la configuración del Reporte 186"
        Case 2:
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró la configuración del Reporte 256"
        Case 3:
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró la configuración del Reporte 257"
        Case Else:
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró la configuración del Reporte 186"
    End Select
    
    Exit Sub
End If
  
'Esta variable guarda si se creo una cabecera de registro de inicio de nomina de apertura por ccaf. Si es cero no creo cabecera
CodCabCCAF = 0
NroLineaCCAF = 0

'Comienzo la transaccion
'MyBeginTrans




UltimoEmpleado = -1
Num_linea = 1


    
    StrSql = "SELECT empleado.ternro, empleado.empleg, cabliq.cliqnro, cabliq.empleado, proceso.pronro FROM proceso"
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro"
    StrSql = StrSql & " INNER JOIN  tipoproc ON proceso.tprocnro = tipoproc.tprocnro" 'para sacar el ajugcias
    StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro"
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
    
    '13/05/2010 - SANT
    StrSql = StrSql & " AND his_estructura.estrnro = " & EmpEstrNro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fechahasta)
    StrSql = StrSql & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(Fechahasta) & ")"
    StrSql = StrSql & " AND his_estructura.tenro = 10"
    StrSql = StrSql & " WHERE proceso.pronro IN (" & Lista_Pro & ")"
    'StrSql = StrSql & " and empleado.empleg = 83815"
    StrSql = StrSql & " ORDER BY empleado.ternro, proceso.pronro"
    OpenRecordset StrSql, rs_Empleados

    
    If rs_Empleados.State = adStateOpen Then
        'Flog.writeline Espacios(Tabulador * 3) & "busco los empleados"
    Else
        Flog.writeline Espacios(Tabulador * 3) & "se supero el tiempo de espera "
        HuboError = True
    End If


If Not HuboError Then
    
        'seteo de las variables de progreso
        'Progreso = 0
          
          'Obtengo la cantidad real de empleados
            StrSql = "SELECT distinct (empleado) FROM proceso  "
            StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro  "
            StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
            StrSql = StrSql & " Where proceso.pronro IN (" & Lista_Pro & ")"
            OpenRecordset StrSql, rs_CantEmpleados
            
          CEmpleadosAProc = rs_CantEmpleados.RecordCount
        If CEmpleadosAProc = 0 Then
           Flog.writeline Espacios(Tabulador * 3) & "No hay empleados a procesar"
           CEmpleadosAProc = 1
        Else
            Flog.writeline Espacios(Tabulador * 3) & "Cantidad de empleados a procesar " & CEmpleadosAProc
        End If
        IncPorc = (ProgresoxEmpresa / CEmpleadosAProc)
        
        'Inicializo la cantidad de empleados con errores a 0
        CantEmplError = 0
        CantEmplSinError = 0
    Do While Not rs_Empleados.EOF
        Fechahasta = FechahastaFiltro
        Fechadesde = FechadesdeFiltro
        'MyBeginTrans
          rs_Confrep.MoveFirst
          nro_linea_Aux85 = 0
          nro_linea_Aux87 = 0
          total_APVI = 0
          total_APVC = 0
          Aux85 = 0
          Aux87 = 0
          
          ReDim Preserve Aux85_nroLinea(nro_linea_Aux85)
          
          Aux85_nroLinea(nro_linea_Aux85) = 0
          
          ReDim Preserve Aux87_nroLinea(nro_linea_Aux87)
          Aux87_nroLinea(nro_linea_Aux87) = 0
          'Verifico si abre por ccaf
          AperCCAF = False
          If ((ListaConcCcaf <> "'0'") And (TipoEstrCcaf <> 0) And (CodCCAF <> 0)) Then
                
                'Busco el codigo de previred de la ccaf del empleado
                ConcCCAFEmp = ""
                StrSql = "SELECT estr_cod.nrocod "
                StrSql = StrSql & " FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = estructura.estrnro"
                StrSql = StrSql & " AND estr_cod.tcodnro = " & CodCCAF
                StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = " & TipoEstrCcaf & " And "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fechahasta) & ") And "
                StrSql = StrSql & " ((" & ConvFecha(Fechahasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                OpenRecordset StrSql, rs_CCAF
                If Not rs_CCAF.EOF Then
                    If Not EsNulo(rs_CCAF!nrocod) Then ConcCCAFEmp = rs_CCAF!nrocod
                End If
                rs_CCAF.Close

                If Len(ConcCCAFEmp) <> 0 Then
                    
                    'Busco si tiene liquidados conceptos de CCAF distintos al de su CCAF
                    '(el conceptos en su parametro imprimible guarda su codigo previred)
                    StrSql = "SELECT dlimonto Monto, dlicant Cant FROM detliq"
                    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                    StrSql = StrSql & " AND concepto.conccod IN (" & ListaConcCcaf & ")"
                    'StrSql = StrSql & " AND concepto.conccod <> '" & ConcCCAFEmp & "'"
                    StrSql = StrSql & " WHERE cliqnro = " & rs_Empleados!cliqnro
                    StrSql = StrSql & " AND detliq.dlicant <> " & CLng(ConcCCAFEmp)
                    OpenRecordset StrSql, rs_CCAF
                    
                    AperCCAF = Not rs_CCAF.EOF
                    
                End If
                
          End If

          
          If rs_Empleados!ternro <> UltimoEmpleado Then  'Es el primero
                    
               UltimoEmpleado = rs_Empleados!ternro
                
                'Buscar el apellido y nombre
                    StrSql = "SELECT terape, terape2, ternom, ternom2, tersex, nacionalnro FROM tercero WHERE ternro = " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Tercero
                    If Not rs_Tercero.EOF Then
                        Apellido = Left(rs_Tercero!terape, 30)
                        Apellido2 = Left(rs_Tercero!terape2, 30)
                        NombreEmp = Left(rs_Tercero!ternom, 22) & " " & Left(rs_Tercero!ternom2, 6)
                    Else
                        Flog.writeline Espacios(Tabulador * 3) & "ERROR datos del Empleado " & rs_Empleados!ternro
                        GoTo SgtEmpl
                    End If


                    'Inicializar Arreglos de totales
                    For Contador = 1 To 110
                        arreglo(Contador) = 0
                        arregloEstruc(Contador) = ""
                        If Contador <= 30 Then
                            arregloMov(Contador) = 0
                            arregloFecD(Contador) = vbNull
                            arregloFecH(Contador) = vbNull
                        End If
                    Next Contador

          End If
                            
                Do While Not rs_Confrep.EOF
                    
                    Select Case UCase(rs_Confrep!conftipo)
                    Case "AC":
                        StrSql = "SELECT almonto FROM acu_liq WHERE cliqnro = " & rs_Empleados!cliqnro & _
                                 " AND acunro =" & rs_Confrep!confval
                        OpenRecordset StrSql, rs_Acu_liq
                        If Not rs_Acu_liq.EOF Then arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Acu_liq!almonto)
                    Case "CO":
                        StrSql = "SELECT concnro, conccod FROM concepto "
                        StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                        StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then
                            'CCAF
                            If (CLng(rs_Confrep!confnrocol = 85) Or CLng(rs_Confrep!confnrocol = 87)) Then
                                
                                'Solo va concepto de CCAF del empleado
                                
                                    StrSql = "SELECT dlimonto, dlicant FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                             " AND cliqnro =" & rs_Empleados!cliqnro
                                    OpenRecordset StrSql, rs_Detliq
                                    If Not rs_Detliq.EOF Then
                                        If rs_Detliq!dlimonto <> 0 Then
                                            If CLng(ConcCCAFEmp) = rs_Detliq!dlicant Then
                                                'arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                                arreglo(rs_Confrep!confnrocol) = Abs(rs_Detliq!dlimonto)
                                            Else
                                                If (CLng(rs_Confrep!confnrocol = 85)) Then
                                                    'Aux85 = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                                    Aux85 = Abs(rs_Detliq!dlimonto)
                                                    nro_linea_Aux85 = nro_linea_Aux85 + 1
                                                    Flog.writeline "nro_linea_Aux85=" & nro_linea_Aux85
                                                    ReDim Preserve Aux85_nroLinea(nro_linea_Aux85)
                                                    'Aux85 = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                                    'Flog.writeline "Aux85=" & Aux85
                                                    Aux85_nroLinea(nro_linea_Aux85) = Abs(rs_Detliq!dlimonto)
                                                    
                                                    
                                                Else
                                                    'Aux87 = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                                    Aux87 = Abs(rs_Detliq!dlimonto)
                                                    nro_linea_Aux87 = nro_linea_Aux87 + 1
                                                    'Flog.writeline "nro_linea_Aux87=" & nro_linea_Aux87
                                                    ReDim Preserve Aux87_nroLinea(nro_linea_Aux87)
                                                    'Aux87 = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                                    Flog.writeline "Aux87=" & Aux87
                                                    Aux87_nroLinea(nro_linea_Aux87) = Abs(rs_Detliq!dlimonto)
                                                End If
                                            End If
                                        End If
                                    End If
                                
                            Else
                                
                                StrSql = "SELECT dlimonto FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                         " AND cliqnro =" & rs_Empleados!cliqnro
                                OpenRecordset StrSql, rs_Detliq
                                If Not rs_Detliq.EOF Then
                                    If rs_Detliq!dlimonto <> 0 Then
                                        arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                    End If
                                End If
                                
                            End If
                        End If
                    
                    Case "PCO":
                        StrSql = "SELECT concnro FROM concepto "
                        StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                        StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then
                            StrSql = "SELECT dlicant FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                     " AND cliqnro =" & rs_Empleados!cliqnro
                            OpenRecordset StrSql, rs_Detliq
                            If Not rs_Detliq.EOF Then
                                If rs_Detliq!dlicant <> 0 Then arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlicant)
                            End If
                        End If
                    
                    Case "APV":
                        StrSql = "SELECT concnro FROM concepto "
                        StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                        StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then
                            StrSql = "SELECT dlicant, dlimonto FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                     " AND cliqnro =" & rs_Empleados!cliqnro
                            OpenRecordset StrSql, rs_Detliq
                            If Not rs_Detliq.EOF Then
                                If rs_Detliq!dlicant <> 0 Then
                                    'arreglo(rs_Confrep!confnrocol) = Abs(rs_Detliq!dlicant)
                                    'arreglo(rs_Confrep!confnrocol + 3) = Abs(rs_Detliq!dlimonto)
                                    Select Case CLng(rs_Confrep!confnrocol)
                                        Case 40:
                                            'Encontre una institucion liquidada
                                            total_APVI = total_APVI + 1
                                            arregloAPVI(total_APVI).Cod = Abs(rs_Detliq!dlicant)
                                            arregloAPVI(total_APVI).Cotiza = Abs(rs_Detliq!dlimonto)
                                            
                                            'Busco OTRO concepto liquidado asociado a la institucuion (mismo parametro)
                                            StrSql = "SELECT detliq.dlimonto"
                                            StrSql = StrSql & " From detliq"
                                            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                                            StrSql = StrSql & " INNER JOIN confrep ON confrep.confval2 = concepto.conccod"
                                            StrSql = StrSql & " AND confrep.repnro = " & rs_Confrep!repnro
                                            StrSql = StrSql & " AND confrep.confnrocol = 44"
                                            StrSql = StrSql & " WHERE cliqnro = " & rs_Empleados!cliqnro
                                            StrSql = StrSql & " AND detliq.dlicant = " & rs_Detliq!dlicant
                                            OpenRecordset StrSql, rs_Aux
                                            If Not rs_Aux.EOF Then
                                                arregloAPVI(total_APVI).Depositos = Abs(rs_Aux!dlimonto)
                                            Else
                                                'El registro no esta completo
                                                arregloAPVI(total_APVI).Depositos = 0
                                            End If
                                            
                                        Case 45:
                                            'Busco OTRO concepto liquidado asociado a la institucuion (mismo parametro)
                                            total_APVC = total_APVC + 1
                                            arregloAPVI(total_APVC).Cod = Abs(rs_Detliq!dlicant)
                                            arregloAPVI(total_APVC).Cotiza = Abs(rs_Detliq!dlimonto)
                                            
                                            'Busco OTRO concepto liquidado asociado a la institucuion (mismo parametro)
                                            StrSql = "SELECT detliq.dlimonto"
                                            StrSql = StrSql & " From detliq"
                                            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                                            StrSql = StrSql & " INNER JOIN confrep ON confrep.confval2 = concepto.conccod"
                                            StrSql = StrSql & " AND confrep.repnro = " & rs_Confrep!repnro
                                            StrSql = StrSql & " AND confrep.confnrocol = 49"
                                            StrSql = StrSql & " WHERE cliqnro = " & rs_Empleados!cliqnro
                                            StrSql = StrSql & " AND detliq.dlicant = " & rs_Detliq!dlicant
                                            OpenRecordset StrSql, rs_Aux
                                            If Not rs_Aux.EOF Then
                                                arregloAPVC(total_APVC).Depositos = Abs(rs_Aux!dlimonto)
                                            Else
                                                arregloAPVC(total_APVC).Depositos = 0
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                        
                    Case "TE": 'tipo estructura
                            
                            StrSql = " SELECT estructura.estrnro, estructura.estrdabr, estructura.estrcodext, htetdesde, htethasta "
                            StrSql = StrSql & " FROM his_estructura "
                            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                            StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro & " AND "
                            StrSql = StrSql & " his_estructura.tenro = " & rs_Confrep!confval & " And "
                            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fechahasta) & ") And "
                            StrSql = StrSql & " ((" & ConvFecha(Fechahasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                            OpenRecordset StrSql, rs_Estructura
                            If Not rs_Estructura.EOF Then
                                If rs_Confrep!confnrocol = 37 Then
                                    arregloEstruc(37) = rs_Estructura!estrdabr
                                Else
                                    StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                                    StrSql = StrSql & " AND tcodnro = 38"
                                    OpenRecordset StrSql, rs_Estr_cod
                                    If Not rs_Estr_cod.EOF Then arregloEstruc(rs_Confrep!confnrocol) = IIf(EsNulo(rs_Estr_cod!nrocod), "", CStr(rs_Estr_cod!nrocod))
                                    rs_Estr_cod.Close
                                End If
                            End If
                            rs_Estructura.Close
                            
                    Case "TM": 'tipo movimiento
                            'Hacer case de tipo de Movimiento y generar el array con las fechas correspondientes
                            total_mov = 0
                            'ALTA
                                StrSql = "SELECT fases.altfec, fases.bajfec FROM fases "
                                StrSql = StrSql & " WHERE fases.real = -1 "
                                StrSql = StrSql & " AND fases.altfec >=" & ConvFecha(Fechadesde)
                                StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(Fechahasta)
                                StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
                                OpenRecordset StrSql, rs_Fases
                                Do While Not rs_Fases.EOF
                                    total_mov = total_mov + 1
                                    arregloMov(total_mov) = 1   'fijo
                        
                                    If Not EsNulo(rs_Fases!altfec) Then
                                       arregloFecD(total_mov) = rs_Fases!altfec
                                       arregloFecH(total_mov) = rs_Fases!altfec
                                    End If
                        
                                    rs_Fases.MoveNext
                                Loop
                            
                            'BAJA"
                            StrSql = "SELECT fases.caunro, fases.altfec, fases.bajfec FROM fases "
                            StrSql = StrSql & " WHERE fases.real = -1 "
                            StrSql = StrSql & " AND fases.bajfec >=" & ConvFecha(Fechadesde)
                            StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(Fechahasta)
                            StrSql = StrSql & " AND fases.empleado = " & rs_Empleados!ternro
                            OpenRecordset StrSql, rs_Fases
                            Do While Not rs_Fases.EOF
                                total_mov = total_mov + 1
                                'segun la causa ==> busco la estructura y el codigo asociado
                                
                                StrSql = "SELECT nrocod FROM estr_cod "
                                StrSql = StrSql & " INNER JOIN causa_sitrev ON causa_sitrev.estrnro = estr_cod.estrnro"
                                StrSql = StrSql & " WHERE tcodnro = 38"
                                StrSql = StrSql & " AND causa_sitrev.caunro = " & rs_Fases!caunro
                                OpenRecordset StrSql, rs_Estr_cod
                                If Not rs_Estr_cod.EOF Then
                                   arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                End If
                            
                                If Not EsNulo(rs_Fases!bajfec) Then
                                   arregloFecD(total_mov) = rs_Fases!bajfec
                                   arregloFecH(total_mov) = rs_Fases!bajfec
                                End If
                                
                                rs_Fases.MoveNext
                            Loop
                            
                            'Licencias
                            StrSql = " SELECT emp_lic.elfechadesde, emp_lic.elfechahasta, emp_lic.tdnro,emp_lic.emp_licnro FROM emp_lic "
                            StrSql = StrSql & " WHERE emp_lic.empleado= " & rs_Empleados!ternro
                            StrSql = StrSql & " AND (emp_lic.elfechadesde <= " & ConvFecha(Fechahasta)
                            StrSql = StrSql & " AND emp_lic.elfechahasta >= " & ConvFecha(Fechadesde) & ")"
                            
                            OpenRecordset StrSql, rs_Aux
                            Do While Not rs_Aux.EOF
                                
                                'Flog.writeline Espacios(Tabulador * 3) & "LIC ENCONTRADA PARA TERNRO " & rs_Empleados!ternro & " NRO LIC " & rs_Aux!emp_licnro & " DESDE " & rs_Aux!elfechadesde & " HASTA " & rs_Aux!elfechahasta & " TIPO " & rs_Aux!tdnro
                                total_mov = total_mov + 1
                                
                                StrSql = "SELECT nrocod FROM estr_cod "
                                StrSql = StrSql & " INNER JOIN csijp_srtd ON estr_cod.estrnro = csijp_srtd.estrnro "
                                StrSql = StrSql & " WHERE tcodnro = 38"
                                StrSql = StrSql & " AND csijp_srtd.tdnro = " & rs_Aux!tdnro
                                OpenRecordset StrSql, rs_Estr_cod
                                If Not rs_Estr_cod.EOF Then
                                   arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                End If
                                
                                If rs_Aux!elfechadesde < Fechadesde Then
                                    arregloFecD(total_mov) = Fechadesde
                                Else
                                    arregloFecD(total_mov) = rs_Aux!elfechadesde
                                End If
                                If rs_Aux!elfechahasta > Fechahasta Then
                                    arregloFecH(total_mov) = Fechahasta
                                Else
                                    arregloFecH(total_mov) = rs_Aux!elfechahasta
                                End If
                            
                                rs_Aux.MoveNext
                            Loop
                            
                    Case "CTE": 'Constante
                                If rs_Confrep!confval2 = "" Or EsNulo(rs_Confrep!confval2) Then
                                    'Numerica
                                    arreglo(rs_Confrep!confnrocol) = rs_Confrep!confval
                                Else
                                    'Alfanumerica
                                    arregloEstruc(rs_Confrep!confnrocol) = rs_Confrep!confval2
                                End If
                    Case Else
                    
                    End Select
                    
                    rs_Confrep.MoveNext
                Loop
                
                'Reviso si es el ultimo empleado
                If EsElUltimoEmpleado(rs_Empleados, UltimoEmpleado) Then
                    'Inicializo
                        HuboError = False 'Para cada empleado
                        'Errores = False 'En el proceso
                        EsFonasa = False
                        
                    '-----------------------------------------------------------------------------------
                    'Lleno con ceros las posiciones vacias del arreglo para que NO tire error al insertar
                    For Contador = 1 To 110
                        If arregloEstruc(Contador) = "" Then
                            arregloEstruc(Contador) = "0"
                        End If
                    Next
                    
                    '-----------------------------------------------------------------------------------
                    ' Bloque Datos del Trabajador
                    ' ----------------------------------------------------------------
                    ' Buscar el Rut DEL EMPLEADO
                    '28/10/2014
                    'StrSql = " SELECT nrodoc FROM tercero " & _
                    '         " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = 1) " & _
                    '         " WHERE tercero.ternro= " & rs_Empleados!ternro
                    'Inicio
                    StrSql = " SELECT nrodoc FROM tercero "
                    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro =tercero.ternro "
                    StrSql = StrSql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
                    StrSql = StrSql & " INNER JOIN tipodocu_pais on tipodocu_pais.tidnro = ter_doc.tidnro "
                    StrSql = StrSql & " AND tipodocu_pais.paisnro = 8"
                    StrSql = StrSql & " AND tipodocu_pais.tidcod = 1"
                    StrSql = StrSql & " WHERE Tercero.ternro = " & rs_Empleados!ternro
                    'fin
                    OpenRecordset StrSql, rs_Rut
                    
                    If Not rs_Rut.EOF Then
                        RUT = Mid(rs_Rut!nrodoc, 1, Len(rs_Rut!nrodoc) - 1)
                        RUT = Replace(RUT, "-", "")
                        DV = Right(rs_Rut!nrodoc, 1)
                    Else
                        RUT = ""
                        DV = ""
                        HuboError = True
                    End If
                    
                    'SEXO
                    If Not rs_Tercero.EOF Then
                        If rs_Tercero!tersex = -1 Then
                            Sexo = "M"
                        Else
                            Sexo = "F"
                        End If
                    Else
                        Sexo = ""
                        HuboError = True
                    End If


                    'Busco la Nacionalidad
                    StrSql = "SELECT nacionaldefault FROM nacionalidad"
                    StrSql = StrSql & " WHERE nacionalnro = " & rs_Tercero!nacionalnro
                    OpenRecordset StrSql, rs_Nacionalidad
                    If rs_Nacionalidad!nacionaldefault = -1 Then
                       Nacionalidad = 0
                    Else
                       Nacionalidad = 1
                    End If
                    
                    TipoPago = tNomina
                    
                    If IsNumeric(arregloEstruc(12)) = False Then
                            arregloEstruc(12) = "0"
                            HuboError = True
                    End If
                    
                    'Tipo de Linea
                    'Fijo 00 (Linea principal Default)
                    TipoLinea = "00"
                    
                    'Codigo de Movimiento del Personal
                    aux = 1
                    Do While (aux <= total_mov) And (Not HuboError)
                        If IsNumeric(arregloMov(aux)) = False Then
                                HuboError = True
                                arregloMov(aux) = 0
                        End If
                        
                        'Fecha Desde
                        'Si Movimiento de personal es 1,3,4,5,6,7,8 el campo es obligatorio. QUEDA PARA REVISAR
                        'FGZ - 31/05/2013 ------------------------------------------------------
                        'If Not (arregloMov(aux) = "1" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "5" Or arregloMov(aux) = "6" Or arregloMov(aux) = "7" Or arregloMov(aux) = "8") Then
                        If Not (arregloMov(aux) = "1" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "5" Or arregloMov(aux) = "6" Or arregloMov(aux) = "7" Or arregloMov(aux) = "8" Or arregloMov(aux) = "11") Then
                        'FGZ - 31/05/2013 ------------------------------------------------------
                            'Cuando es un retiro, la fecha del retiro se guarda como desde en RH Pro pero en Previred se informa en Hasta
                            If arregloMov(aux) = "2" Then arregloFecH(aux) = arregloFecD(aux)
                            arregloFecD(aux) = vbNull
                        End If
                        
                        'Fecha Hasta
                        'Si Movimiento de personal es 2,3,4,6 el campo es obligatorio. QUEDA PARA REVISAR
                        'FGZ - 31/05/2013 ------------------------------------------------------
                        'If Not (arregloMov(aux) = "2" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "6") Then arregloFecH(aux) = vbNull
                        If Not (arregloMov(aux) = "2" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "6" Or arregloMov(aux) = "11") Then arregloFecH(aux) = vbNull
                        'FGZ - 31/05/2013 ------------------------------------------------------
                        aux = aux + 1
                    Loop
                    
                    'Tramo de Asignacion Familiar
                    If arreglo(18) = 0 And arregloEstruc(18) = "0" Then
                        arregloEstruc(18) = "D"
                    Else
                        If arreglo(18) <> 0 Then
                            Select Case arreglo(18)
                                Case 1:
                                    arregloEstruc(18) = "A"
                                Case 2:
                                    arregloEstruc(18) = "B"
                                Case 3:
                                    arregloEstruc(18) = "C"
                                Case 4:
                                    arregloEstruc(18) = "D"
                                Case Else
                                    arregloEstruc(18) = "Sin"
                            End Select
                        End If
                        'Flog.writeline Espacios(Tabulador * 3) & "Campo Obtenido"
                    End If
                    'Flog.writeline
                    
                    SolicSubsidioJoven = "N"
                    
                    '-----------------------------------------------------------------------------------
                    ' Bloque de AFP
                    '-----------------------------------------------------------------------------------
                    If (IsNumeric(arregloEstruc(26)) = False) Or (IsNumeric(arreglo(26)) = False) Then
                        arreglo(26) = 0
                        arregloEstruc(26) = "0"
                        HuboError = True
                    End If
                    
                    
                    'Puesto de Trabajo Pesado
                    If arregloEstruc(37) = "0" Then arregloEstruc(37) = ""

                    '-----------------------------------------------------------------------------------
                    'Bloque de APVI
                    '-----------------------------------------------------------------------------------
                   'Numero de Contrato APVI
                   If total_APVI > 0 Then
                        If Not ((arregloEstruc(41) = "0") Or (arregloEstruc(41) = "")) Then
                            If IsNumeric(arregloEstruc(41)) = False Then
                                arregloEstruc(41) = "0"
                                HuboError = True
                            End If
                        End If
                   Else
                        arregloEstruc(41) = "0"
                   End If
                   
                   'Forma de PAGO APVI
                   If total_APVI > 0 Then
                        If arregloEstruc(42) = "0" And arregloEstruc(42) <> "000" Then
                             HuboError = True
                             total_APVI = 0
                         Else
                             If IsNumeric(arregloEstruc(42)) = False Then
                                 arregloEstruc(42) = "0"
                                 HuboError = True
                                 total_APVI = 0
                             End If
                         End If
                   Else
                        arregloEstruc(42) = "0"
                   End If
                   
                   
                    '-----------------------------------------------------------------------------------
                    'Bloque de APVC
                    '-----------------------------------------------------------------------------------
                   'Codigo de la Institucion Autorizada APVC
                   'Validado arriba en confrep
                   
                   'Numero de Contrato APVC
                   If total_APVC > 0 Then
                        Flog.writeline Espacios(Tabulador * 2) & "Procesando Campo 46: Numero de Contrato APVC"
                        If Not ((arregloEstruc(46) = "0") Or (arregloEstruc(46) = "")) Then
                            If IsNumeric(arregloEstruc(46)) = False Then
                                arregloEstruc(46) = "0"
                                HuboError = True
                            End If
                        End If
                    Else
                        arregloEstruc(46) = ""
                    End If
                   
                   'Forma de PAGO APVC
                   If total_APVC > 0 Then
                        If arregloEstruc(47) = "0" And arregloEstruc(47) <> "000" Then
                             HuboError = True
                         Else
                             If IsNumeric(arregloEstruc(47)) = False Then
                                 arregloEstruc(47) = "0"
                                 HuboError = True
                             End If
                         End If
                   Else
                        arregloEstruc(47) = "0"
                   End If
                   
                    
                   '-----------------------------------------------------------------------------------
                   'Bloque de IPS - Fonasa
                   '-----------------------------------------------------------------------------------
                    'Codigo EX caja Regimen
                     If IsNumeric(arregloEstruc(62)) = False Then
                            arregloEstruc(62) = "0"
                            'HuboError = True
                     End If
                     
                    
                   'Cotizacion Obligatoria IPS
                    EstaIPS = (UCase(arregloEstruc(11)) = "IPS")
                    
                    'Codigo Ex caja Regimen Regimen Desahucio
                    If IsNumeric(arregloEstruc(67)) = False Then
                            arregloEstruc(67) = "0"
                            HuboError = True
                     End If
                    
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de ISAPRE
                    '-----------------------------------------------------------------------------------
                    If IsNumeric(arregloEstruc(75)) = False Then
                        arregloEstruc(75) = "0"
                        HuboError = True
                    End If
                    
                    'FUN DEL EMPLEADO
                    '28/10/2014
                    'StrSql = " SELECT nrodoc FROM tercero" & _
                    '         " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro)" & _
                    '         " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro and tipodocu.tidsigla='Fun'" & _
                    '         " WHERE tercero.ternro= " & rs_Empleados!ternro
                    
                    'Inicio
                    StrSql = " SELECT nrodoc FROM tercero "
                    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro =tercero.ternro "
                    StrSql = StrSql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
                    StrSql = StrSql & " INNER JOIN tipodocu_pais tdp ON tipodocu.tidnro = tdp.tidnro AND tdp.paisnro = 8"
                    StrSql = StrSql & " AND UPPER(tipodocu.tidsigla) = '" & UCase("Fun") & "'"
                    StrSql = StrSql & " WHERE Tercero.ternro = " & rs_Empleados!ternro
                    'fin
                    OpenRecordset StrSql, rs_Rut
                    If Not rs_Rut.EOF Then
                        FUN = rs_Rut!nrodoc
                    Else
                        FUN = 0
                    End If
                    
                    'Moneda del Plan Pactado
                    If IsNumeric(arreglo(78)) = False Then arregloEstruc(78) = "0"
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de CCAF
                    '-----------------------------------------------------------------------------------
                    If IsNumeric(arregloEstruc(83)) = False Then arregloEstruc(83) = "0"
                    
                    'Codigo de Sucursal
                    If IsNumeric(arregloEstruc(95)) = False Then
                        arregloEstruc(95) = "0"
                        HuboError = True
                    End If
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de Mutual de Seguridad
                    '-----------------------------------------------------------------------------------
                    
                    'Codigo Mutual
                    If IsNumeric(arregloEstruc(96)) = False Then arregloEstruc(96) = "0"
                    
                    
                   'Sucursal para pago mutual
                    If IsNumeric(arregloEstruc(99)) = False Then arregloEstruc(99) = "0"
                                     
                   'Renta imponible Seguro Cesantia
                    If arreglo(100) < 0 Then HuboError = True
                    
                    SeguroCesantia = False
                    If arreglo(100) <> 0 Then SeguroCesantia = True
                    
                    
                  'Fin Validaciones
                  '-----------------------------------------------------------------------------------
                'Controlo errores en el empleado
                'If Not HuboError Then
                    
                    Select Case TipoPago
                        Case 1, 3: 'Remuneraciones
                           'Inserto en rep_previred
                            StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                            StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                            StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                            StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                            StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                            StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                            
                            StrSql = StrSql & ") VALUES ("
                            
                            StrSql = StrSql & NroProcesoBatch & ","
                            StrSql = StrSql & PMReginoNro & ","
                            StrSql = StrSql & rs_Empleados!ternro & ","
                            StrSql = StrSql & Num_linea & ","
                            StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & Empresa & ","
                            StrSql = StrSql & "'" & RUT & "',"
                            StrSql = StrSql & "'" & DV & "',"
                            StrSql = StrSql & "'" & Apellido & "',"
                            StrSql = StrSql & "'" & Apellido2 & "',"
                            StrSql = StrSql & "'" & NombreEmp & "',"
                            StrSql = StrSql & "'" & Sexo & "',"
                            StrSql = StrSql & "'" & Nacionalidad & "',"
                            StrSql = StrSql & TipoPago & ","
                            StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                            StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                            StrSql = StrSql & arreglo(10) & "," 'Renta imponible
                            StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                            StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                            StrSql = StrSql & arreglo(13) & "," 'Dias Trabajados
                            StrSql = StrSql & TipoLinea & "," 'Tipo de Linea
                            StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                            StrSql = StrSql & IIf(arregloFecD(1) = vbNull, "Null", ConvFecha(arregloFecD(1))) & "," 'Fecha Desde para el movimiento
                            StrSql = StrSql & IIf(arregloFecH(1) = vbNull, "Null", ConvFecha(arregloFecH(1))) & "," 'Fecha Hasta para el movimiento
                            StrSql = StrSql & "'" & arregloEstruc(18) & "'," 'Tramo Asignacion Familiar
                            StrSql = StrSql & arreglo(19) & "," 'Numero de cargas simples
                            StrSql = StrSql & arreglo(20) & "," 'Numero de cargas Maternales
                            StrSql = StrSql & arreglo(21) & "," 'Numero de cargas Invalidas
                            StrSql = StrSql & arreglo(22) & "," 'ASignacion Familiar
                            StrSql = StrSql & arreglo(23) & "," 'ASignacion Familiar Retroactiva
                            StrSql = StrSql & arreglo(24) & "," 'Renta Carga Familiares
                            StrSql = StrSql & "'" & SolicSubsidioJoven & "'," 'Solicitud Subsidio Trabajador Joven
                            If arreglo(26) = 0 Then
                                StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                            Else
                                StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                            End If
                            StrSql = StrSql & arreglo(27) & "," 'Renta Imponible AFP
                            StrSql = StrSql & arreglo(28) & "," 'Cotizacion Obligatoria AFP
                            StrSql = StrSql & arreglo(29) & "," 'Aporte Seguro Invalidez y Sobervivencia
                            StrSql = StrSql & arreglo(30) & "," 'Cuenta de Ahorro Voluntaria
                            StrSql = StrSql & arreglo(31) & "," 'Renta Imponible Sust a AFP
                            StrSql = StrSql & arreglo(32) & "," 'Tasa Pactada
                            StrSql = StrSql & arreglo(33) & "," 'Aporte Indem
                            StrSql = StrSql & 0 & "," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                            StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                            StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                            StrSql = StrSql & "'" & Mid(arregloEstruc(37), 1, 40) & "'," 'Puesto de trabajo Pesado
                            StrSql = StrSql & arreglo(38) & "," 'Porcentaje Cotizacion Trabajo Pesado
                            StrSql = StrSql & arreglo(39) & "," 'Cotizacion Trabajo Pesado
                            If total_APVI > 0 Then
                                StrSql = StrSql & arregloAPVI(1).Cod & "," 'Inst Autor APVI
                                StrSql = StrSql & arregloEstruc(41) & "," 'Numero de contrato APVI
                                StrSql = StrSql & arregloEstruc(42) & "," 'Forma de Pago APVI
                                StrSql = StrSql & arregloAPVI(1).Cotiza & "," 'Cotizacion APVI
                                StrSql = StrSql & arregloAPVI(1).Depositos & "," 'Cotizacion Depositos convenidos
                            Else
                                StrSql = StrSql & "0," 'Inst Autor APVI
                                StrSql = StrSql & "0," 'Numero de contrato APVI
                                StrSql = StrSql & "0," 'Forma de Pago APVI
                                StrSql = StrSql & "0," 'Cotizacion APVI
                                StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                            End If
                            If total_APVC > 0 Then
                                StrSql = StrSql & arregloAPVC(1).Cod & ","  'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & arregloEstruc(46) & ","  'Numero de Contrato APVC
                                StrSql = StrSql & arregloEstruc(47) & ","  'Forma de Pago APVC
                                StrSql = StrSql & arregloAPVC(1).Cotiza & "," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & arregloAPVC(1).Depositos & "," 'Cotizacion Empleador APVC
                            Else
                                StrSql = StrSql & "0,"  'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & "0,"  'Numero de Contrato APVC
                                StrSql = StrSql & "0,"  'Forma de Pago APVC
                                StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                            End If
                            StrSql = StrSql & "'" & "0" & "'," 'RUT Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                            StrSql = StrSql & 0 & "," 'Codigo Movimiento Personal Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                            StrSql = StrSql & 0 & "," 'Codigo de la AFP
                            StrSql = StrSql & 0 & "," 'Monto Capitalizacion Voluntaria
                            StrSql = StrSql & 0 & "," 'Monto Ahorro Voluntario
                            StrSql = StrSql & 0 & "," 'Nu8mero de Periodos de Cotizacion
                            'StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                            If arreglo(62) = 0 Then
                                StrSql = StrSql & arregloEstruc(62) & "," 'Codigo EX caja Regimen
                            Else
                                StrSql = StrSql & arreglo(62) & "," 'Codigo EX caja Regimen
                            End If
                            StrSql = StrSql & arreglo(63) & "," 'Tasa cotizacion EX cajas de Regimen
                            StrSql = StrSql & arreglo(64) & "," 'Renta Imponible IPS
                            StrSql = StrSql & arreglo(65) & "," 'Cotizacion Obligatoria INP
                            StrSql = StrSql & arreglo(66) & "," 'Renta Imponible Desahucio
                            StrSql = StrSql & arregloEstruc(67) & "," 'Codigo ex caja Regimen
                            StrSql = StrSql & arreglo(68) & "," 'Tasa Cotizacion Desahucio
                            StrSql = StrSql & arreglo(69) & "," 'Cotizacion Desahucio
                            StrSql = StrSql & arreglo(70) & "," 'Cotizacion Fonasa
                            StrSql = StrSql & arreglo(71) & "," 'Cotizacion Accidente de Trabajo
                            StrSql = StrSql & arreglo(72) & "," 'Bonificacion ley 15.386
                            StrSql = StrSql & arreglo(73) & "," 'Descuento por Cargas Familiares
                            StrSql = StrSql & arreglo(74) & "," 'Bonos Gobierno
                            StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                            StrSql = StrSql & FUN & ","  'FUN
                            StrSql = StrSql & arreglo(77) & "," 'Renta Imponible Isapre
                            StrSql = StrSql & arreglo(78) & "," 'Moneda del plan pactado con Isapre
                            StrSql = StrSql & arreglo(79) & "," 'Cotizacion Pactada
                            StrSql = StrSql & arreglo(80) & "," 'Cotizacion Obligatoria ISAPRE
                            StrSql = StrSql & arreglo(81) & "," 'Cotizacion adicional Voluntaria
                            StrSql = StrSql & arreglo(82) & "," 'Monto Garantia Explicito
                            StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                            StrSql = StrSql & arreglo(84) & "," 'Renta imponible CCAF
                            StrSql = StrSql & arreglo(85) & "," 'Creditos Personales CCAF
                            StrSql = StrSql & arreglo(86) & "," 'Descuento Dental
                            StrSql = StrSql & arreglo(87) & "," 'Descuento por Leasing
                            StrSql = StrSql & arreglo(88) & "," 'Descuentos por Seguro de Vida
                            StrSql = StrSql & arreglo(89) & "," 'Otros Descuentos CCAF
                            StrSql = StrSql & arreglo(90) & "," 'Cotiz CCAF de no Afiliados ISAPRE
                            StrSql = StrSql & arreglo(91) & "," 'Descuentos Cargas Familiares CCAF
                            StrSql = StrSql & arreglo(92) & "," 'Otros Descuentos CCAF 1
                            StrSql = StrSql & arreglo(93) & "," 'Otros Descuentos CCAF 2
                            StrSql = StrSql & arreglo(94) & "," 'Bonos Gobierno
                            StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                            StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                            StrSql = StrSql & arreglo(97) & "," 'Renta Imponible Mutual
                            StrSql = StrSql & arreglo(98) & "," 'Cotizacion Accidente del trabajo
                            StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(100)) & "," 'Renta total Imponible Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(101)) & "," 'Aporte Trabajador Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(102)) & "," 'Aporte Empleador Seguro de Cesantia
                            StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                            StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                            StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                            StrSql = StrSql & IIf(HuboError, -1, 0)
                            StrSql = StrSql & ")"
                            'Flog.writeline
                            'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            'Flog.writeline
                            'Flog.writeline
                            'Sumo el numero de linea
                            Num_linea = Num_linea + 1
                            CantEmplSinError = CantEmplSinError + 1
                            'Flog.writeline
                            'Flog.writeline Espacios(Tabulador * 2) & "SE GRABO EL EMPLEADO "
                            'Flog.writeline
                            
                            '----------------------------------------------------------------------------------------------------------
                            'Aperturas de lineas x ccaf
                            '----------------------------------------------------------------------------------------------------------
                            If AperCCAF Then
                            
                                'Verifico si ya se inserto una cabecera
                                If CodCabCCAF = 0 Then
                                    StrSql = "SELECT * FROM PMRegino WHERE PMreginonro = " & PMReginoNro
                                    OpenRecordset StrSql, rs_Aux
                                    
                                    If Not rs_Aux.EOF Then
                                        StrSql = "INSERT INTO PMRegino"
                                        StrSql = StrSql & "(bpronro,identificador,nomina"
                                        StrSql = StrSql & ",RUTPag,DVPag,TipoNom,CodForm"
                                        StrSql = StrSql & ",Periodo,CantReg,rol,mail)"
                                        StrSql = StrSql & "VALUES("
                                        StrSql = StrSql & "  " & NroProcesoBatch
                                        StrSql = StrSql & ",'" & rs_Aux!identificador & "'"
                                        StrSql = StrSql & ",'" & Mid(rs_Aux!nomina & " CCAF", 1, 50) & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!RUTPag & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!DVPag & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!TipoNom & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!CodForm & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!periodo & "'"
                                        StrSql = StrSql & ",0"
                                        StrSql = StrSql & ",'" & rs_Aux!rol & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!mail & "'"
                                        StrSql = StrSql & ")"
                                        objConn.Execute StrSql, , adExecuteNoRecords
        
                                        CodCabCCAF = getLastIdentity(objConn, "PMRegino")
                                        
                                        'Actualizo la cantidad de nominas
                                        CantEmpr = CantEmpr + 1
                                        
                                    End If
                                    
                                    rs_Aux.Close
                                End If
                                
                                'Verifico nuevamente para ver si se inserto el registro
                                If CodCabCCAF <> 0 Then
                                    
                                    Do While Not rs_CCAF.EOF
                                        
                                        NroLineaCCAF = NroLineaCCAF + 1
                                        
                                        StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                        StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                        StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                        StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                        StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                        StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                        StrSql = StrSql & ") VALUES ("
                                        StrSql = StrSql & NroProcesoBatch & ","
                                        StrSql = StrSql & CodCabCCAF & ","
                                        StrSql = StrSql & rs_Empleados!ternro & ","
                                        StrSql = StrSql & NroLineaCCAF & ","
                                        StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                        StrSql = StrSql & "Null,"
                                        StrSql = StrSql & "Null,"
                                        StrSql = StrSql & Empresa & ","
                                        StrSql = StrSql & "'" & RUT & "',"
                                        StrSql = StrSql & "'" & DV & "',"
                                        StrSql = StrSql & "'" & Apellido & "',"
                                        StrSql = StrSql & "'" & Apellido2 & "',"
                                        StrSql = StrSql & "'" & NombreEmp & "',"
                                        StrSql = StrSql & "'" & Sexo & "',"
                                        StrSql = StrSql & "'" & Nacionalidad & "',"
                                        'StrSql = StrSql & TipoPago & ","
                                        StrSql = StrSql & "1,"
                                        StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                        StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                        StrSql = StrSql & "0," 'Renta imponible
                                        StrSql = StrSql & "'SIP'," 'Regimen Previsional
                                        StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                        StrSql = StrSql & "0," 'Dias Trabajados
                                        'StrSql = StrSql & TipoLinea & "," 'Tipo de Linea
                                        StrSql = StrSql & "2," 'Tipo de Linea
                                        StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                                        StrSql = StrSql & IIf(arregloFecD(1) = vbNull, "Null", ConvFecha(arregloFecD(1))) & "," 'Fecha Desde para el movimiento
                                        StrSql = StrSql & IIf(arregloFecH(1) = vbNull, "Null", ConvFecha(arregloFecH(1))) & "," 'Fecha Hasta para el movimiento
                                        StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                        StrSql = StrSql & "0," 'Numero de cargas simples
                                        StrSql = StrSql & "0," 'Numero de cargas Maternales
                                        StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                        StrSql = StrSql & "0," 'ASignacion Familiar
                                        StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                        StrSql = StrSql & "0," 'Renta Carga Familiares
                                        StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                        
                                        StrSql = StrSql & "0," 'Codigo AFP
                                        
                                        StrSql = StrSql & "0," 'Renta Imponible AFP
                                        StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                        StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                        StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                        StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                        StrSql = StrSql & "0," 'Tasa Pactada
                                        StrSql = StrSql & "0," 'Aporte Indem
                                        StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                        StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                        StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                        StrSql = StrSql & "' '," 'Puesto de trabajo Pesado
                                        StrSql = StrSql & "0," 'Porcentaje Cotizacion Trabajo Pesado
                                        StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                                        StrSql = StrSql & "0," 'Inst Autor APVI
                                        StrSql = StrSql & "0," 'Numero de contrato APVI
                                        StrSql = StrSql & "0," 'Forma de Pago APVI
                                        StrSql = StrSql & "0," 'Cotizacion APVI
                                        StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                                        StrSql = StrSql & "0,"  'Codigo Institucion Autorizada APVC
                                        StrSql = StrSql & "0,"  'Numero de Contrato APVC
                                        StrSql = StrSql & "0,"  'Forma de Pago APVC
                                        StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                        StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                                        StrSql = StrSql & "'" & "0" & "'," 'RUT Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                        StrSql = StrSql & 0 & "," 'Codigo Movimiento Personal Afiliado Voluntario
                                        StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                        StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                        StrSql = StrSql & 0 & "," 'Codigo de la AFP
                                        StrSql = StrSql & 0 & "," 'Monto Capitalizacion Voluntaria
                                        StrSql = StrSql & 0 & "," 'Monto Ahorro Voluntario
                                        StrSql = StrSql & 0 & "," 'Nu8mero de Periodos de Cotizacion
                                        StrSql = StrSql & "0," 'Codigo Caja Regimen
                                        StrSql = StrSql & "0," 'Tasa cotizacion EX cajas de Regimen
                                        StrSql = StrSql & "0," 'Renta Imponible IPS
                                        StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                        StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                        StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                        StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                        StrSql = StrSql & "0," 'Cotizacion Desahucio
                                        StrSql = StrSql & "0," 'Cotizacion Fonasa
                                        StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                        StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                        StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                        StrSql = StrSql & "0," 'Bonos Gobierno
                                        StrSql = StrSql & "0," 'Codigo Institucion de Salud
                                        StrSql = StrSql & "'',"  'FUN
                                        StrSql = StrSql & "0," 'Renta Imponible Isapre
                                        StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                        StrSql = StrSql & "0," 'Cotizacion Pactada
                                        StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                        StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                        StrSql = StrSql & "0," 'Monto Garantia Explicito
                                        StrSql = StrSql & " " & Fix(rs_CCAF!Cant) & "," 'Codigo CCAF
                                        StrSql = StrSql & "0," 'Renta imponible CCAF
                                        
                                        'StrSql = StrSql & Aux85 & ","  'Creditos Personales CCAF
                                        If (UBound(Aux85_nroLinea) >= 1) And (UBound(Aux85_nroLinea) >= NroLineaCCAF) Then
                                            StrSql = StrSql & Aux85_nroLinea(NroLineaCCAF) & ","
                                        Else
                                            StrSql = StrSql & Aux85 & ","  'Creditos Personales CCAF
                                        End If
                                        
                                        StrSql = StrSql & "0," 'Descuento Dental
                                        
                                        'StrSql = StrSql & " " & Abs(rs_CCAF!Monto) & "," 'Descuento por Leasing
                                        'StrSql = StrSql & " " & Aux87 & "," 'Descuento por Leasing
                                       
                                        If (UBound(Aux87_nroLinea) >= 1) And (UBound(Aux87_nroLinea) >= NroLineaCCAF) Then
                                            StrSql = StrSql & Aux87_nroLinea(NroLineaCCAF) & ","
                                        Else
                                            StrSql = StrSql & Aux87 & ","  'Creditos Personales CCAF
                                        End If
                                        
                                        StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                        StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                        StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                        StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                        StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                        StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                        StrSql = StrSql & "0," 'Bonos Gobierno
                                        StrSql = StrSql & "' '," 'Codigo de Sucursal
                                        StrSql = StrSql & "0," 'Codigo Mutual
                                        StrSql = StrSql & "0," 'Renta Imponible Mutual
                                        StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                        StrSql = StrSql & "0," 'Sucursal Para Pago Mutual
                                        StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                        StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                        StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                        StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                        StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                        StrSql = StrSql & "'',"   'Centro Costo, sucursal, etc
                                        StrSql = StrSql & IIf(HuboError, -1, 0)
                                        StrSql = StrSql & ")"
                                        'Flog.writeline
                                        'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                                        objConn.Execute StrSql, , adExecuteNoRecords

                                        rs_CCAF.MoveNext
                                    Loop
                                End If
                            End If 'If AperCCAF Then
                            
                            aux = 2
                            Do While aux <= total_mov And (aux < 30)
                               'Inserto en rep_previred las lineas adicionales
                                StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                
                                StrSql = StrSql & ") VALUES ("
                                
                                StrSql = StrSql & NroProcesoBatch & ","
                                StrSql = StrSql & PMReginoNro & ","
                                StrSql = StrSql & rs_Empleados!ternro & ","
                                StrSql = StrSql & Num_linea & ","
                                StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & Empresa & ","
                                StrSql = StrSql & "'" & RUT & "',"
                                StrSql = StrSql & "'" & DV & "',"
                                StrSql = StrSql & "'" & Apellido & "',"
                                StrSql = StrSql & "'" & Apellido2 & "',"
                                StrSql = StrSql & "'" & NombreEmp & "',"
                                StrSql = StrSql & "'" & Sexo & "',"
                                StrSql = StrSql & "'" & Nacionalidad & "',"
                                StrSql = StrSql & TipoPago & ","
                                StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                StrSql = StrSql & "0," 'Renta imponible
                                StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                                StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                StrSql = StrSql & "0," 'Dias Trabajados
                                StrSql = StrSql & "01" & "," 'Tipo de Linea Adicional
                                StrSql = StrSql & arregloMov(aux) & "," 'Codigo Movimiento de personal
                                StrSql = StrSql & IIf(arregloFecD(aux) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & IIf(arregloFecH(aux) = vbNull, "Null", ConvFecha(arregloFecH(aux))) & "," 'Fecha Hasta para el movimiento
                                StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                StrSql = StrSql & "0," 'Numero de cargas simples
                                StrSql = StrSql & "0," 'Numero de cargas Maternales
                                StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                StrSql = StrSql & "0," 'ASignacion Familiar
                                StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                StrSql = StrSql & "0," 'Renta Carga Familiares
                                StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                If arreglo(26) = 0 Then
                                    StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                Else
                                    StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                                End If
                                'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                StrSql = StrSql & "0," 'Renta Imponible AFP
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                StrSql = StrSql & "0.00," 'Tasa Pactada
                                StrSql = StrSql & "0," 'Aporte Indem
                                StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                                StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                                StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                                StrSql = StrSql & "000," 'Inst Autor APVI
                                StrSql = StrSql & "0," 'Numero de contrato APVI
                                StrSql = StrSql & "0,"  'Forma de Pago APVI
                                StrSql = StrSql & "0," 'Cotizacion APVI
                                StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                                StrSql = StrSql & "0," 'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & "'0'," 'Numero de Contrato APVC
                                StrSql = StrSql & "0," 'Forma de Pago APVC
                                StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                                StrSql = StrSql & "'" & "0" & "'," 'RUT Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo de la AFP
                                StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                                StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                                StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                                'StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                                If arreglo(62) = 0 Then
                                    StrSql = StrSql & arregloEstruc(62) & "," 'Codigo EX caja Regimen
                                Else
                                    StrSql = StrSql & arreglo(62) & "," 'Codigo EX caja Regimen
                                End If
                                StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                                StrSql = StrSql & "0," 'Renta Imponible IPS
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Fonasa
                                StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                                StrSql = StrSql & "0,"  'FUN
                                StrSql = StrSql & "0," 'Renta Imponible Isapre
                                StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                StrSql = StrSql & "0," 'Cotizacion Pactada
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                                StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                                StrSql = StrSql & "0," 'Renta imponible CCAF
                                StrSql = StrSql & "0,"  'Creditos Personales CCAF
                                StrSql = StrSql & "0," 'Descuento Dental
                                StrSql = StrSql & "0," 'Descuento por Leasing
                                StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                                StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                                StrSql = StrSql & "0," 'Renta Imponible Mutual
                                StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                                StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                                StrSql = StrSql & IIf(HuboError, -1, 0)
                                StrSql = StrSql & ")"
                                'Flog.writeline
                                'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'Flog.writeline
                                'Flog.writeline
                                'Sumo el numero de linea
                                Num_linea = Num_linea + 1
                                'Flog.writeline
                                'Flog.writeline Espacios(Tabulador * 2) & "SE GRABO LINEA ADICIONAL POR MOVIMIENTO"
                                'Flog.writeline
            
                                aux = aux + 1
                            Loop
                            
                            'Aperturas de lineas de APVI
                            aux = 2
                            Do While (aux <= total_APVI) And (aux < 30)
                               'Inserto en rep_previred las lineas adicionales APVI
                                StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                
                                StrSql = StrSql & ") VALUES ("
                                
                                StrSql = StrSql & NroProcesoBatch & ","
                                StrSql = StrSql & PMReginoNro & ","
                                StrSql = StrSql & rs_Empleados!ternro & ","
                                StrSql = StrSql & Num_linea & ","
                                StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & Empresa & ","
                                StrSql = StrSql & "'" & RUT & "',"
                                StrSql = StrSql & "'" & DV & "',"
                                StrSql = StrSql & "'" & Apellido & "',"
                                StrSql = StrSql & "'" & Apellido2 & "',"
                                StrSql = StrSql & "'" & NombreEmp & "',"
                                StrSql = StrSql & "'" & Sexo & "',"
                                StrSql = StrSql & "'" & Nacionalidad & "',"
                                StrSql = StrSql & TipoPago & ","
                                StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                StrSql = StrSql & "0," 'Renta imponible
                                StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                                StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                StrSql = StrSql & "0," 'Dias Trabajados
                                StrSql = StrSql & "01" & "," 'Tipo de Linea
                                '06/10/2010 - Martin Ferraro - Codigo Movimiento de personal Fijo 0
                                'StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                                StrSql = StrSql & "0," 'Codigo Movimiento de personal
                                StrSql = StrSql & IIf(arregloFecD(aux) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & IIf(arregloFecH(aux) = vbNull, "Null", ConvFecha(arregloFecH(aux))) & "," 'Fecha Hasta para el movimiento
                                StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                StrSql = StrSql & "0," 'Numero de cargas simples
                                StrSql = StrSql & "0," 'Numero de cargas Maternales
                                StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                StrSql = StrSql & "0," 'ASignacion Familiar
                                StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                StrSql = StrSql & "0," 'Renta Carga Familiares
                                StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                If arreglo(26) = 0 Then
                                    StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                Else
                                    StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                                End If
                                StrSql = StrSql & "0," 'Renta Imponible AFP
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                StrSql = StrSql & "0.00," 'Tasa Pactada
                                StrSql = StrSql & "0," 'Aporte Indem
                                StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                                StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                                StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado

                                StrSql = StrSql & arregloAPVI(aux).Cod & "," 'Inst Autor APVI
                                StrSql = StrSql & arregloEstruc(41) & "," 'Numero de contrato APVI
                                StrSql = StrSql & arregloEstruc(42) & "," 'Forma de Pago APVI
                                StrSql = StrSql & arregloAPVI(aux).Cotiza & "," 'Cotizacion APVI
                                StrSql = StrSql & arregloAPVI(aux).Depositos & "," 'Cotizacion Depositos convenidos
                                
                                If total_APVC > 0 Then
                                    StrSql = StrSql & arregloAPVC(1).Cod & ","  'Codigo Institucion Autorizada APVC
                                    StrSql = StrSql & arregloEstruc(46) & ","  'Numero de Contrato APVC
                                    StrSql = StrSql & arregloEstruc(47) & ","  'Forma de Pago APVC
                                    StrSql = StrSql & arregloAPVC(1).Cotiza & "," 'Cotizacion Trabajador APVC
                                    StrSql = StrSql & arregloAPVC(1).Depositos & "," 'Cotizacion Empleador APVC
                                Else
                                    StrSql = StrSql & "0,"  'Codigo Institucion Autorizada APVC
                                    StrSql = StrSql & "0,"  'Numero de Contrato APVC
                                    StrSql = StrSql & "0,"  'Forma de Pago APVC
                                    StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                    StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                                End If
                                
                                StrSql = StrSql & "'" & "0" & "'," 'RUT Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo de la AFP
                                StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                                StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                                StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                                'StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                                If arreglo(62) = 0 Then
                                    StrSql = StrSql & arregloEstruc(62) & "," 'Codigo EX caja Regimen
                                Else
                                    StrSql = StrSql & arreglo(62) & "," 'Codigo EX caja Regimen
                                End If
                                StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                                StrSql = StrSql & "0," 'Renta Imponible IPS
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Fonasa
                                StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                                StrSql = StrSql & "0,"  'FUN
                                StrSql = StrSql & "0," 'Renta Imponible Isapre
                                StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                StrSql = StrSql & "0," 'Cotizacion Pactada
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                                StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                                StrSql = StrSql & "0," 'Renta imponible CCAF
                                StrSql = StrSql & "0,"  'Creditos Personales CCAF
                                StrSql = StrSql & "0," 'Descuento Dental
                                StrSql = StrSql & "0," 'Descuento por Leasing
                                StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                                StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                                StrSql = StrSql & "0," 'Renta Imponible Mutual
                                StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                                StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                                StrSql = StrSql & IIf(HuboError, -1, 0)
                                
                                StrSql = StrSql & ")"
                                'Flog.writeline
                                'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'Flog.writeline
                                'Flog.writeline
                                'Sumo el numero de linea
                                Num_linea = Num_linea + 1
                                'Flog.writeline
                                'Flog.writeline Espacios(Tabulador * 2) & "SE GRABO LINEA ADICIONAL DE APVI"
                                'Flog.writeline
            
                                aux = aux + 1
                            Loop
                            
                            'Aperturas de lineas de APVC
                            aux = 2
                            Do While aux <= total_APVC And (aux < 30)
                               'Inserto en rep_previred las lineas adicionales
                                StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                
                                StrSql = StrSql & ") VALUES ("
                                
                                StrSql = StrSql & NroProcesoBatch & ","
                                StrSql = StrSql & PMReginoNro & ","
                                StrSql = StrSql & rs_Empleados!ternro & ","
                                StrSql = StrSql & Num_linea & ","
                                StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & Empresa & ","
                                StrSql = StrSql & "'" & RUT & "',"
                                StrSql = StrSql & "'" & DV & "',"
                                StrSql = StrSql & "'" & Apellido & "',"
                                StrSql = StrSql & "'" & Apellido2 & "',"
                                StrSql = StrSql & "'" & NombreEmp & "',"
                                StrSql = StrSql & "'" & Sexo & "',"
                                StrSql = StrSql & "'" & Nacionalidad & "',"
                                StrSql = StrSql & TipoPago & ","
                                StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                StrSql = StrSql & "0," 'Renta imponible
                                StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                                StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                StrSql = StrSql & "0," 'Dias Trabajados
                                StrSql = StrSql & "01" & "," 'Tipo de Linea
                                '06/10/2010 - Martin Ferraro - Codigo Movimiento de personal Fijo 0
                                'StrSql = StrSql & arregloMov(aux) & "," 'Codigo Movimiento de personal
                                StrSql = StrSql & "0," 'Codigo Movimiento de personal
                                StrSql = StrSql & IIf(arregloFecD(aux) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & IIf(arregloFecH(aux) = vbNull, "Null", ConvFecha(arregloFecH(aux))) & "," 'Fecha Hasta para el movimiento
                                StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                StrSql = StrSql & "0," 'Numero de cargas simples
                                StrSql = StrSql & "0," 'Numero de cargas Maternales
                                StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                StrSql = StrSql & "0," 'ASignacion Familiar
                                StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                StrSql = StrSql & "0," 'Renta Carga Familiares
                                StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                If arreglo(26) = 0 Then
                                   StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                Else
                                   StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                                End If
                                StrSql = StrSql & "0," 'Renta Imponible AFP
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                StrSql = StrSql & "0.00," 'Tasa Pactada
                                StrSql = StrSql & "0," 'Aporte Indem
                                StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                                StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                                StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                                
                                If total_APVI > 0 Then
                                    StrSql = StrSql & arregloAPVI(1).Cod & "," 'Inst Autor APVI
                                    StrSql = StrSql & arregloEstruc(41) & "," 'Numero de contrato APVI
                                    StrSql = StrSql & arregloEstruc(42) & "," 'Forma de Pago APVI
                                    StrSql = StrSql & arregloAPVI(1).Cotiza & "," 'Cotizacion APVI
                                    StrSql = StrSql & arregloAPVI(1).Depositos & "," 'Cotizacion Depositos convenidos
                                Else
                                    StrSql = StrSql & "0," 'Inst Autor APVI
                                    StrSql = StrSql & "0," 'Numero de contrato APVI
                                    StrSql = StrSql & "0," 'Forma de Pago APVI
                                    StrSql = StrSql & "0," 'Cotizacion APVI
                                    StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                                End If
                                
                                StrSql = StrSql & arregloAPVC(aux).Cod & ","  'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & arregloEstruc(46) & ","  'Numero de Contrato APVC
                                StrSql = StrSql & arregloEstruc(47) & ","  'Forma de Pago APVC
                                StrSql = StrSql & arregloAPVC(aux).Cotiza & "," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & arregloAPVC(aux).Depositos & "," 'Cotizacion Empleador APVC
                                
                                StrSql = StrSql & "'" & "0" & "'," 'RUT Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo de la AFP
                                StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                                StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                                StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                                'StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                                If arreglo(62) = 0 Then
                                    StrSql = StrSql & arregloEstruc(62) & "," 'Codigo EX caja Regimen
                                Else
                                    StrSql = StrSql & arreglo(62) & "," 'Codigo EX caja Regimen
                                End If
                                StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                                StrSql = StrSql & "0," 'Renta Imponible IPS
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Fonasa
                                StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                                StrSql = StrSql & "0,"  'FUN
                                StrSql = StrSql & "0," 'Renta Imponible Isapre
                                StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                StrSql = StrSql & "0," 'Cotizacion Pactada
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                                StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                                StrSql = StrSql & "0," 'Renta imponible CCAF
                                StrSql = StrSql & "0," 'Creditos Personales CCAF
                                StrSql = StrSql & "0," 'Descuento Dental
                                StrSql = StrSql & "0," 'Descuento por Leasing
                                StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                                StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                                StrSql = StrSql & "0," 'Renta Imponible Mutual
                                StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                                StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                                StrSql = StrSql & IIf(HuboError, -1, 0)
                                
                                StrSql = StrSql & ")"
                                'Flog.writeline
                                'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'Flog.writeline
                                'Flog.writeline
                                'Sumo el numero de linea
                                Num_linea = Num_linea + 1
                                'Flog.writeline
                                'Flog.writeline Espacios(Tabulador * 2) & "SE GRABO LINEA ADICIONAL POR APVC"
                                'Flog.writeline
            
                                aux = aux + 1
                            Loop
                            
                        Case 2: 'Gratificaciones
                            
                            'Busco las fechas desde/hasta de los periodos de reliquidacion
                            
                            FechaDesdePeri = Fechadesde
                            FechaHastaPeri = Fechahasta
                            
                            StrSql = "SELECT " & TOP(1) & " pliqdesde FROM impuni_peri "
                            StrSql = StrSql & " inner join periodo on periodo.pliqnro = impuni_peri.pliqnro "
                            StrSql = StrSql & " Where pronro = " & rs_Empleados!pronro
                            StrSql = StrSql & " order by pliqdesde"
                            OpenRecordset StrSql, rs_impuniperi
                            If Not rs_impuniperi.EOF Then
                                    FechaDesdePeri = rs_impuniperi!pliqdesde
                            End If
                            
                            StrSql = "SELECT " & TOP(1) & " pliqhasta FROM impuni_peri "
                            StrSql = StrSql & " inner join periodo on periodo.pliqnro = impuni_peri.pliqnro "
                            StrSql = StrSql & " Where pronro = " & rs_Empleados!pronro
                            StrSql = StrSql & " order by pliqhasta desc"
                            OpenRecordset StrSql, rs_impuniperi
                            If Not rs_impuniperi.EOF Then
                                    FechaHastaPeri = rs_impuniperi!pliqhasta
                            End If
                            
                            
                            '//sebastian stremel 14/09/2012 si el empleado no tiene la empresa en la fecha del impuni_peri no lo muestro
                            
                            '//hasta aca
                            
                            'Inserto en rep_previred
                            
                            StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                            StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                            StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                            StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                            StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                            StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                            
                            StrSql = StrSql & ") VALUES ("
                            
                            StrSql = StrSql & NroProcesoBatch & ","
                            StrSql = StrSql & PMReginoNro & ","
                            StrSql = StrSql & rs_Empleados!ternro & ","
                            StrSql = StrSql & Num_linea & ","
                            StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & Empresa & ","
                            StrSql = StrSql & "'" & RUT & "',"
                            StrSql = StrSql & "'" & DV & "',"
                            StrSql = StrSql & "'" & Apellido & "',"
                            StrSql = StrSql & "'" & Apellido2 & "',"
                            StrSql = StrSql & "'" & NombreEmp & "',"
                            StrSql = StrSql & "'" & Sexo & "',"
                            StrSql = StrSql & "'" & Nacionalidad & "',"
                            StrSql = StrSql & TipoPago & ","
                            StrSql = StrSql & "'" & Month(FechaDesdePeri) & Year(FechaDesdePeri) & "',"
                            StrSql = StrSql & "'" & Month(FechaHastaPeri) & Year(FechaHastaPeri) & "',"
                            StrSql = StrSql & arreglo(10) & "," 'Renta imponible
                            StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                            StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                            StrSql = StrSql & "30," 'Dias Trabajados
                            StrSql = StrSql & "'00'," 'Tipo de Linea
                            StrSql = StrSql & "0," 'Codigo Movimiento de personal
                            StrSql = StrSql & "null," 'Fecha Desde para el movimiento
                            StrSql = StrSql & "null," 'Fecha Hasta para el movimiento
                            StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                            StrSql = StrSql & "0," 'Numero de cargas simples
                            StrSql = StrSql & "0," 'Numero de cargas Maternales
                            StrSql = StrSql & "0," 'Numero de cargas Invalidas
                            StrSql = StrSql & "0," 'ASignacion Familiar
                            StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                            StrSql = StrSql & "0," 'Renta Carga Familiares
                            StrSql = StrSql & "'N'," 'Solicitud Subsidio Trabajador Joven
                            'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                            If arreglo(26) = 0 Then
                                StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                            Else
                                StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                            End If
                            StrSql = StrSql & arreglo(27) & "," 'Renta Imponible AFP
                            StrSql = StrSql & arreglo(28) & "," 'Cotizacion Obligatoria AFP
                            StrSql = StrSql & arreglo(29) & "," 'Aporte Seguro Invalidez y Sobervivencia
                            StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                            StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                            StrSql = StrSql & "0.00," 'Tasa Pactada
                            StrSql = StrSql & "0," 'Aporte Indem
                            StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                            StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                            StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                            StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                            StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                            StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                            StrSql = StrSql & "000," 'Inst Autor APVI
                            StrSql = StrSql & "'" & "" & "'," 'Numero de contrato APVI
                            StrSql = StrSql & "0,"  'Forma de Pago APVI
                            StrSql = StrSql & "0," 'Cotizacion APVI
                            StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                            StrSql = StrSql & "000," 'Codigo Institucion Autorizada APVC
                            StrSql = StrSql & "'" & "" & "'," 'Numero de Contrato APVC
                            StrSql = StrSql & "0," 'Forma de Pago APVC
                            StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                            StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                            StrSql = StrSql & "'0'," 'RUT Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                            StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                            StrSql = StrSql & "0," 'Codigo de la AFP
                            StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                            StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                            StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                            'StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                            If arreglo(62) = 0 Then
                                StrSql = StrSql & arregloEstruc(62) & "," 'Codigo EX caja Regimen
                            Else
                                StrSql = StrSql & arreglo(62) & "," 'Codigo EX caja Regimen
                            End If
                            StrSql = StrSql & arreglo(63) & "," 'Tasa cotizacion EX cajas de Regimen
                            StrSql = StrSql & arreglo(64) & "," 'Renta Imponible IPS
                            StrSql = StrSql & arreglo(65) & "," 'Cotizacion Obligatoria INP
                            StrSql = StrSql & arreglo(66) & "," 'Renta Imponible Desahucio
                            StrSql = StrSql & arreglo(67) & "," 'Codigo ex caja Regimen
                            StrSql = StrSql & arreglo(68) & "," 'Tasa Cotizacion Desahucio
                            StrSql = StrSql & arreglo(69) & "," 'Cotizacion Desahucio
                            StrSql = StrSql & arreglo(70) & "," 'Cotizacion Fonasa
                            StrSql = StrSql & arreglo(71) & "," 'Cotizacion Accidente de Trabajo
                            StrSql = StrSql & "0," 'Bonificacion ley 15.386
                            StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                            StrSql = StrSql & "0," 'Bonos Gobierno
                            StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                            StrSql = StrSql & "0,"  'FUN
                            StrSql = StrSql & arreglo(77) & "," 'Renta Imponible Isapre
                            StrSql = StrSql & "1," 'Moneda del plan pactado con Isapre
                            StrSql = StrSql & "0," 'Cotizacion Pactada
                            StrSql = StrSql & arreglo(80) & "," 'Cotizacion Obligatoria ISAPRE
                            StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                            StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                            StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                            'StrSql = StrSql & "0," 'Renta imponible CCAF
                            StrSql = StrSql & arreglo(84) & "," 'Renta imponible CCAF
                            StrSql = StrSql & "0," 'Creditos Personales CCAF
                            StrSql = StrSql & "0," 'Descuento Dental
                            StrSql = StrSql & "0," 'Descuento por Leasing
                            StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                            StrSql = StrSql & "0," 'Otros Descuentos CCAF
                            StrSql = StrSql & arreglo(90) & "," 'Cotiz CCAF de no Afiliados ISAPRE
                            StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                            StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                            StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                            StrSql = StrSql & "0," 'Bonos Gobierno
                            StrSql = StrSql & "0," 'Codigo de Sucursal
                            StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                            StrSql = StrSql & arreglo(97) & "," 'Renta Imponible Mutual
                            StrSql = StrSql & arreglo(98) & "," 'Cotizacion Accidente del trabajo
                            StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(100)) & "," 'Renta total Imponible Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(101)) & "," 'Aporte Trabajador Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(102)) & "," 'Aporte Empleador Seguro de Cesantia
                            StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                            StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                            StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                            StrSql = StrSql & IIf(HuboError, -1, 0)
                            
                            StrSql = StrSql & ")"
                            'Flog.writeline
                            'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            'Flog.writeline
                            'Flog.writeline
                            'Sumo el numero de linea
                            Num_linea = Num_linea + 1
                            'Flog.writeline
                            'Flog.writeline Espacios(Tabulador * 2) & "SE GRABO LINEA DE GRATIFICACION"
                            'Flog.writeline
                            
                            
                            '----------------------------------------------------------------------------------------------------------
                            'Aperturas de lineas x ccaf
                            '----------------------------------------------------------------------------------------------------------
                            If AperCCAF Then
                            
                                'Verifico si ya se inserto una cabecera
                                If CodCabCCAF = 0 Then
                                    StrSql = "SELECT * FROM PMRegino WHERE PMreginonro = " & PMReginoNro
                                    OpenRecordset StrSql, rs_Aux
                                    
                                    If Not rs_Aux.EOF Then
                                        StrSql = "INSERT INTO PMRegino"
                                        StrSql = StrSql & "(bpronro,identificador,nomina"
                                        StrSql = StrSql & ",RUTPag,DVPag,TipoNom,CodForm"
                                        StrSql = StrSql & ",Periodo,CantReg,rol,mail)"
                                        StrSql = StrSql & "VALUES("
                                        StrSql = StrSql & "  " & NroProcesoBatch
                                        StrSql = StrSql & ",'" & rs_Aux!identificador & "'"
                                        StrSql = StrSql & ",'" & Mid(rs_Aux!nomina & " CCAF", 1, 50) & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!RUTPag & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!DVPag & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!TipoNom & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!CodForm & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!periodo & "'"
                                        StrSql = StrSql & ",0"
                                        StrSql = StrSql & ",'" & rs_Aux!rol & "'"
                                        StrSql = StrSql & ",'" & rs_Aux!mail & "'"
                                        StrSql = StrSql & ")"
                                        objConn.Execute StrSql, , adExecuteNoRecords
        
                                        CodCabCCAF = getLastIdentity(objConn, "PMRegino")
                                        
                                        'Actualizo la cantidad de nominas
                                        CantEmpr = CantEmpr + 1
                                        
                                    End If
                                    
                                    rs_Aux.Close
                                End If
                                
                                'Verifico nuevamente para ver si se inserto el registro
                                If CodCabCCAF <> 0 Then
                                    
                                    Do While Not rs_CCAF.EOF
                                        
                                        NroLineaCCAF = NroLineaCCAF + 1
                                        
                                        StrSql = "INSERT INTO PMprevired (bpronro, PMreginonro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                        StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                        StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                        StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                        StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                        StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                        StrSql = StrSql & ") VALUES ("
                                        StrSql = StrSql & NroProcesoBatch & ","
                                        StrSql = StrSql & CodCabCCAF & ","
                                        StrSql = StrSql & rs_Empleados!ternro & ","
                                        StrSql = StrSql & NroLineaCCAF & ","
                                        StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                        StrSql = StrSql & "Null,"
                                        StrSql = StrSql & "Null,"
                                        StrSql = StrSql & Empresa & ","
                                        StrSql = StrSql & "'" & RUT & "',"
                                        StrSql = StrSql & "'" & DV & "',"
                                        StrSql = StrSql & "'" & Apellido & "',"
                                        StrSql = StrSql & "'" & Apellido2 & "',"
                                        StrSql = StrSql & "'" & NombreEmp & "',"
                                        StrSql = StrSql & "'" & Sexo & "',"
                                        StrSql = StrSql & "'" & Nacionalidad & "',"
                                        'StrSql = StrSql & TipoPago & ","
                                        StrSql = StrSql & "1,"
                                        StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                        StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                        StrSql = StrSql & "0," 'Renta imponible
                                        StrSql = StrSql & "'SIP'," 'Regimen Previsional
                                        StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                        StrSql = StrSql & "0," 'Dias Trabajados
                                        'StrSql = StrSql & TipoLinea & "," 'Tipo de Linea
                                        StrSql = StrSql & "2," 'Tipo de Linea
                                        StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                                        StrSql = StrSql & IIf(arregloFecD(1) = vbNull, "Null", ConvFecha(arregloFecD(1))) & "," 'Fecha Desde para el movimiento
                                        StrSql = StrSql & IIf(arregloFecH(1) = vbNull, "Null", ConvFecha(arregloFecH(1))) & "," 'Fecha Hasta para el movimiento
                                        StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                        StrSql = StrSql & "0," 'Numero de cargas simples
                                        StrSql = StrSql & "0," 'Numero de cargas Maternales
                                        StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                        StrSql = StrSql & "0," 'ASignacion Familiar
                                        StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                        StrSql = StrSql & "0," 'Renta Carga Familiares
                                        StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                        
                                        StrSql = StrSql & "0," 'Codigo AFP
                                        
                                        StrSql = StrSql & "0," 'Renta Imponible AFP
                                        StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                        StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                        StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                        StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                        StrSql = StrSql & "0," 'Tasa Pactada
                                        StrSql = StrSql & "0," 'Aporte Indem
                                        StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                        StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                        StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                        StrSql = StrSql & "' '," 'Puesto de trabajo Pesado
                                        StrSql = StrSql & "0," 'Porcentaje Cotizacion Trabajo Pesado
                                        StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                                        StrSql = StrSql & "0," 'Inst Autor APVI
                                        StrSql = StrSql & "0," 'Numero de contrato APVI
                                        StrSql = StrSql & "0," 'Forma de Pago APVI
                                        StrSql = StrSql & "0," 'Cotizacion APVI
                                        StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                                        StrSql = StrSql & "0,"  'Codigo Institucion Autorizada APVC
                                        StrSql = StrSql & "0,"  'Numero de Contrato APVC
                                        StrSql = StrSql & "0,"  'Forma de Pago APVC
                                        StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                        StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                                        StrSql = StrSql & "'" & "0" & "'," 'RUT Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                        StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                        StrSql = StrSql & 0 & "," 'Codigo Movimiento Personal Afiliado Voluntario
                                        StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                        StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                        StrSql = StrSql & 0 & "," 'Codigo de la AFP
                                        StrSql = StrSql & 0 & "," 'Monto Capitalizacion Voluntaria
                                        StrSql = StrSql & 0 & "," 'Monto Ahorro Voluntario
                                        StrSql = StrSql & 0 & "," 'Nu8mero de Periodos de Cotizacion
                                        StrSql = StrSql & "0," 'Codigo Caja Regimen
                                        StrSql = StrSql & "0," 'Tasa cotizacion EX cajas de Regimen
                                        StrSql = StrSql & "0," 'Renta Imponible IPS
                                        StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                        StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                        StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                        StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                        StrSql = StrSql & "0," 'Cotizacion Desahucio
                                        StrSql = StrSql & "0," 'Cotizacion Fonasa
                                        StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                        StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                        StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                        StrSql = StrSql & "0," 'Bonos Gobierno
                                        StrSql = StrSql & "0," 'Codigo Institucion de Salud
                                        StrSql = StrSql & "'',"  'FUN
                                        StrSql = StrSql & "0," 'Renta Imponible Isapre
                                        StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                        StrSql = StrSql & "0," 'Cotizacion Pactada
                                        StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                        StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                        StrSql = StrSql & "0," 'Monto Garantia Explicito
                                        StrSql = StrSql & " " & Fix(rs_CCAF!Cant) & "," 'Codigo CCAF
                                        StrSql = StrSql & "0," 'Renta imponible CCAF
                                        
                                        'StrSql = StrSql & Aux85 & "," 'Creditos Personales CCAF
                                        
                                        'StrSql = StrSql & Aux85 & ","  'Creditos Personales CCAF
                                        If (UBound(Aux85_nroLinea) >= 1) And (UBound(Aux85_nroLinea) >= NroLineaCCAF) Then
                                            StrSql = StrSql & Aux85_nroLinea(NroLineaCCAF) & ","
                                        Else
                                            StrSql = StrSql & Aux85 & ","  'Creditos Personales CCAF
                                        End If
                                        
                                        StrSql = StrSql & "0," 'Descuento Dental
                                        
                                        'StrSql = StrSql & " " & Aux87 & "," 'Descuento por Leasing
                                        'StrSql = StrSql & Aux85 & ","  'Creditos Personales CCAF
                                        If (UBound(Aux87_nroLinea) >= 1) And (UBound(Aux87_nroLinea) >= NroLineaCCAF) Then
                                            StrSql = StrSql & Aux87_nroLinea(NroLineaCCAF) & ","
                                        Else
                                            StrSql = StrSql & Aux87 & ","  'Creditos Personales CCAF
                                        End If
                                        
                                        
                                        StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                        StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                        StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                        StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                        StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                        StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                        StrSql = StrSql & "0," 'Bonos Gobierno
                                        StrSql = StrSql & "' '," 'Codigo de Sucursal
                                        StrSql = StrSql & "0," 'Codigo Mutual
                                        StrSql = StrSql & "0," 'Renta Imponible Mutual
                                        StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                        StrSql = StrSql & "0," 'Sucursal Para Pago Mutual
                                        StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                        StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                        StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                        StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                        StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                        StrSql = StrSql & "'',"   'Centro Costo, sucursal, etc
                                        StrSql = StrSql & IIf(HuboError, -1, 0)
                                        StrSql = StrSql & ")"
                                        'Flog.writeline
                                        'Flog.writeline Espacios(Tabulador * 2) & "Insertando : " & StrSql
                                        objConn.Execute StrSql, , adExecuteNoRecords

                                        rs_CCAF.MoveNext
                                    Loop
                                End If
                            End If 'If AperCCAF Then
                            
                            
                    End Select
                    

                    
                'Else
                If HuboError Then
                    'Sumo 1 A la cantidad de errores
                    CantEmplError = CantEmplError + 1
                    'Flog.writeline
                    'Flog.writeline Espacios(Tabulador * 2) & "SE DETECTARON ERRORES EN EL EMPLEADO "
                    'Flog.writeline
                    'Errores = True
                End If
                    
                'Actualizo el progreso
                Progreso = Progreso + IncPorc
                TiempoAcumulado = GetTickCount
                
                'If Errores = False Then
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                    "' WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                'Else
                '    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                '    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                '    "',bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
                '    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                'End If
                
                ' ----------------------------------------------------------------
                
                End If
                
                
                
                'Paso al siguiente Empleado
SgtEmpl:        rs_Empleados.MoveNext

                'MyCommitTrans
    Loop
    
    
    'Si realizo apertura de CCAF entonces imprimo el pie
    'If AperCCAF Then
    If CodCabCCAF <> 0 Then
        StrSql = "INSERT INTO PMRegfno"
        StrSql = StrSql & " (bpronro,PMreginonro,descripcion)"
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & "  " & NroProcesoBatch
        StrSql = StrSql & " ," & CodCabCCAF
        StrSql = StrSql & " ,'" & Mid(DescPie & " CCAF", 1, 555) & "'"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Actualizo la cantidad de reg de la cabecera CCAF
        StrSql = "UPDATE PMRegino"
        StrSql = StrSql & " SET CantReg = " & NroLineaCCAF
        StrSql = StrSql & " WHERE PMreginonro = " & CodCabCCAF
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    
    'Actualizo o borro la cantidad de reg de la cabecera
    If Num_linea > 1 Then
        StrSql = "UPDATE PMRegino"
        StrSql = StrSql & " SET CantReg = " & Num_linea - 1
        StrSql = StrSql & " WHERE PMreginonro = " & PMReginoNro
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "DELETE PMRegino"
        StrSql = StrSql & " WHERE PMreginonro = " & PMReginoNro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If CantEmpr > 0 Then CantEmpr = CantEmpr - 1
    End If
    
    
End If 'If Not HuboError

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_CantEmpleados.State = adStateOpen Then rs_CantEmpleados.Close
If rs_Acu_liq.State = adStateOpen Then rs_Acu_liq.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Conceptos.State = adStateOpen Then rs_Conceptos.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Rut.State = adStateOpen Then rs_Rut.Close
If rs_Estr_cod.State = adStateOpen Then rs_Estr_cod.Close
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close
If rs_Nacionalidad.State = adStateOpen Then rs_Nacionalidad.Close
If rs_CCAF.State = adStateOpen Then rs_CCAF.Close


Set rs_Empleados = Nothing
Set rs_CantEmpleados = Nothing
Set rs_Acu_liq = Nothing
Set rs_Confrep = Nothing
Set rs_Conceptos = Nothing
Set rs_Detliq = Nothing
Set rs_Tercero = Nothing
Set rs_Estructura = Nothing
Set rs_Rut = Nothing
Set rs_Estr_cod = Nothing
Set rs_Fases = Nothing
Set rs_Aux = Nothing
Set rs_Nacionalidad = Nothing
Set rs_CCAF = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en Previred"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    'MyRollbackTrans
    'MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
    
    Errores = True
    Flog.writeline " Error: " & Err.Description

End Sub


