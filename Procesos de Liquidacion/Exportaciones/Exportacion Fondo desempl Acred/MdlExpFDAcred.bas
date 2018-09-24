Attribute VB_Name = "MdlExpFDAcred"
Option Explicit

Global Const Version = "1.00" 'Exportacion fondo desempleo banco nacion - Cuenta Legajo
Global Const FechaModificacion = "08/10/2007"
Global Const UltimaModificacion = " " 'Martin Ferraro - Version Inicial


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Exportacion de Fondo desempleo Cuenta Legajo.
' Autor      : Martin Ferraro
' Fecha      : 27/09/2007
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
    
    Nombre_Arch = PathFLog & "Exp. FD CuentaLegajo" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 202 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ExportLegFondoDesempl(NroProcesoBatch, bprcparam)
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


Public Sub ExportLegFondoDesempl(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Exportacion de Cuenta legajo de fondo de desempleo de banco nacion
' Autor      : Martin Ferraro
' Fecha      : 08/10/2007
' --------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim LegDesde As Long
Dim LegHasta As Long
Dim ParEstado As Integer
Dim Empnro As Long
Dim PliqNro As Long
Dim ListaProc As String
Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long
Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long
Dim FecEstr As Date

'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------
Dim ArrPar
Dim ternro As Long
Dim Directorio As String
Dim Separador As String
Dim SeparadorDecimal As String
Dim DescripcionModelo As String
Dim Archivo As String
Dim fExport
Dim carpeta
Dim Aux_Linea As String
Dim casa As String
Dim FormaPagoNro As Long
Dim Filler As String
Dim Acum As Boolean
Dim ConcAcumCod As String
Dim ConcAcumNro As Long
Dim concnro As Long
Dim Monto As Double
Dim MontoStr As String
Dim CuentaNro As String
Dim Transacion As String

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset

'Inicio codigo ejecutable
On Error GoTo E_LegFondoDesemple


'-------------------------------------------------------------------------------------------------
'Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, "@")
        If UBound(ArrPar) = 12 Then
            
            LegDesde = IIf(ArrPar(0) = "", 0, CLng(ArrPar(0)))
            LegHasta = IIf(ArrPar(1) = "", 0, CLng(ArrPar(1)))
            Flog.writeline Espacios(Tabulador * 0) & "Empleados Desde " & LegDesde & " Hasta " & LegHasta
            
            ParEstado = CInt(ArrPar(2))
            Select Case ParEstado
                Case -1:
                    Flog.writeline Espacios(Tabulador * 0) & "Activos"
                Case 0:
                    Flog.writeline Espacios(Tabulador * 0) & "Inactivos"
                Case 1:
                    Flog.writeline Espacios(Tabulador * 0) & "Ambos"
            End Select
            
            Empnro = CLng(ArrPar(3))
            Flog.writeline Espacios(Tabulador * 0) & "Estructura Empresa = " & Empnro

            PliqNro = CLng(ArrPar(4))
            Flog.writeline Espacios(Tabulador * 0) & "Periodo = " & PliqNro
            
            ListaProc = ArrPar(5)
            ListaProc = Replace(ListaProc, "*", ",")
            Flog.writeline Espacios(Tabulador * 0) & "Procesos = " & ListaProc
            
            Tenro1 = CLng(ArrPar(6))
            Flog.writeline Espacios(Tabulador * 0) & "Nivel TE1 = " & Tenro1
            
            Tenro2 = CLng(ArrPar(7))
            Flog.writeline Espacios(Tabulador * 0) & "Nivel TE2 = " & Tenro2
            
            Tenro3 = CLng(ArrPar(8))
            Flog.writeline Espacios(Tabulador * 0) & "Nivel TE3 = " & Tenro3
            
            Estrnro1 = CLng(ArrPar(9))
            Flog.writeline Espacios(Tabulador * 0) & "Estruct1 = " & Estrnro1
            
            Estrnro2 = CLng(ArrPar(10))
            Flog.writeline Espacios(Tabulador * 0) & "Estruct2 = " & Estrnro2
            
            Estrnro3 = CLng(ArrPar(11))
            Flog.writeline Espacios(Tabulador * 0) & "Estruct3 = " & Estrnro3
            
            FecEstr = CDate(ArrPar(12))
            Flog.writeline Espacios(Tabulador * 0) & "Fecha Estruct = " & FecEstr
            
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Numero de parametros erroneo."
            Exit Sub
        End If
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    Exit Sub
End If
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------

casa = ""
FormaPagoNro = 0
Filler = ""
Acum = True
ConcAcumCod = ""
ConcAcumNro = 0


Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte."
StrSql = "SELECT * FROM confrep WHERE repnro = 216 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró la configuración del Reporte."
    HuboError = True
    Exit Sub
Else
    Do While Not rs_Consult.EOF
                
        Flog.writeline Espacios(Tabulador * 1) & "Columna " & rs_Consult!confnrocol & " " & rs_Consult!confetiq & " VNUM = " & rs_Consult!confval & " VALF = " & rs_Consult!confval2
        
        Select Case rs_Consult!confnrocol
            Case 1:
                If EsNulo(rs_Consult!confval2) Then
                    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Falta configurar el valor de CASA en el campo alfanumerico de la columna 1."
                    HuboError = True
                    Exit Sub
                Else
                    casa = rs_Consult!confval2
                End If
            Case 2:
                If EsNulo(rs_Consult!confval) Then
                    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Falta configurar el valor de Forma de Pago en el campo numerico de la columna 2."
                    HuboError = True
                    Exit Sub
                Else
                    FormaPagoNro = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
                End If
            Case 3:
                If Not EsNulo(rs_Consult!confval2) Then
                    Filler = rs_Consult!confval2
                End If
            Case 4:
                If EsNulo(rs_Consult!confval2) Then
                    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Falta configurar el valor de Acumulador/Concepto en el campo alfanumerico de la columna 4."
                    HuboError = True
                    Exit Sub
                Else
                    Select Case UCase(rs_Consult!conftipo)
                        Case "AC":
                            Acum = True
                        Case "CO":
                            Acum = False
                        Case Else:
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El Tipo de la columna 4 debe ser AC o CO."
                            HuboError = True
                            Exit Sub
                    End Select
                    
                    ConcAcumCod = rs_Consult!confval2
                    
                End If
            Case 5:
                If Not EsNulo(rs_Consult!confval2) Then
                    Transacion = rs_Consult!confval2
                End If
        End Select
        
        rs_Consult.MoveNext
    Loop
End If
rs_Consult.Close
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Validaciones campos obligatorios confrep
'-------------------------------------------------------------------------------------------------
If Len(casa) = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Falta configurar el valor de CASA en el campo alfanumerico de la columna 1."
    HuboError = True
    Exit Sub
End If

If FormaPagoNro = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Falta configurar el valor de Forma de Pago en el campo numerico de la columna 2."
    HuboError = True
    Exit Sub
End If
    
ConcAcumCod = Trim(ConcAcumCod)
If Len(ConcAcumCod) = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Falta configurar el valor de Acumulador/Concepto en el campo alfanumerico de la columna 4."
    HuboError = True
    Exit Sub
End If
    
    
'-------------------------------------------------------------------------------------------------
'Busco el codigo interno del concepto
'-------------------------------------------------------------------------------------------------
If Not Acum Then
    StrSql = "SELECT concnro, conccod, concabr"
    StrSql = StrSql & " FROM concepto"
    StrSql = StrSql & " WHERE conccod = '" & ConcAcumCod & "'"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        ConcAcumNro = rs_Consult!concnro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el codigo interno del concepto " & ConcAcumCod
        HuboError = True
        Exit Sub
    End If
    rs_Consult.Close
Else
    ConcAcumNro = CLng(ConcAcumCod)
End If

'-------------------------------------------------------------------------------------------------
'Configuracion del Directorio de salida
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio de salida."
StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Trim(rs_Consult!sis_direntradas)
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el registro de la tabla sistema nro 1"
    HuboError = True
    Exit Sub
End If
rs_Consult.Close
Flog.writeline
    
    
'-------------------------------------------------------------------------------------------------
'Configuracion del Modelo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Modelo Interface."
StrSql = "SELECT * FROM modelo WHERE modnro = 916"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Directorio & Trim(rs_Consult!modarchdefault)
    Separador = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, ",")
    SeparadorDecimal = IIf(Not IsNull(rs_Consult!modsepdec), rs_Consult!modsepdec, "")
    DescripcionModelo = rs_Consult!moddesc
    
    Flog.writeline Espacios(Tabulador * 1) & "Modelo 916 " & rs_Consult!moddesc
    Flog.writeline Espacios(Tabulador * 1) & "Directorio de Exportacion : " & Directorio
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo 916."
    HuboError = True
    Exit Sub
End If
rs_Consult.Close
Flog.writeline



'-------------------------------------------------------------------------------------------------
'Busqueda de empleados
'-------------------------------------------------------------------------------------------------
'Buscando todos los empleados
StrSql = "SELECT distinct empleado.empleg, empleado.ternro, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2"
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
'Empresa
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
StrSql = StrSql & " AND his_estructura.htetdesde <=" & ConvFecha(FecEstr) & " AND"
StrSql = StrSql & " (his_estructura.htethasta >= " & ConvFecha(FecEstr) & " OR his_estructura.htethasta IS NULL)"
StrSql = StrSql & " AND his_estructura.tenro  = 10"
StrSql = StrSql & " AND his_estructura.estrnro = " & Empnro
'Que el empleado tenga la cuenta bancaria con forma de pago Fondo desempleo
StrSql = StrSql & " INNER JOIN ctabancaria ON ctabancaria.ternro = empleado.ternro"
StrSql = StrSql & " AND ctabancaria.fpagnro = " & FormaPagoNro
'Filtro por los tres niveles de estructuras
If Tenro1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1
    StrSql = StrSql & " AND estact1.htetdesde<=" & ConvFecha(FecEstr) & " AND "
    StrSql = StrSql & " (estact1.htethasta >= " & ConvFecha(FecEstr) & " OR estact1.htethasta IS NULL)"
    If Estrnro1 <> -1 Then
        StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
    End If
End If
If Tenro2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2
    StrSql = StrSql & " AND estact2.htetdesde<=" & ConvFecha(FecEstr) & " AND "
    StrSql = StrSql & " (estact2.htethasta >= " & ConvFecha(FecEstr) & " OR estact2.htethasta IS NULL)"
    If Estrnro2 <> -1 Then
        StrSql = StrSql & " AND estact2.estrnro =" & Estrnro2
    End If
End If
If Tenro3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & Tenro3
    StrSql = StrSql & " AND estact3.htetdesde<=" & ConvFecha(FecEstr) & " AND "
    StrSql = StrSql & " (estact3.htethasta >= " & ConvFecha(FecEstr) & " OR estact3.htethasta IS NULL)"
    If Estrnro3 <> -1 Then
        StrSql = StrSql & " AND estact3.estrnro =" & Estrnro3
    End If
End If

StrSql = StrSql & " WHERE"
StrSql = StrSql & " empleado.empleg >= " & LegDesde
StrSql = StrSql & " AND empleado.empleg <= " & LegHasta
StrSql = StrSql & " AND cabliq.pronro IN (" & ListaProc & ") "
Select Case ParEstado
    Case -1:
        StrSql = StrSql & " AND empleado.empest = -1"
    Case 0:
        StrSql = StrSql & " AND empleado.empest = 0"
End Select

StrSql = StrSql & " ORDER BY empleado.empleg"

OpenRecordset StrSql, rs_Empleados

'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay empleados"
    CEmpleadosAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Empleados: " & CEmpleadosAProc
End If
IncPorc = (100 / CEmpleadosAProc)
        
        
'-------------------------------------------------------------------------------------------------
'Creacion del archivo
'-------------------------------------------------------------------------------------------------
If Not rs_Empleados.EOF Then
    'Seteo el nombre del archivo generado
    Archivo = Directorio & "\ACREFD_" & Format(Date, "dd-mm-yyyy") & "_" & bpronro & ".txt"
    Set fs = CreateObject("Scripting.FileSystemObject")
    'Activo el manejador de errores
    On Error Resume Next
    Set fExport = fs.CreateTextFile(Archivo, True)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs.CreateFolder(Directorio)
        Set fExport = fs.CreateTextFile(Archivo, True)
    End If
    On Error GoTo E_LegFondoDesemple
    Flog.writeline Espacios(Tabulador * 0) & "Archivo Creado: " & "CTALEG_" & Format(Date, "dd-mm-yyyy") & "_" & bpronro & ".txt"
End If
  
        
'-------------------------------------------------------------------------------------------------
'Comienzo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el Procesamiento de Empleados"
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Do While Not rs_Empleados.EOF
    
    ternro = rs_Empleados!ternro
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO EMPLEADO: " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!ternom
    
    
    Monto = 0
    MontoStr = "0"
    'Busco el monto del concepto/acumulador segun corresponda
    If Not Acum Then
        'concepto
        StrSql = "SELECT sum(detliq.dlimonto) monto"
        StrSql = StrSql & " FROM cabliq"
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
        StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro AND cabliq.empleado = " & ternro
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND detliq.concnro = " & ConcAcumNro
        StrSql = StrSql & " WHERE cabliq.pronro IN (" & ListaProc & ")"
    Else
        'acumulador
        StrSql = "SELECT sum(acu_liq.almonto) monto"
        StrSql = StrSql & " FROM cabliq"
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
        StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro
        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro AND acu_liq.acunro = " & ConcAcumNro
        StrSql = StrSql & " WHERE cabliq.pronro IN (" & ListaProc & ")"
    End If
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!Monto) Then Monto = rs_Consult!Monto
    End If
    rs_Consult.Close
    
    'Si no tiene monto no lo considero
    If Monto = 0 Then GoTo Siguiente_empl
    
    'Formateo del monto a lo configurado en el modelo
    MontoStr = CStr(Format(Monto, "##########0.00"))
    MontoStr = Replace(MontoStr, ".", SeparadorDecimal)
    
    'Busco la cuenta bancaria Fondo Desempleo
    CuentaNro = ""
    StrSql = "SELECT ctabancaria.ctabnro"
    StrSql = StrSql & " FROM ctabancaria"
    StrSql = StrSql & " WHERE ctabancaria.ternro = " & ternro
    StrSql = StrSql & " AND ctabancaria.fpagnro = " & FormaPagoNro
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!ctabnro) Then CuentaNro = rs_Consult!ctabnro
    End If
    rs_Consult.Close
    
    
    'Casa N(4)
    Aux_Linea = Format_StrNro(Left(casa, 4), 4, True, "0")
    'Cuenta N(10)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(CuentaNro, 10), 10, True, "0")
    'Transaccio N(3)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(Transacion, 3), 3, True, "0")
    'Filler C(19)
    Aux_Linea = Aux_Linea & Separador & Format_Str(Left(Filler, 19), 19, True, " ")
    'Monto N(13)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(MontoStr, 13), 13, True, "0")
    
    'Imprimo la linea
    fExport.writeline Aux_Linea
    
    
    '---------------------------------------------------------------------------------------------------------------
    'ACTUALIZO EL PROGRESO------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Siguiente_empl:
    rs_Empleados.MoveNext
    
Loop


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close


Set rs_Empleados = Nothing
Set rs_Consult = Nothing

Exit Sub

E_LegFondoDesemple:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: E_LegFondoDesemple"
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

