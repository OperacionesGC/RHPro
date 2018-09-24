Attribute VB_Name = "DotacyRotac"

' __________________________________________________________________________________________________
' Descripcion:
' Autor      : Leticia Amadio
' Fecha      : 02-08-2007
' Ultima Mod : 08-08-2008 - Gustavo Ring - Se modifico el rep. ahora se maneja sin fases
'                                          Se corrigieron todas las consultas que fallaban
' Descripcion:
' ___________________________________________________________________________________________________

Option Explicit


'Global Const Version = "1.00"
'Global Const FechaModificacion = "02-08-2007 "
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "14-08-2008 "
'Global Const UltimaModificacion = " " 'Gustavo Ring

'Global Const Version = "1.02"
'Global Const FechaModificacion = "21-10-2008 "
'Global Const UltimaModificacion = " " 'Gustavo Ring

Global Const Version = "1.03"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'Martin Ferraro - Encriptacion de string connection

' ________________________________________________________________________________________


Global fs, f
'Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset


'NUEVAS
Global EmpErrores As Boolean  ' VERRRRRRRRR

Global filtro As String  ' filtro trae si el empleado es activo o no, y legajo desde -hasta
Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global agencia As Integer
Global fecestr As Date
Global Cargo As Integer   ' verrrrrrrrrrrrr si lo pongo o no o estrnroppal!!!
Global repnro As Integer

Dim tenrofuncion As Integer
Dim estrnroTCIndef As String
Dim estrnroTCPlazoF As String
Dim estrnroTCPract As String
Dim causaRenucia As Integer
Dim causaDespido As Integer
Dim causaFinContr As Integer

Dim estrnro1Ant As Integer
Dim estrnro2Ant As Integer
Dim estrnro3Ant As Integer
Dim cargoAnt As Integer
Dim estcargo As Integer
Dim listaTernoAct As String
Dim listaTernoAnt As String
Dim fecestrAnt As Date

Dim movimientos(20) As Integer


Dim IdUser As String
Dim bpfecha As Date
Dim bphora As String

        ' Global fecestr As String
'Global TituloRep As String


Private Sub Main()
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String

Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim totalEmpleados
Dim cantRegistros

Dim Desde As Date
Dim Hasta As Date
Dim fecestrAnt As Date

On Error GoTo ME_Main

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "DotacionyRotacion" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    Flog.writeline
    
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
    
    Flog.writeline "Inicio Proceso de Dotacion y Rotacion de Personal : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        IdUser = rs!IdUser
        bpfecha = rs!bprcfecha
        bphora = rs!bprchora
        Parametros = rs!bprcparam
        
        ArrParametros = Split(Parametros, "@")
             
        Call levantarParametros(ArrParametros)
        
          
        'Cargo la configuracion del reporte
        Call CargarConfiguracionReporte(repnro)

              
        cantRegistros = 0
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 "
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        ' _____________________________________________________
        ' armar consulta Ppal según Filtro - empls con estructuras activas a la Fecha
        ' ____________________________________________________
        Call cargarConsulta(StrSql, fecestr)
        OpenRecordset StrSql, rs1
       
        'seteo de las variables de progreso
        Progreso = 0
        cantRegistros = rs1.RecordCount
        totalEmpleados = rs1.RecordCount
           
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados para el Filtro."
        End If
        IncPorc = (100 / cantRegistros)
          
        'Inicializo variables de la estructura
        Call inicVariables
        Call inicMovimientos
        
        If Not rs1.EOF Then
        
            Call InsertarDatosCab
            
            cargoAnt = rs1!estrnrocargo
            If tenro1 <> 0 Then
            estrnro1Ant = rs1!estrnro1
            End If
            If tenro2 <> 0 Then
            estrnro2Ant = rs1!estrnro2
            End If
            If tenro3 <> 0 Then
            estrnro3Ant = rs1!estrnro3
            End If
            
            
            Desde = primer_dia_mes(Month(fecestr), Year(fecestr))
            Hasta = ultimo_dia_mes(Month(fecestr), Year(fecestr)) ' ultimo_dia_mes(mes_actual, anio_actual)
            If Month(fecestr) = 1 Then
                fecestrAnt = ultimo_dia_mes(12, Year(fecestr) - 1)
            Else
                fecestrAnt = ultimo_dia_mes(Month(fecestr) - 1, Year(fecestr))
            End If
            
            
            Do While Not rs1.EOF
                
                If EstructurasIguales() Then
                    listaTernoAct = listaTernoAct & ", " & rs1!ternro
                Else
                
                    ' buscar dotacion mes anterior , con la mismas estructuras (estructuras anteriores)
                    Call dotacionAnterior(fecestrAnt, listaTernoAnt, movimientos)
                    Call ingresos(listaTernoAct, Desde, Hasta, movimientos)
                    Call TrasferenciaFunc(Desde, Hasta, listaTernoAct, movimientos)
                    Call fasesBaja(listaTernoAnt, Desde, Hasta, movimientos)
                    Call dotacTipoContrato(fecestr, listaTernoAct, movimientos)
                    Call dotacionAntNeta(fecestrAnt, listaTernoAnt, movimientos) ' solo contratos indef
                    Call InsertarDatosDet
                    Call actualizarEstrAnt
                    Call inicMovimientos
                    listaTernoAct = rs1!ternro
                    
                End If
            
                
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                cantRegistros = cantRegistros - 1
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados)
                StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                objConn.Execute StrSql, , adExecuteNoRecords
     
            rs1.MoveNext
            Loop
            Call dotacionAnterior(fecestrAnt, listaTernoAnt, movimientos)
            Call ingresos(listaTernoAct, Desde, Hasta, movimientos)
            Call TrasferenciaFunc(Desde, Hasta, listaTernoAct, movimientos)
            Call fasesBaja(listaTernoAnt, Desde, Hasta, movimientos)
            Call dotacTipoContrato(fecestr, listaTernoAct, movimientos)
            Call dotacionAntNeta(fecestrAnt, listaTernoAnt, movimientos)
            
         
            Call InsertarDatosDet
            
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            cantRegistros = cantRegistros - 1
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados)
            StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        
        End If
        
                  
        rs1.Close
        
    
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
          
    'If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    'Set rs_Modelo = Nothing

    'Actualizo el estado del proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Flog.writeline "cant open " & Cantidad_de_OpenRecordset

    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
    
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
    
End Sub

' ___________________________________________________________________________________________________
' procedimiento que inserta los dato cabecera en la tabla
' faltaria chequear si ya existe la cabecera, no volverla a insertar o borra lo que existe de antes
' ___________________________________________________________________________________________________
Sub InsertarDatosCab()
Dim Campos As String
Dim Valores As String

On Error GoTo MError

Flog.writeline " "
Flog.writeline Espacios(Tabulador * 1) & "Insertar datos de la cabecera  "


Campos = " (bpronro,Fecha,Hora, iduser, repnro , cargo, filtro,  "    'Fecha - por ahora la del proceso -
Campos = Campos & " tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3, agencia, fecestr )"

Valores = "("
Valores = Valores & NroProceso & "," & ConvFecha(bpfecha) & ",'" & bphora & "', '" & IdUser & "' , "
Valores = Valores & repnro & "," & Cargo & ", '" & filtro & "' , "
Valores = Valores & tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & estrnro3 & ","
Valores = Valores & agencia & "," & ConvFecha(fecestr)
Valores = Valores & ")"

StrSql = " INSERT INTO rep_dotac_rotac_cab " & Campos & " VALUES " & Valores
objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub


MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
 
 
 ' ____________________________________________________________
' procedimiento que inserta los datos detalle  en la tabla
' ____________________________________________________________
Sub InsertarDatosDet()
Dim Campos As String
Dim Valores As String

On Error GoTo MError

Flog.writeline Espacios(Tabulador * 2) & "Insertar Datos del Detalle: cargo: " & cargoAnt & " estrnro1: " & estrnro1Ant & " estrnro2: " & estrnro2Ant & " estrnro3: " & estrnro3Ant

Campos = " (bpronro, estrcargo, estrnro1,  estrnro2,  estrnro3, "
Campos = Campos & " dotmesantf, dotmesantm , ingresof, ingresom, practicasf, practicasm, "
Campos = Campos & " renuciaf, renuciam, functraspf, functraspm,  "
Campos = Campos & " despidof, despidom , fincontrf, fincontrm, "
Campos = Campos & " contrindef, contrpfijo, contrpract, dotnetaant"
Campos = Campos & " )"

Valores = "("
Valores = Valores & NroProceso & "," & cargoAnt & "," & estrnro1Ant & "," & estrnro2Ant & "," & estrnro3Ant & ","
Valores = Valores & movimientos(0) & "," & movimientos(1) & "," & movimientos(2) & "," & movimientos(3) & "," & movimientos(4) & "," & movimientos(5) & ","
Valores = Valores & movimientos(6) & "," & movimientos(7) & "," & movimientos(8) & "," & movimientos(9) & ","
Valores = Valores & movimientos(10) & "," & movimientos(11) & "," & movimientos(12) & "," & movimientos(13) & ","
Valores = Valores & movimientos(14) & "," & movimientos(15) & "," & movimientos(16) & "," & movimientos(0) + movimientos(1)
Valores = Valores & ")"

StrSql = " INSERT INTO rep_dotac_rotac_det " & Campos & " VALUES " & Valores
objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub



' ____________________________________________________________
' procedimiento que
' ____________________________________________________________
Sub levantarParametros(ArrParametros)

On Error GoTo ME_param


    filtro = ArrParametros(0)
    tenro1 = CInt(ArrParametros(1))
    estrnro1 = CInt(ArrParametros(2))
    tenro2 = CInt(ArrParametros(3))
    estrnro2 = CInt(ArrParametros(4))
    tenro3 = CInt(ArrParametros(5))
    estrnro3 = CInt(ArrParametros(6))
    agencia = CInt(ArrParametros(7))
    fecestr = ArrParametros(8)
    Cargo = CInt(ArrParametros(9))
    repnro = CInt(ArrParametros(10))

       
    Flog.writeline Espacios(Tabulador * 0) & "PARAMETROS"
    Flog.writeline Espacios(Tabulador * 0) & "Filtro: " & filtro
    Flog.writeline Espacios(Tabulador * 0) & "Tenro1: " & tenro1
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro1: " & estrnro1
    Flog.writeline Espacios(Tabulador * 0) & "Tenro2: " & tenro2
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro2: " & estrnro2
    Flog.writeline Espacios(Tabulador * 0) & "Tenro3: " & tenro3
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro3: " & estrnro3
    Flog.writeline Espacios(Tabulador * 0) & "Agencia: " & agencia
    Flog.writeline Espacios(Tabulador * 0) & "Fecha p/Estruct: " & fecestr
    Flog.writeline Espacios(Tabulador * 0) & "Cargo: " & Cargo
    Flog.writeline Espacios(Tabulador * 0) & "Nro Reporte: " & repnro

Exit Sub

ME_param:
    Flog.writeline "    Error: Error en la carga de Parametros "
    
End Sub



' ____________________________________________________________
' procedimiento que
' ____________________________________________________________
Sub inicVariables()

Flog.writeline Espacios(Tabulador * 1) & "Inicializa variables.  "
cargoAnt = 0
estrnro1Ant = 0
estrnro2Ant = 0
estrnro3Ant = 0

listaTernoAct = 0
listaTernoAnt = 0

End Sub


' ____________________________________________________________
' procedimiento que
' movimientos(0-1)   - dotmesantf - dotmesantf - dotacion mes anterior
' movimientos(2-3)   - ingresof - ingresom - altas en fases
' movimientos(4-5)   - practicasf - practicasm - altas
' movimientos(6-7)   - renuciaf - renuciam - bajas
' movimientos(8-9)   - functraspf - functraspm
' movimientos(10-11) - despidof - despidom
' movimientos(12-13) - fincontrf - fincontrm
' movimientos(14) - contrindef
' movimientos(15) - contrpfijo
' movimientos(16) - contrpract
' movimientos(17) - dotnetaant
' ____________________________________________________________
Sub inicMovimientos()
Dim I


For I = 0 To 20
    movimientos(I) = 0
Next

End Sub

' ____________________________________________________________
' procedimiento que
' conftipo
' I : tipo Contrato Indefinido
' P : tipo Contrato Practicas
' F : tipo Contrato Plazo Fijo
' FR : fase - causa baja: renucia
' FD : fase - causa baja: despido
' FF : fase - causa baja: fin de contrato
' TEF : Tipo estructura Funcion
' ____________________________________________________________

Sub CargarConfiguracionReporte(repnro)

'Dim I 'Dim columnaActual
'Dim Nro_col 'Dim Valor As Long
Dim rs2 As New ADODB.Recordset

On Error GoTo ME_conf

Flog.writeline Espacios(Tabulador * 1) & "Buscar la configuracion del Reporte - confrep  "

estrnroTCIndef = "0"
estrnroTCPract = "0"
estrnroTCPlazoF = "0"

tenrofuncion = 0
causaRenucia = 0
causaDespido = 0
causaFinContr = 0


StrSql = " SELECT * FROM confrep WHERE confrep.repnro= " & repnro
StrSql = StrSql & " ORDER BY confnrocol "
OpenRecordset StrSql, rs2


If rs2.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & " *** No se encontro la configuracion del reporte " & repnro
    Exit Sub
Else

    Do While Not rs2.EOF
       
        Select Case Trim(rs2!conftipo)
        
        Case "I"
            estrnroTCIndef = estrnroTCIndef & "," & rs2!confval
        Case "P"
            estrnroTCPract = estrnroTCPract & "," & rs2!confval
        Case "F"
            estrnroTCPlazoF = estrnroTCPlazoF & "," & rs2!confval
        Case "FR"
            causaRenucia = rs2!confval
        Case "FD"
            causaDespido = rs2!confval
        Case "FF"
            causaFinContr = rs2!confval
        Case "TEF"
            tenrofuncion = rs2!confval
        End Select
       
       rs2.MoveNext
       
    Loop

Flog.writeline Espacios(Tabulador * 1) & "Datos Configurado en el Reporte: "
Flog.writeline Espacios(Tabulador * 2) & "I -  Tipo Contrato Indefinido - Estructuras: " & estrnroTCIndef
Flog.writeline Espacios(Tabulador * 2) & "P -  Tipo Contrato Practicas  - Estructuras: " & estrnroTCPract
Flog.writeline Espacios(Tabulador * 2) & "F -  Tipo Contrato Plazo Fijo - Estructuras: " & estrnroTCPlazoF

Flog.writeline Espacios(Tabulador * 2) & "FR - Fase - causa baja: Renucia: " & causaRenucia
Flog.writeline Espacios(Tabulador * 2) & "FD - Fase - causa baja: Despido: " & causaDespido
Flog.writeline Espacios(Tabulador * 2) & "FF - Fase - causa baja: Fin de Contrato: " & causaFinContr
Flog.writeline Espacios(Tabulador * 2) & "TEF - Tipo estructura Funcion: " & tenrofuncion

End If

rs2.Close
    


Exit Sub

ME_conf:
    ' Flog.Writeline "    Error - Empleado: " & Empleado
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


' ____________________________________________________________
' funcion que
' ____________________________________________________________
Function EstructurasIguales() As Boolean
Dim igual As String

igual = "SI"

If cargoAnt <> rs1!estrnrocargo Then
    igual = "NO"
End If

If tenro1 <> 0 Then
    If estrnro1Ant <> rs1!estrnro1 Then
        igual = "NO"
    End If
End If
If tenro2 <> 0 Then
    If estrnro2Ant <> rs1!estrnro2 Then
        igual = "NO"
    End If
End If
If tenro3 <> 0 Then
    If estrnro3Ant <> rs1!estrnro3 Then
        igual = "NO"
    End If
End If

If igual = "NO" Then
    EstructurasIguales = False
Else
    EstructurasIguales = True
End If
            
            
End Function

' ____________________________________________________________
' procedimiento que
' ____________________________________________________________
Sub actualizarEstrAnt()

cargoAnt = rs1!estrnrocargo

If tenro1 <> 0 Then
estrnro1Ant = rs1!estrnro1
End If
If tenro2 <> 0 Then
estrnro2Ant = rs1!estrnro2
End If
If tenro3 <> 0 Then
estrnro3Ant = rs1!estrnro3
End If

End Sub

' ____________________________________________________________
' procedimiento que
' ____________________________________________________________
Sub cargarConsulta(ByRef StrSql As String, ByVal Fecha As Date)

Dim StrAgencia As String
Dim StrSelect As String
Dim strjoin As String
Dim StrOrder As String
Dim fecdes As String
Dim fechas As String

On Error GoTo ME_armarsql

StrSql = ""
StrSelect = ""
strjoin = ""
StrOrder = ""



fecdes = primer_dia_mes(Month(fecestr), Year(fecestr))
fechas = ultimo_dia_mes(Month(fecestr), Year(fecestr))

StrAgencia = "" ' cuando queremos todos los empleados

If agencia = "-1" Then
    StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
    StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
    StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
Else
    If agencia = "-2" Then
        StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
        StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
    Else
        If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
            StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
            StrAgencia = StrAgencia & "  AND (agencia.htetdesde<=" & ConvFecha(Fecha)
            StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
        End If
    End If
End If
 
 
 
If tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    
    StrSelect = StrSelect & "  , estact1.tenro tenro1, estact1.estrnro estrnro1, estructura1.estrdabr  estrdabr1 "
    
    strjoin = strjoin & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    strjoin = strjoin & "  AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro1 <> 0 Then
        strjoin = strjoin & " AND estact1.estrnro =" & estrnro1
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro1, estrnro1 "
End If

If tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel
    StrSelect = StrSelect & " , estact2.tenro tenro2, estact2.estrnro estrnro2, estructura2.estrdabr estrdabr2  "
    
    strjoin = strjoin & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    strjoin = strjoin & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro2 <> 0 Then
        strjoin = strjoin & " AND estact2.estrnro =" & estrnro2
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro2, estrnro2 "
End If

If tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles
    StrSelect = StrSelect & " , estact3.tenro tenro3, estact3.estrnro estrnro3, estructura3.estrdabr estrdabr3 "

    strjoin = strjoin & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & tenro3
    strjoin = strjoin & "   AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta IS NULL OR  estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
        strjoin = strjoin & " AND estact3.estrnro =" & estrnro3
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "

    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro3, estrnro3 "
    
End If

                      
StrSql = " SELECT DISTINCT empleado.ternro, empleado.empleg, cargo.estrnro estrnrocargo, estrcargo.estrcodext   "   '  empleado.empest, tercero.tersex,
StrSql = StrSql & StrSelect
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN his_estructura cargo ON empleado.ternro = cargo.ternro  AND cargo.tenro = 4 " ' cargo- tenro=4, puesto base rhpro
StrSql = StrSql & "  AND (cargo.htetdesde<=" & ConvFecha(Fecha) & " AND (cargo.htethasta IS NULL OR cargo.htethasta>=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & " INNER JOIN estructura estrcargo ON estrcargo.estrnro = cargo.estrnro "
StrSql = StrSql & strjoin
StrSql = StrSql & " WHERE " & filtro & StrAgencia
StrSql = StrSql & " ORDER BY "
If StrOrder <> "" Then
    StrSql = StrSql & StrOrder & ", "
End If
StrSql = StrSql & " estrcargo.estrcodext "
     
  
 
Exit Sub

ME_armarsql:
    Flog.writeline " Error: Armar consulta del Filtro.- " & Err.Description
    Flog.writeline " Consulta Armada " & StrSql
    
    
End Sub

' ____________________________________________________________________________________
' procedimiento que busca la dotacion del mes anterior y la carga en movimientos
' movimientos(0-1)   - dotmesantf - dotmesantm - dotacion mes anterior
' ____________________________________________________________________________________
Sub dotacionAnterior(Fecha, ByRef listaTernoAnt, ByRef movimientos)  ' fecha: fecestrAnt

Dim rs2 As New ADODB.Recordset
Dim StrAgencia As String
Dim StrSelect As String
Dim strjoin As String
Dim StrOrder As String

Dim dotacAntF As Integer
Dim dotacAntM As Integer

On Error GoTo ME_dotAnt

StrSql = ""
StrSelect = ""
strjoin = ""
StrOrder = ""

StrAgencia = "" ' cuando queremos todos los empleados

listaTernoAnt = "0"
dotacAntF = 0
dotacAntM = 0

If agencia = "-1" Then
    StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
    StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
    StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
Else
    If agencia = "-2" Then
        StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
        StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
    Else
        If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
            StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
            StrAgencia = StrAgencia & "  AND (agencia.htetdesde<=" & ConvFecha(Fecha)
            StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
        End If
    End If
End If

 
If tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    
    StrSelect = StrSelect & "  , estact1.tenro tenro1, estact1.estrnro estrnro1, estructura1.estrdabr  estrdabr1 "
    
    strjoin = strjoin & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    strjoin = strjoin & "  AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    strjoin = strjoin & "  AND estact1.estrnro =" & estrnro1Ant
    strjoin = strjoin & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro1, estrnro1 "
End If

If tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel
    StrSelect = StrSelect & " , estact2.tenro tenro2, estact2.estrnro estrnro2, estructura2.estrdabr estrdabr2  "
    
    strjoin = strjoin & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    strjoin = strjoin & "  AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    strjoin = strjoin & "  AND estact2.estrnro =" & estrnro2Ant
    strjoin = strjoin & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro2, estrnro2 "
End If


If tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles
    StrSelect = StrSelect & " , estact3.tenro tenro3, estact3.estrnro estrnro3, estructura3.estrdabr estrdabr3 "

    strjoin = strjoin & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & tenro3
    strjoin = strjoin & "   AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta IS NULL OR  estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    strjoin = strjoin & " AND estact3.estrnro =" & estrnro3Ant
    strjoin = strjoin & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "

    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro3, estrnro3 "
    
End If

                              
StrSql = " SELECT DISTINCT empleado.ternro, empleado.empleg, tercero.tersex, estrcargo.estrcodext  "   '  empleado.empest, tercero.tersex,
StrSql = StrSql & StrSelect
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & " INNER JOIN his_estructura cargo ON empleado.ternro = cargo.ternro  AND cargo.tenro = 4 " ' cargo- tenro=4, puesto base rhpro
StrSql = StrSql & " AND (cargo.htetdesde<=" & ConvFecha(Fecha) & " AND (cargo.htethasta IS NULL OR cargo.htethasta>=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & " INNER JOIN estructura estrcargo ON estrcargo.estrnro = cargo.estrnro "
StrSql = StrSql & " AND estrcargo.estrnro= " & cargoAnt
StrSql = StrSql & strjoin
StrSql = StrSql & " WHERE " & filtro & StrAgencia
StrSql = StrSql & " ORDER BY estrcargo.estrcodext "

OpenRecordset StrSql, rs2

Do While Not rs2.EOF
    
    listaTernoAnt = listaTernoAnt & "," & rs2!ternro
    
    If rs2!tersex = 0 Then
        dotacAntF = dotacAntF + 1
    Else
        dotacAntM = dotacAntM + 1
    End If
    
rs2.MoveNext
Loop

rs2.Close

' carga resultados en el arreglo
movimientos(0) = dotacAntF
movimientos(1) = dotacAntM


Exit Sub


ME_dotAnt:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Sub

Sub ingresos(listaTernro, Desde, Hasta, ByRef movimientos)

' 13/08/2008 Gustavo Ring - Cuenta las altas de la estructura tipo de contrato seleccionadas en el confrep según el puesto

Dim strPracticas   As String
Dim rs2 As New ADODB.Recordset
Dim StrAgencia As String

On Error GoTo ME_fases

    StrAgencia = "" ' cuando queremos todos los empleados

    If agencia = "-1" Then
        StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Hasta)
        StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Hasta) & ")) )"
    Else
        If agencia = "-2" Then
            StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Hasta)
            StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Hasta) & ")) )"
        Else
            If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
                StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
                StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
                StrAgencia = StrAgencia & "  AND (agencia.htetdesde<=" & ConvFecha(Hasta)
                StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Hasta) & ")) )"
            End If
        End If
    End If
    
    strPracticas = "SELECT ternro "
    strPracticas = strPracticas & " FROM his_estructura tipocontrato "
    strPracticas = strPracticas & " WHERE tipocontrato.ternro = empleado.ternro "
    strPracticas = strPracticas & "  AND tipocontrato.estrnro IN (" & estrnroTCPract & ")"
    strPracticas = strPracticas & "  AND (tipocontrato.htetdesde<= " & ConvFecha(Hasta)
    strPracticas = strPracticas & "  AND (tipocontrato.htethasta IS NULL OR tipocontrato.htethasta>=" & ConvFecha(Hasta) & " )) "
    
'----------------------------------------------------------------------------------------
' Calculo los ingresos que tuvo el tipo de estructura que no estan en practicas (mujeres)
'----------------------------------------------------------------------------------------

    StrSql = "SELECT count(Distinct empleg) cantf "
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
    StrSql = StrSql & " INNER JOIN estructura     ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
    StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(fecestr) & "))"
    StrSql = StrSql & " WHERE tersex = 0 "
    StrSql = StrSql & " AND his_estructura.tenro = " & tenrofuncion  ' De la estructura del tipo de estructura seleccionado en el confrep
    StrSql = StrSql & " AND (his_estructura.htetdesde >=" & ConvFecha(Desde) & " AND his_estructura.htetdesde <= " & ConvFecha(Hasta) & ")"
    StrSql = StrSql & " AND empleado.ternro IN ( " & listaTernro & ")"
    StrSql = StrSql & " AND his_estructura.ternro NOT IN ( " & strPracticas & ")"
    StrSql = StrSql & StrAgencia
    OpenRecordset StrSql, rs2
    
    If Not rs2.EOF Then
        movimientos(2) = rs2!cantf
    Else
        movimientos(2) = 0
    End If
    rs2.Close
    
'-------------------------------------------------------------------------------------------
' Calculo los ingresos que tuvo el tipo de estructura que no estan en practicas (hombres)
'-------------------------------------------------------------------------------------------
    
    StrSql = "SELECT count(Distinct empleg) cantm "
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
    StrSql = StrSql & " INNER JOIN estructura     ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
    StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(fecestr) & "))"
    StrSql = StrSql & " WHERE tersex = -1 "
    StrSql = StrSql & " AND his_estructura.tenro = " & tenrofuncion  ' El tipo de estructura seleccionado en el confrep
    StrSql = StrSql & " AND (his_estructura.htetdesde >=" & ConvFecha(Desde) & " AND his_estructura.htetdesde <= " & ConvFecha(Hasta) & ")"
    StrSql = StrSql & " AND empleado.ternro IN ( " & listaTernro & ")"
    StrSql = StrSql & " AND his_estructura.ternro NOT IN ( " & strPracticas & ")"
    StrSql = StrSql & StrAgencia
    OpenRecordset StrSql, rs2
    
    If Not rs2.EOF Then
        movimientos(3) = rs2!cantm
    Else
        movimientos(3) = 0
    End If
    rs2.Close
      
'--------------------------------------------------------------------------------------------------
' Calculo los ingresos que tuvo el tipo de estructura que estan en practicas (mujeres)
'---------------------------------------------------------------------------------------------------

    
    StrSql = "SELECT count(Distinct empleg) cantm "
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
    StrSql = StrSql & " INNER JOIN estructura     ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
    StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(fecestr) & "))"
    StrSql = StrSql & " WHERE tersex = 0 "
    StrSql = StrSql & " AND his_estructura.tenro = " & tenrofuncion  ' El tipo de estructura seleccionado en el confrep
    StrSql = StrSql & " AND (his_estructura.htetdesde >=" & ConvFecha(Desde) & " AND his_estructura.htetdesde <= " & ConvFecha(Hasta) & ")"
    StrSql = StrSql & " AND empleado.ternro IN ( " & listaTernro & ")"
    StrSql = StrSql & " AND his_estructura.ternro IN ( " & strPracticas & ")"
    StrSql = StrSql & StrAgencia
    OpenRecordset StrSql, rs2
    
    If Not rs2.EOF Then
        movimientos(4) = rs2!cantm
    Else
        movimientos(4) = 0
    End If
    rs2.Close

'-------------------------------------------------------------------------------------------------------
' Calculo los ingresos que tuvo el tipo de estructura que estan en prácticas (hombres)
'-----------------------------------------------------------------------------------------------------------
    
    StrSql = "SELECT count(Distinct empleg) cantm "
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
    StrSql = StrSql & " INNER JOIN estructura     ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
    StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(fecestr) & "))"
    StrSql = StrSql & " WHERE tersex = -1 "
    StrSql = StrSql & " AND his_estructura.tenro = " & tenrofuncion  ' El tipo de estructura seleccionado en el confrep
    StrSql = StrSql & " AND (his_estructura.htetdesde >=" & ConvFecha(Desde) & " AND his_estructura.htetdesde <= " & ConvFecha(Hasta) & ")"
    StrSql = StrSql & " AND empleado.ternro IN ( " & listaTernro & ")"
    StrSql = StrSql & " AND his_estructura.ternro IN ( " & strPracticas & ")"
    StrSql = StrSql & StrAgencia
    OpenRecordset StrSql, rs2
    
    If Not rs2.EOF Then
        movimientos(5) = rs2!cantm
    Else
        movimientos(5) = 0
    End If
    rs2.Close

Exit Sub

ME_fases:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Sub

' ____________________________________________________________
' procedimiento  que cuenta las bajas en las fases
' movimientos(6-7)   - renuciaf - renuciam - bajas
' movimientos(10-11) - despidof - despidom
' movimientos(12-13) - fincontrf - fincontrm
' ____________________________________________________________
Sub fasesBaja(listaTerno, Desde, Hasta, ByRef movimientos)
Dim rs2 As New ADODB.Recordset
Dim StrAgencia As String
Dim StrSelect As String
Dim strjoin As String

On Error GoTo ME_fases

' 13/08/2008 Gustavo Ring - Cuenta las bajas de la estructura tipo de contrato seleccionadas en el confrep según el puesto


    StrAgencia = "" ' cuando queremos todos los empleados

    If agencia = "-1" Then
        StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<= fases.bajfec"
        StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>= fases.bajfec )"
    Else
        If agencia = "-2" Then
            StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=  fases.bajfec "
            StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta >=  fases.bajfec ))"
        Else
            If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
                StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
                StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
                StrAgencia = StrAgencia & "  AND (agencia.htetdesde<= fases.bajfec "
                StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>= fases.bajfec )))"
            End If
        End If
    End If
    
    StrSql = "SELECT distinct(empleado.ternro), e1.estrdabr,tersex,e1.estrnro,he1.estrnro, caunro" & StrSelect
    StrSql = StrSql & " FROM  Empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro AND fases.bajfec >= " & ConvFecha(Desde) & " AND fases.bajfec <= " & ConvFecha(Hasta)
    StrSql = StrSql & " INNER JOIN his_estructura h1 ON h1.ternro = empleado.ternro AND h1.estrnro = " & cargoAnt
    StrSql = StrSql & " INNER JOIN estructura     e1 ON e1.estrnro = h1.estrnro AND e1.tenro = h1.tenro AND e1.estrnro = " & cargoAnt
    StrSql = StrSql & " INNER JOIN his_estructura he1 ON he1.ternro = empleado.ternro "
    StrSql = StrSql & " AND he1.tenro = " & tenrofuncion
    StrSql = StrSql & " AND h1.htetdesde <= fases.bajfec AND (h1.htethasta IS NULL OR h1.htethasta >= fases.bajfec) "
    StrSql = StrSql & " AND he1.htetdesde <= fases.bajfec AND (he1.htethasta IS NULL OR he1.htethasta >= fases.bajfec) "

    If tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
        StrSql = StrSql & "  AND (estact1.htetdesde<= fases.bajfec AND (estact1.htethasta IS NULL OR estact1.htethasta>=fases.bajfec))"
        StrSql = StrSql & "  AND estact1.estrnro =" & estrnro1Ant
        StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
    End If

    If tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel
    
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
        StrSql = StrSql & " AND (estact2.htetdesde<=fases.bajfec AND (estact2.htethasta IS NULL OR estact2.htethasta>=fases.bajfec))"
        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2Ant
        StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
    End If

    If tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles

        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & tenro3
        StrSql = StrSql & " AND (estact3.htetdesde<= fases.bajfec AND (estact3.htethasta IS NULL OR  estact3.htethasta>= fases.bajfec))"
        StrSql = StrSql & " AND estact3.estrnro =" & estrnro3Ant
        StrSql = StrSql & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "
    
    End If

    StrSql = StrSql & " WHERE he1.tenro = " & tenrofuncion
    'StrSql = StrSql & " AND fases.bajfec >= " & ConvFecha(Desde) & " AND fases.bajfec <= " & ConvFecha(Hasta)
    
    OpenRecordset StrSql, rs2
    
    movimientos(6) = 0
    movimientos(7) = 0
    movimientos(10) = 0
    movimientos(11) = 0
    movimientos(12) = 0
    movimientos(13) = 0
    
    While Not rs2.EOF
    
        If rs2!tersex = 0 Then
            If rs2!caunro = causaRenucia Then
                movimientos(6) = movimientos(6) + 1
            End If
            If rs2!caunro = causaDespido Then
                movimientos(10) = movimientos(10) + 1
            End If
            If rs2!caunro = causaFinContr Then
                movimientos(12) = movimientos(12) + 1
            End If
        Else
            If rs2!caunro = causaRenucia Then
                movimientos(7) = movimientos(7) + 1
            End If
            If rs2!caunro = causaDespido Then
                movimientos(11) = movimientos(11) + 1
            End If
            If rs2!caunro = causaFinContr Then
                movimientos(13) = movimientos(13) + 1
            End If
        End If
        rs2.MoveNext
    Wend
    
Exit Sub

ME_fases:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Sub
    
' ____________________________________________________________
' procedimiento que
' movimientos(14) - contrindef
' movimientos(15) - contrpfijo
' movimientos(16) - contrpract
' ____________________________________________________________
Sub dotacTipoContrato(Fecha, listaTernoAct, ByRef movimientos)
Dim rs2 As New ADODB.Recordset

On Error GoTo ME_dotac

' Tipo Contrato Indefinido - movimientos(14) - contrindef
StrSql = " SELECT COUNT(DISTINCT empleado.ternro) cant "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN his_estructura tipoContrato ON empleado.ternro = tipoContrato.ternro "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & "   AND (tipoContrato.htetdesde<=" & ConvFecha(Fecha) & " AND (tipoContrato.htethasta IS NULL OR tipoContrato.htethasta>=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & "   AND empleado.ternro IN ( " & listaTernoAct & ")"
StrSql = StrSql & "   AND tipoContrato.estrnro IN (" & estrnroTCIndef & ")"

OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    movimientos(14) = rs2!cant
Else
    movimientos(14) = 0
End If
rs2.Close


' Tipo Contrato Plazo Fijo - movimientos(15) - contrpfijo
StrSql = " SELECT COUNT(DISTINCT empleado.ternro) cant "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & " INNER JOIN his_estructura tipoContrato ON empleado.ternro = tipoContrato.ternro "
StrSql = StrSql & "   AND (tipoContrato.htetdesde<=" & ConvFecha(Fecha) & " AND (tipoContrato.htethasta IS NULL OR tipoContrato.htethasta>=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & "   AND empleado.ternro IN ( " & listaTernoAct & ")"
StrSql = StrSql & "   AND tipoContrato.estrnro IN (" & estrnroTCPlazoF & ")"


OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    movimientos(15) = rs2!cant
Else
    movimientos(15) = 0
End If
rs2.Close


' Tipo Contrato Practicas - movimientos(16) - contrpract
StrSql = " SELECT COUNT(DISTINCT empleado.ternro) cant "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & " INNER JOIN his_estructura tipoContrato ON empleado.ternro = tipoContrato.ternro "
StrSql = StrSql & "   AND (tipoContrato.htetdesde<=" & ConvFecha(Fecha) & " AND (tipoContrato.htethasta IS NULL OR tipoContrato.htethasta>=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & "   AND empleado.ternro IN ( " & listaTernoAct & ")"
StrSql = StrSql & "   AND tipoContrato.estrnro IN (" & estrnroTCPract & ")"


OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    movimientos(16) = rs2!cant
Else
    movimientos(16) = 0
End If
rs2.Close

Exit Sub

ME_dotac:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Sub


' ______________________________________________________________________
' procedimiento que
' Si hay cambio de FUNCION -  Hay cambio de CARGO en DEl CHILE
' movimientos(8-9)   - functraspf - functraspm - traspaso entre funciones
' __________________________________________________________________

Sub TrasferenciaFunc(Desde, Hasta, listaTerno, ByRef movimientos)
Dim rs2 As New ADODB.Recordset
Dim transferenciaF As Integer
Dim transferenciaM As Integer
Dim StrAgencia As String

On Error GoTo ME_transf

transferenciaF = 0
transferenciaM = 0

    StrAgencia = "" ' cuando queremos todos los empleados

    If agencia = "-1" Then
        StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Hasta)
        StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Hasta) & ")) )"
    Else
        If agencia = "-2" Then
            StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Hasta)
            StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Hasta) & ")) )"
        Else
            If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
                StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
                StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
                StrAgencia = StrAgencia & "  AND (agencia.htetdesde<=" & ConvFecha(Hasta)
                StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Hasta) & ")) )"
            End If
        End If
    End If

' busco los cambios los cambios dedel mes cambian, las estructuras son las que tiene el empleado
' al momento del cambio del tipo de estructura configurado. Gustavo Ring

StrSql = " SELECT distinct(empleado.ternro),e1.estrdabr,e2.estrdabr,tersex,"
StrSql = StrSql & " h1.htethasta,h2.htetdesde, ee1.estrnro,ee1.estrdabr"
StrSql = StrSql & " FROM Empleado"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(fecestr) & "))"
StrSql = StrSql & " INNER JOIN his_estructura h1 ON h1.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN estructura     e1 ON e1.estrnro = h1.estrnro AND e1.tenro = h1.tenro "
StrSql = StrSql & " INNER JOIN his_estructura h2 ON h2.ternro = empleado.ternro AND h1.tenro = h2.tenro AND h2.htetdesde >= h1.htethasta "
StrSql = StrSql & " INNER JOIN estructura     e2 ON e2.estrnro = h2.estrnro AND e2.tenro = h2.tenro "
StrSql = StrSql & " INNER JOIN his_estructura he1 ON he1.ternro = empleado.ternro AND he1.estrnro = " & cargoAnt
StrSql = StrSql & " AND he1.tenro = " & tenrofuncion
StrSql = StrSql & " AND he1.htetdesde <= h1.htethasta AND (he1.htethasta IS NULL OR he1.htethasta >= h1.htethasta)"
StrSql = StrSql & " INNER JOIN estructura ee1 ON ee1.estrnro = he1.estrnro AND he1.tenro = ee1.tenro AND ee1.estrnro =" & cargoAnt

If tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    StrSql = StrSql & "  AND (estact1.htetdesde<= h1.htethasta AND (estact1.htethasta IS NULL OR estact1.htethasta>=h1.htethasta))"
    StrSql = StrSql & "  AND estact1.estrnro =" & estrnro1Ant
    StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
End If

If tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel
    
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    StrSql = StrSql & " AND (estact2.htetdesde<=h1.htethasta AND (estact2.htethasta IS NULL OR estact2.htethasta>=h1.htethasta))"
    StrSql = StrSql & " AND estact2.estrnro =" & estrnro2Ant
    StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
End If

If tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles

    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & tenro3
    StrSql = StrSql & " AND (estact3.htetdesde<= h1.htethasta AND (estact3.htethasta IS NULL OR  estact3.htethasta>= h1.htethasta))"
    StrSql = StrSql & " AND estact3.estrnro =" & estrnro3Ant
    StrSql = StrSql & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "
    
End If

StrSql = StrSql & " AND h1.tenro = " & tenrofuncion
StrSql = StrSql & " AND h1.htethasta >= " & ConvFecha(Desde) & " AND h1.htethasta <= " & ConvFecha(Hasta)
 
OpenRecordset StrSql, rs2

Do While Not rs2.EOF
    
            If rs2!tersex = 0 Then
                transferenciaF = transferenciaF + 1
            Else
                transferenciaM = transferenciaM + 1
            End If
    
rs2.MoveNext
Loop

rs2.Close

movimientos(8) = transferenciaF
movimientos(9) = transferenciaM


Exit Sub

ME_transf:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "


End Sub


' ____________________________________________________________
' procedimiento que
' movimientos(17) - dotnetaant
' ____________________________________________________________
Sub dotacionAntNeta(Fecha, listaTernoAnt, ByRef movimientos)
Dim rs2 As New ADODB.Recordset

On Error GoTo ME_dotac
 
' Tipo Contrato Indefinido - movimientos(17) - contrindef - del mes anterior
StrSql = " SELECT COUNT(DISTINCT empleado.ternro) cant "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = tercero.ternro "
StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec IS NULL OR fases.bajfec >=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & " INNER JOIN his_estructura tipoContrato ON empleado.ternro = tipoContrato.ternro "
StrSql = StrSql & "   AND (tipoContrato.htetdesde<=" & ConvFecha(Fecha) & " AND (tipoContrato.htethasta IS NULL OR tipoContrato.htethasta>=" & ConvFecha(Fecha) & "))"
StrSql = StrSql & "   AND empleado.ternro IN ( " & listaTernoAnt & ")"
StrSql = StrSql & "   AND tipoContrato.estrnro IN (" & estrnroTCIndef & ")"


OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    movimientos(17) = rs2!cant
Else
    movimientos(17) = 0
End If
rs2.Close

Exit Sub

ME_dotac:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Sub

Function primer_dia_mes(mes As Integer, Anio As Integer) As Date
Dim aux As String
    primer_dia_mes = C_Date("01/" & mes & "/" & Anio)
    
End Function



Function ultimo_dia_mes(mes As Integer, Anio As Integer) As Date

Dim mes_sgt As Integer
Dim anio_sgt As Integer

    If mes = 12 Then
        mes_sgt = 1
        anio_sgt = Anio + 1
    Else
        mes_sgt = mes + 1
        anio_sgt = Anio
    End If
    
    ultimo_dia_mes = DateAdd("d", -1, primer_dia_mes(mes_sgt, anio_sgt))
    
End Function


