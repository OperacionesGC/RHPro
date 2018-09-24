Attribute VB_Name = "InterfazOrgyMet"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "24/10/2006"
Global Const UltimaModificacion = " " 'FAF - Version Inicial

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

'NUEVAS
Global EmpErrores As Boolean

Global ArchExp1
Global ArchExp2
Global UsaEncabezado As Integer
Global Encabezado As Boolean

Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion de datos Board meeting.
' Autor      : FAF
' Fecha      : 10/10/2006
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim Sep As String
Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim param
Dim ternro
Dim totalEmpleados
Dim cantRegistros
Dim orden
Dim NroModelo As Long
Dim Directorio As String
Dim Carpeta
Dim fs1
Dim mes As Integer
Dim Anio As Integer
Dim tenro_celula
Dim tenro_linea
Dim fechadesde As Date
Dim fechahasta As Date
Dim objFeriado As New Feriado
Dim codext_celula
Dim dabr_celula
Dim codext_linea
Dim dabr_linea
Dim linestrnro_ant
Dim celestrnro_ant
Dim fec_cel_desde
Dim fec_cel_hasta
Dim fec_lin_desde
Dim fec_lin_hasta
Dim fec_lic_desde
Dim fec_lic_hasta
Dim elfechadesde
Dim elfechahasta
Dim fecCalculo As Date
Dim cant
Dim imprimir As Boolean

Dim rs As New ADODB.Recordset
Dim rsEmpl As New ADODB.Recordset
Dim objRs As New ADODB.Recordset
Dim rsLic As New ADODB.Recordset

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    OpenConnection strconexion, objConn

    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "InterfazOyM" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Interfaz Organización y Métodos: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 140"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        'Obtengo los parametros del proceso
        Parametros = rs!bprcparam
        ArrParametros = Split(Parametros, "@")
        
        Anio = CInt(ArrParametros(0))
        mes = CInt(ArrParametros(1))
        
        fechadesde = CDate("01/" & mes & "/" & Anio)
        If mes = 12 Then
             fechahasta = CDate("31/" & mes & "/" & Anio)
        Else
             fechahasta = DateAdd("d", -1, CDate("01/" & mes + 1 & "/" & Anio))
        End If
        
        'Directorio de exportacion
        StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Directorio = Trim(rs!sis_dirsalidas) & "\InterfazOyM"
        End If
        
        Nombre_Arch = Directorio & "\Licencias.txt"
        Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
        Set fs = CreateObject("Scripting.FileSystemObject")
        'desactivo el manejador de errores
        On Error Resume Next
        
        Set Carpeta = fs.getFolder(Directorio)
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & " no existe. Se creará."
            Err.Number = 0
            Set Carpeta = fs.CreateFolder(Directorio)
            
            If Err.Number <> 0 Then
                Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio " & Directorio & ". Verifique los derechos de acceso o puede crearlo."
                HuboErrores = True
                GoTo Fin
            End If
        End If
        
        Set ArchExp1 = fs.CreateTextFile(Nombre_Arch, True)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio
            HuboErrores = True
            GoTo Fin
        End If
        
        Nombre_Arch = Directorio & "\Lic" & Format(fechadesde, "yyyy") & Format(fechadesde, "mm") & ".txt"
        Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
        Set fs = CreateObject("Scripting.FileSystemObject")
        'desactivo el manejador de errores
        On Error Resume Next
        
        Set Carpeta = fs.getFolder(Directorio)
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & " no existe. Se creará."
            Err.Number = 0
            Set Carpeta = fs.CreateFolder(Directorio)
            
            If Err.Number <> 0 Then
                Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio " & Directorio & ". Verifique los derechos de acceso o puede crearlo."
                HuboErrores = True
                GoTo Fin
            End If
        End If
        
        Set ArchExp2 = fs.CreateTextFile(Nombre_Arch, True)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio
            HuboErrores = True
            GoTo Fin
        End If
        
        On Error GoTo ME_Main
        
        'Obtengo los valores de las 2 estructuras del confrep
        tenro_celula = 0
        tenro_linea = 0
        
        StrSql = "SELECT * FROM confrep WHERE repnro = 181"
        OpenRecordset StrSql, objRs
        Do Until objRs.EOF
            If objRs!confnrocol = 1 Then
                'Tipo de estructura para celula
                tenro_celula = objRs!confval
            End If
            If objRs!confnrocol = 2 Then
                'Tipo de estructura para linea
                tenro_linea = objRs!confval
            End If
            objRs.MoveNext
        Loop
        objRs.Close
        
        If tenro_celula = 0 Then
             Flog.writeline Espacios(Tabulador * 1) & "Error. No se encuenta configurado el tipo de estructura Celula en la columna 1 del confrep (reporte 181)."
        End If
        If tenro_linea = 0 Then
             Flog.writeline Espacios(Tabulador * 1) & "Error. No se encuenta configurado el tipo de estructura Linea en la columna 2 del confrep (reporte 181)."
        End If
        
        StrSql = "SELECT empleado.ternro, h_cel.estrnro celestrnro, h_cel.htetdesde celdesde, h_cel.htethasta celhasta, h_lin.estrnro linestrnro, h_lin.htetdesde lindesde, h_lin.htethasta linhasta FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura h_cel ON h_cel.ternro = empleado.ternro "
        StrSql = StrSql & " AND h_cel.tenro = " & tenro_celula & " AND h_cel.htetdesde <= " & ConvFecha(fechahasta)
        StrSql = StrSql & " AND (h_cel.htethasta is null OR h_cel.htethasta >= " & ConvFecha(fechadesde) & ") "
        StrSql = StrSql & " INNER JOIN his_estructura h_lin ON h_lin.ternro = empleado.ternro "
        StrSql = StrSql & " AND h_lin.tenro = " & tenro_linea & " AND h_lin.htetdesde <= " & ConvFecha(fechahasta)
        StrSql = StrSql & " AND (h_lin.htethasta is null OR h_lin.htethasta >= " & ConvFecha(fechadesde) & ") "
        StrSql = StrSql & " ORDER BY linestrnro, celestrnro"
        OpenRecordset StrSql, rsEmpl
        
        'seteo de las variables de progreso
        Progreso = 0
        cantRegistros = rsEmpl.RecordCount
        totalEmpleados = rsEmpl.RecordCount
        If cantRegistros = 0 Then
            cantRegistros = 1
            Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        End If
        IncPorc = (100 / cantRegistros)
           
        Sep = ";"
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        ' Encabezado Archivo 2
        Call imprimirTexto("""Departamento""", ArchExp2, 2, True)
        Call imprimirTexto(Sep, ArchExp2, 2, True)
        
        Call imprimirTexto("""Nombre""", ArchExp2, 2, True)
        Call imprimirTexto(Sep, ArchExp2, 2, True)
        
        Call imprimirTexto("""Sector""", ArchExp2, 2, True)
        Call imprimirTexto(Sep, ArchExp2, 2, True)
        
        Call imprimirTexto("""Nombre""", ArchExp2, 2, True)
        Call imprimirTexto(Sep, ArchExp2, 2, True)
        
        Call imprimirTexto("""Fecha""", ArchExp2, 2, True)
        Call imprimirTexto(Sep, ArchExp2, 2, True)
        
        Call imprimirTexto("""Licencias""", ArchExp2, 2, True)
        
        'Salto de linea
        ArchExp2.writeline ""
        
        cant = 0
        'Genero por cada empleado un registro
        Do Until rsEmpl.EOF
            ternro = rsEmpl!ternro
            
            linestrnro_ant = rsEmpl!linestrnro
            celestrnro_ant = rsEmpl!celestrnro
            
            'Determino la fecha hasta de la estructura celula
            If EsNulo(rsEmpl!celhasta) Then
                fec_cel_hasta = fechahasta
            Else
                fec_cel_hasta = rsEmpl!celhasta
            End If
       
            'Determino la fecha hasta de la estructura linea
            If EsNulo(rsEmpl!linhasta) Then
                fec_lin_hasta = fechahasta
            Else
                fec_lin_hasta = rsEmpl!linhasta
            End If
            
            'Determino si se superponen los rangos de fechas de ambas estructuras
            If Not (fec_cel_hasta < rsEmpl!lindesde Or fec_lin_hasta < rsEmpl!celdesde) Then
                If rsEmpl!celdesde > rsEmpl!lindesde Then
                    fec_lic_desde = rsEmpl!celdesde
                Else
                    fec_lic_desde = rsEmpl!lindesde
                End If
                
                If fec_lic_desde < fechadesde Then
                    fec_lic_desde = fechadesde
                End If
                
                If fec_cel_hasta < fec_lin_hasta Then
                    fec_lic_hasta = fec_cel_hasta
                Else
                    fec_lic_hasta = fec_lin_hasta
                End If
                
                'Busco si tiene licencias en el rago defecha fec_lic_desde - fec_lic_hasta
                StrSql = "SELECT * FROM emp_lic WHERE empleado = " & rsEmpl!ternro
                StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(fec_lic_hasta) & " AND elfechahasta >= " & ConvFecha(fec_lic_desde)
                OpenRecordset StrSql, rsLic
                
                Do Until rsLic.EOF
                    If rsLic!elfechadesde > fec_lic_desde Then
                        elfechadesde = rsLic!elfechadesde
                    Else
                        elfechadesde = fec_lic_desde
                    End If
                    
                    If rsLic!elfechahasta < fec_lic_hasta Then
                        elfechahasta = rsLic!elfechahasta
                    Else
                        elfechahasta = fec_lic_hasta
                    End If
                    
                    'Calculo la cantidad de dias
                    fecCalculo = elfechadesde
                    Do While fecCalculo <= elfechahasta
                        'Determino si es un dia habil. Que no sea feriado, sabado o domingo
                        If Not (Weekday(fecCalculo) = 7 Or Weekday(fecCalculo) = 1 Or objFeriado.Feriado(fecCalculo, rsEmpl!ternro, False)) Then
                            cant = cant + 1
                        End If
                        
                        fecCalculo = DateAdd("d", 1, fecCalculo)
                        
                    Loop
                    rsLic.MoveNext
                    
                Loop
                
                rsLic.Close
                
            End If
            
            rsEmpl.MoveNext
            imprimir = False
            If rsEmpl.EOF Then
                imprimir = True
            Else
                If rsEmpl!celestrnro <> celestrnro_ant Or rsEmpl!linestrnro <> linestrnro_ant Then
                    imprimir = True
                End If
            End If
            
            
            If imprimir And cant <> 0 Then
                'Busco el codigo externo y la descripcion de la estructura celula
                StrSql = "SELECT * FROM estructura WHERE estrnro=" & celestrnro_ant
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    codext_celula = objRs!estrcodext
                    dabr_celula = objRs!estrdabr
                End If
                objRs.Close
                
                'Busco el codigo externo y la descripcion de la estructura linea
                StrSql = "SELECT * FROM estructura WHERE estrnro=" & linestrnro_ant
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    codext_linea = objRs!estrcodext
                    dabr_linea = objRs!estrdabr
                End If
                objRs.Close
                
                ' Archivo 1
                Call imprimirTexto("""" & codext_linea & """", ArchExp1, Len(codext_linea) + 2, True)
                Call imprimirTexto(Sep, ArchExp1, 2, True)
                
                Call imprimirTexto("""" & codext_celula & """", ArchExp1, Len(codext_celula) + 2, True)
                Call imprimirTexto(Sep, ArchExp1, 2, True)
                
                Call imprimirTexto("""" & Format(fechadesde, "yyyy") & Format(fechadesde, "mm") & """", ArchExp1, 8, True)
                Call imprimirTexto(Sep, ArchExp1, 2, True)
                
                Call imprimirTexto(cant, ArchExp1, Len(cant), True)
                
                'Salto de linea
                ArchExp1.writeline ""
                
                
                ' Archivo 2
                Call imprimirTexto("""" & codext_linea & """", ArchExp2, Len(codext_linea) + 2, True)
                Call imprimirTexto(Sep, ArchExp2, 2, True)
                
                Call imprimirTexto("""" & dabr_linea & """", ArchExp2, Len(dabr_linea) + 2, True)
                Call imprimirTexto(Sep, ArchExp2, 2, True)
                
                Call imprimirTexto("""" & codext_celula & """", ArchExp2, Len(codext_celula) + 2, True)
                Call imprimirTexto(Sep, ArchExp2, 2, True)
                
                Call imprimirTexto("""" & dabr_celula & """", ArchExp2, Len(dabr_celula) + 2, True)
                Call imprimirTexto(Sep, ArchExp2, 2, True)
                
                Call imprimirTexto("""" & Format(fechadesde, "yyyy") & Format(fechadesde, "mm") & """", ArchExp2, 8, True)
                Call imprimirTexto(Sep, ArchExp2, 2, True)
                
                Call imprimirTexto(cant, ArchExp2, Len(cant), True)
                
                'Salto de linea
                ArchExp2.writeline ""
                
                cant = 0
                
            End If
        
        
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
               
            cantRegistros = cantRegistros - 1
               
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
        Loop
        
        ArchExp1.Close
        ArchExp2.Close
              
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    If rs.State = adStateOpen Then rs.Close
    If rsEmpl.State = adStateOpen Then rsEmpl.Close
    If objRs.State = adStateOpen Then objRs.Close
    If rsLic.State = adStateOpen Then rsLic.Close
    Set rs = Nothing
    Set rsEmpl = Nothing
    Set objRs = Nothing
    Set rsLic = Nothing
    
Fin:
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

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
    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo Fin
End Sub

Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 1
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    archivo.Write cadena
    
End Sub
