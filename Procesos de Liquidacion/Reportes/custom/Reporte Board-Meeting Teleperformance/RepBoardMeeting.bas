Attribute VB_Name = "RepBoardMeeting"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "10/10/2006"
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

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global tenro1_confrep As Integer
Global tenro2_confrep As Integer
Global tenro3_confrep As Integer
Global fecEstr As String
Global Formato As Integer
Global Modelo As Long
Global TituloRep As String
Global descDesde
Global descHasta
Global ArchExp
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

Dim rs As New ADODB.Recordset
Dim pliqNro As Long
Dim Lista_Pronro As String
Dim Sep As String
Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim param
Dim listapronro
Dim proNro
Dim ternro
Dim arrpronro
Dim Periodos
Dim rsEmpl As New ADODB.Recordset
Dim totalEmpleados
Dim cantRegistros
Dim objRs As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim orden
Dim NroModelo As Long
Dim Directorio As String
Dim Carpeta
'Dim Nombre_Arch As String
'Dim rs As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim fs1


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
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "RepBoardMeeting" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Exportación para el Reporte Board Meeting : " & Now
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
    StrSql = StrSql & " AND btprcnro = 138"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(0))
       estrnro1 = CInt(ArrParametros(1))
       tenro2 = CInt(ArrParametros(2))
       estrnro2 = CInt(ArrParametros(3))
       tenro3 = CInt(ArrParametros(4))
       estrnro3 = CInt(ArrParametros(5))
       fecEstr = ArrParametros(6)
       
       'Directorio de exportacion
       StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
       If rs.State = adStateOpen Then rs.Close
       OpenRecordset StrSql, rs
       If Not rs.EOF Then
          Directorio = Trim(rs!sis_dirsalidas) & "\ExpBoardMeeting"
       End If
     
       Nombre_Arch = Directorio & "\Rep_Board_Meeting.csv"
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
    
       Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
    
       If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio
            HuboErrores = True
            GoTo Fin
       End If

       On Error GoTo ME_Main
       
'       Nombre_Arch = Directorio & "\Rep_Board_Meeting.csv"
'       Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
'       Set fs = CreateObject("Scripting.FileSystemObject")
'       On Error Resume Next
'       If Err.Number <> 0 Then
'          Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
'          Set Carpeta = fs1.CreateFolder(Directorio)
'       End If
'       'desactivo el manejador de errores
'       Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
      
       'Obtengo los empleados sobre los que tengo que generar los recibos
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
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
       
       Encabezado = True
       
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
           EmpErrores = False
           ternro = rsEmpl!ternro
           orden = rsEmpl!Estado
              
           Flog.writeline Espacios(Tabulador * 1) & "Se comienza a procesar los datos"
    
           'Genero una entrada para el empleado por cada periodo
           Call Generar_Archivo(ternro, Sep)
              
           Flog.writeline Espacios(Tabulador * 1) & "Se Terminaron de Procesar los datos del empleado " & ternro
    
           'Actualizo el estado del proceso
           TiempoAcumulado = GetTickCount
              
           cantRegistros = cantRegistros - 1
              
           StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                 
           objConn.Execute StrSql, , adExecuteNoRecords
              
           'Si se generaron todos los datos del empleado correctamente lo borro
           If Not EmpErrores Then
               StrSql = " DELETE FROM batch_empleado "
               StrSql = StrSql & " WHERE bpronro = " & NroProceso
               StrSql = StrSql & " AND ternro = " & ternro
             
               objConn.Execute StrSql, , adExecuteNoRecords
           End If
            
           rsEmpl.MoveNext
       Loop
       
       ArchExp.Close
              
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    Set rs_Modelo = Nothing
    
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
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub

Function cambiarFormato(cadena, separador)

 Dim l_salida
 
 l_salida = cadena

 If (InStr(cadena, ".")) Then
    If separador = "," Then
       l_salida = Replace(cadena, ".", ",")
    End If
 Else
     If (InStr(cadena, ",")) Then
        If separador = "." Then
           l_salida = Replace(cadena, ",", ".")
        End If
     End If
 End If

 cambiarFormato = l_salida

End Function 'cambiarFormato(cadena,separador)

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
Sub imprimirTextoConCeros(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 1
        cadena = "0"
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    archivo.Write cadena
    
End Sub



Private Sub Generar_Archivo(ternro, Sep)
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : FAF
' Fecha      : 01/10/2006
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim I As Integer
Dim cantRegistros As Long

Dim Legajo As Long
Dim Apellido As String
Dim FechaAlta
Dim FechaBaja
Dim PuestoContable
Dim CentroCosto
Dim NroCC
Dim Empresa
Dim FecIngEmp
Dim Estado
Dim CausaBaja

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim testrnomb1
Dim testrnomb2
Dim testrnomb3

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsCuenta As New ADODB.Recordset
Dim lineaEncabezado As String

    On Error GoTo ME_Local

    estrnomb1 = ""
    estrnomb2 = ""
    estrnomb3 = ""

    '------------------------------------------------------------------
    'Busco los datos del empleado
    '------------------------------------------------------------------
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE ternro = " & ternro
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
        Legajo = rsConsult!empleg
        Apellido = rsConsult!terape & " " & rsConsult!terape2
        FecIngEmp = rsConsult!empfaltagr
        Estado = rsConsult!empest
    Else
        Flog.writeline "Error al obtener los datos del empleado. SQL --> " & StrSql
    End If
    
    rsConsult.Close

    '------------------------------------------------------------------
    'Busco los datos del tipos de estructura 1
    '------------------------------------------------------------------
    If tenro1 <> 0 Then
        
        StrSql = " SELECT estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro1
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
        End If
                   
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           estrnomb1 = rsConsult!estrdabr
           testrnomb1 = rsConsult!tedabr
        End If
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del tipos de estructura 2
    '------------------------------------------------------------------
    If tenro2 <> 0 Then
        
        StrSql = " SELECT estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro2
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
        End If
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           estrnomb2 = rsConsult!estrdabr
           testrnomb2 = rsConsult!tedabr
        End If
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del tipos de estructura 3
    '------------------------------------------------------------------
    If tenro3 <> 0 Then
        
        StrSql = " SELECT estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro3
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
        End If
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           estrnomb3 = rsConsult!estrdabr
           testrnomb3 = rsConsult!tedabr
        End If
    End If
    
    '------------------------------------------------------------------
    'Busco los datos de la empresa
    '------------------------------------------------------------------
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 10 "
    StrSql = StrSql & "    AND htethasta is null "
        
    OpenRecordset StrSql, rsConsult
    Empresa = " "
    If Not rsConsult.EOF Then
        Empresa = rsConsult!estrdabr
    Else
        StrSql = " SELECT estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 10 "
        StrSql = StrSql & "    ORDER BY htetdesde DESC "
            
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
            Empresa = rsConsult!estrdabr
        End If
        
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del CC
    '------------------------------------------------------------------
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 5 "
    StrSql = StrSql & "    AND htethasta is null "
        
    OpenRecordset StrSql, rsConsult
    CentroCosto = " "
    If Not rsConsult.EOF Then
        CentroCosto = rsConsult!estrdabr
        NroCC = rsConsult!estrcodext
    Else
        StrSql = " SELECT estrdabr, estrcodext "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 5 "
        StrSql = StrSql & "    ORDER BY htetdesde DESC "
            
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
            CentroCosto = rsConsult!estrdabr
            NroCC = rsConsult!estrcodext
        End If
        
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del Puesto Contable
    '------------------------------------------------------------------
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 48 "
    StrSql = StrSql & "    AND htethasta is null "
        
    OpenRecordset StrSql, rsConsult
    PuestoContable = " "
    If Not rsConsult.EOF Then
        PuestoContable = rsConsult!estrdabr
    Else
        StrSql = " SELECT estrdabr, estrcodext "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 48 "
        StrSql = StrSql & "    ORDER BY htetdesde DESC "
            
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
            PuestoContable = rsConsult!estrdabr
        End If
        
    End If
    
    '------------------------------------------------------------------
    'Busco la Fecha de Alta
    '------------------------------------------------------------------
    StrSql = " SELECT * FROM fases WHERE empleado = " & ternro
    StrSql = StrSql & " ORDER BY bajfec DESC"
    
    OpenRecordset StrSql, rsConsult
    
    FechaAlta = " "
    FechaBaja = " "
    CausaBaja = " "
    If Not rsConsult.EOF Then
       FechaAlta = rsConsult!altfec
       FechaBaja = rsConsult!bajfec
       If Not EsNulo(FechaBaja) And rsConsult!caunro <> 32 And rsConsult!caunro <> 36 Then
            StrSql = "SELECT * FROM causa WHERE caunro = " & rsConsult!caunro
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                CausaBaja = rsConsult!caudes
            End If
       End If
    End If
    
    rsConsult.Close
    
    
    
    If Encabezado = True Then
        If tenro1 <> 0 Then
            Call imprimirTexto(testrnomb1, ArchExp, 6, False)        'Estrucura 1
            Call imprimirTexto(Sep, ArchExp, 2, True)                'Separador
        End If
        If tenro2 <> 0 Then
           Call imprimirTexto(testrnomb2, ArchExp, 6, False)         'Estrucura 2
           Call imprimirTexto(Sep, ArchExp, 2, True)                 'Separador
        End If
        If tenro3 <> 0 Then
           Call imprimirTexto(testrnomb3, ArchExp, 6, False)         'Estrucura 3
           Call imprimirTexto(Sep, ArchExp, 2, True)                 'Separador
        End If
       lineaEncabezado = "Empleado;Apellido;FecIngEmp;FechaAlta;FechaBaja;Causa;"
       lineaEncabezado = lineaEncabezado & "Estado;Empresa;Centros de Costo;Nº CC;"
       lineaEncabezado = lineaEncabezado & "Puesto Contable;"
    
       Call imprimirTexto(lineaEncabezado, ArchExp, 11, True)        'Encabezado
       
       'Salto de linea
       ArchExp.writeline ""
       
       Encabezado = False
    End If
        
            
    If tenro1 <> 0 Then
        Call imprimirTexto(estrnomb1, ArchExp, 6, False)             'Estrucura 1
        Call imprimirTexto(Sep, ArchExp, 2, True)                    'Separador
    End If
    If tenro2 <> 0 Then
       Call imprimirTexto(estrnomb2, ArchExp, 6, False)              'Estrucura 2
       Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
    End If
    If tenro3 <> 0 Then
       Call imprimirTexto(estrnomb3, ArchExp, 6, False)              'Estrucura 3
       Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
    End If
    
    Call imprimirTexto(Legajo, ArchExp, 11, True)                    'Legajo
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    Call imprimirTexto(Apellido, ArchExp, 2, True)                   'Apellido
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    Call imprimirTexto(FecIngEmp, ArchExp, 2, True)                  'Fecha Ingreso Empresa
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
    Call imprimirTexto(FechaAlta, ArchExp, 2, True)                  'Fecha Alta
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
    Call imprimirTexto(FechaBaja, ArchExp, 2, True)                  'Fecha Baja
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
    Call imprimirTexto(CausaBaja, ArchExp, 2, True)                  'Causa Baja
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    If Estado = -1 Then
        Call imprimirTexto("Activo", ArchExp, 2, True)               'Estado
    Else
        Call imprimirTexto("Inactivo", ArchExp, 2, True)             'Estado
    End If
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    Call imprimirTexto(Empresa, ArchExp, 2, True)                    'Empresa
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    Call imprimirTexto(CentroCosto, ArchExp, 2, True)                'centro costo
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    Call imprimirTexto(NroCC, ArchExp, 2, True)                      'nro cc
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    Call imprimirTexto(PuestoContable, ArchExp, 2, True)             'puesto contable
    Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
    
    'Salto de linea
    ArchExp.writeline ""
    
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    If rsConsult.State = adStateOpen Then rsConsult.Close
    
    Set rs = Nothing
    Set objRs = Nothing
Exit Sub

ME_Local:
    Flog.writeline
    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado "
    StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
    StrEmpl = StrEmpl & " ORDER BY estado "
    
    OpenRecordset StrEmpl, rsEmpl
End Sub

