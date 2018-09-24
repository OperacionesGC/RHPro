Attribute VB_Name = "ExpConcLiq"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion de conceptos liquidados
' Autor      : JMH
' Fecha      : 09/12/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "09/12/2005"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "28/08/2006"
'Global Const UltimaModificacion = " " 'Se agregaron hasta 3 estructuras por confrep y la cuenta contable y el modelo de asiento

'Global Const Version = "1.03"
'Global Const FechaModificacion = "11/10/2006"
'Global Const UltimaModificacion = " " 'Se muestra la cuenta contable (linacuenta. Sin resolver) y el modelo de asiento

Global Const Version = "1.04"
Global Const FechaModificacion = "01/08/2008"
Global Const UltimaModificacion = " "
'   01/08/2008 - Fernando Favre - CUSTOM - Arlei - Se agrego la posibilidad de generar el archivo en .../In-OutPorUsr/<user>/..
'                Se esta manera cada reporte generado no es compartido por el resto de usuarios.
'                El manejo de la seguridad de los directorios queda en manos del administrador de la empresa
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
Global tprocNro As Integer
Global tprocDesc As String
Global proDesc As String
Global concnro As Integer
Global Conccod As String
Global concabr As String
Global tconnro As Integer
Global tconDesc As String
Global concimp As Integer
Global concpuente As Integer
Global fecEstr As String
Global Formato As Integer
Global Modelo As Long
Global TituloRep As String
Global descDesde
Global descHasta
Global FechaHasta
Global FechaDesde
Global ArchExp
Global UsaEncabezado As Integer
Global Encabezado As Boolean

Global IdUser As String
Global Fecha As Date
Global Hora As String


Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion de conceptos liquidados.
' Autor      : JMH
' Fecha      : 09/12/2005
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
    
    Nombre_Arch = PathFLog & "ExportacionConcLiq" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Exportación Conceptos Liquidados : " & Now
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
    StrSql = StrSql & " AND btprcnro = 117"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       IdUser = rs!IdUser
       Fecha = rs!bprcfecha
       Hora = rs!bprchora
       
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)
       
       'Obtengo el periodo desde
       pliqdesde = CLng(ArrParametros(1))
       
       'Obtengo el periodo hasta
       pliqhasta = CLng(ArrParametros(2))
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(3))
       estrnro1 = CInt(ArrParametros(4))
       tenro2 = CInt(ArrParametros(5))
       estrnro2 = CInt(ArrParametros(6))
       tenro3 = CInt(ArrParametros(7))
       estrnro3 = CInt(ArrParametros(8))
       fecEstr = ArrParametros(9)
       
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqdesde
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          FechaDesde = objRs!pliqdesde
          descDesde = objRs!pliqDesc
       Else
          Flog.writeline "No se encontro el periodo desde."
          Exit Sub
       End If
        
       objRs.Close
       
       'Busco el periodo hasta
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqhasta
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          FechaHasta = objRs!pliqhasta
          descHasta = objRs!pliqDesc
       Else
          Flog.writeline "No se encontro el periodo hasta."
          Exit Sub
       End If
        
       objRs.Close
       
       'Busco en el confrep
       Flog.writeline "Buscando valores de tipo de estructura en configuración del reporte."
       StrSql = "SELECT * FROM confrep WHERE repnro = 169 AND conftipo = 'TE'"
       StrSql = StrSql & " ORDER BY confnrocol "
       OpenRecordset StrSql, objRs
       If Not objRs.EOF Then
          tenro1_confrep = objRs!confval
          objRs.MoveNext
       Else
          Flog.writeline "No se encontro ninguna configuración de tipo TE en la configuración del reporte."
       End If
       
       If Not objRs.EOF Then
          tenro2_confrep = objRs!confval
          objRs.MoveNext
       End If
       
       If Not objRs.EOF Then
          tenro3_confrep = objRs!confval
          objRs.MoveNext
       End If
       
       If Not objRs.EOF Then
          Flog.writeline "Se aceptan hasta 3 posibles TE en la configuración del reporte."
       End If
       objRs.Close
       
       NroModelo = 269
    
       'Directorio de exportacion
       StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
       If rs.State = adStateOpen Then rs.Close
       OpenRecordset StrSql, rs
       If Not rs.EOF Then
          Directorio = Trim(rs!sis_dirsalidas)
       End If
     
       StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
       OpenRecordset StrSql, rs_Modelo
       If Not rs_Modelo.EOF Then
          If Not IsNull(rs_Modelo!modarchdefault) Then
             Directorio = Directorio & "PorUsr\" & IdUser & Trim(rs_Modelo!modarchdefault)
          Else
             Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
          End If
       Else
          Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
       End If
                
       'Obtengo los datos del separador
       Sep = rs_Modelo!modseparador
       UsaEncabezado = rs_Modelo!modencab
       
       If UsaEncabezado = -1 Then
          Encabezado = True
       Else
          Encabezado = False
       End If
       
       Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
        
       'Activo el manejador de errores
       On Error Resume Next

       'Archivo para el detalle del Pedido de Pago
       Nombre_Arch = Directorio & "\conceptos_liquidados.csv"
       Set fs = CreateObject("Scripting.FileSystemObject")
       Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
       If Err.Number <> 0 Then
            Flog.writeline "La carpeta Destino no existe. Se creará."
            Set Carpeta = fs.CreateFolder(Directorio)
            Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
       End If
       'desactivo el manejador de errores

       On Error GoTo ME_Main

'       Nombre_Arch = Directorio & "\conceptos_liquidados.csv"
'       Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
'       Set fs = CreateObject("Scripting.FileSystemObject")
'       On Error Resume Next
'       If Err.Number <> 0 Then
'          Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
'          Set Carpeta = fs1.CreateFolder(Directorio)
'       End If
       'desactivo el manejador de errores
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
          
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
          StrSql = "SELECT pliqnro FROM periodo WHERE "
          StrSql = StrSql & " pliqdesde >= " & ConvFecha(FechaDesde)
          StrSql = StrSql & " AND pliqhasta <= " & ConvFecha(FechaHasta)
              
          OpenRecordset StrSql, rsPeriodos
              
          EmpErrores = False
          ternro = rsEmpl!ternro
          orden = rsEmpl!estado
              
          Flog.writeline Espacios(Tabulador * 1) & "Se comienza a procesar los datos"
    
          'Genero una entrada para el empleado por cada periodo
           Do Until rsPeriodos.EOF
              Flog.writeline "Generando datos empleado " & ternro & " para el periodo " & rsPeriodos!pliqNro
              
              Call Generar_Archivo_Con_Liq(listapronro, rsPeriodos!pliqNro, ternro, Sep)
              
              rsPeriodos.MoveNext
           Loop
              
           rsPeriodos.Close
              
           Flog.writeline Espacios(Tabulador * 1) & "Se Terminaron de Procesar los datos"
    
           'Actualizo el estado del proceso
           TiempoAcumulado = GetTickCount
              
           cantRegistros = cantRegistros - 1
              
           StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                 
           objConn.Execute StrSql, , adExecuteNoRecords
              
           rsEmpl.MoveNext
       Loop
       
       ArchExp.Close
       'Si se generaron todos los datos del empleado correctamente lo borro
       If Not EmpErrores Then
          StrSql = " DELETE FROM batch_empleado "
          StrSql = StrSql & " WHERE bpronro = " & NroProceso
          StrSql = StrSql & " AND ternro = " & ternro
        
          objConn.Execute StrSql, , adExecuteNoRecords
       End If
              
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
    GoTo Fin
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



Private Sub Generar_Archivo_Con_Liq(ListaProcesos, pliqNro, ternro, Sep)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion de conceptos liquidados
' Autor      : JMH
' Fecha      : 09/12/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Formato
'------------------------------------------------------------
'Nro Campo                                  Desde   Longitud
'------------------------------------------------------------
'1   Legajo                                 0       0
'2   Apellido                               0       0
'3   Nombre                                 0       0
'4   Estructura 1 del confrep (opcional)    0       0
'5   Estructura 2 del confrep (opcional)    0       0
'6   Estrucutra 3 del confrep (opcional)    0       0
'7   Estructura 1 (opcional)                0       0
'8   Estructura 2 (opcional)                0       0
'9   Estrucutra 3 (opcional)                0       0
'10  Periodo                                0       0
'11  Modelo                                 0       0
'12  Concepto                               0       0
'13  Descripcion                            0       0
'14  Monto                                  0       0
'15  Cantidad                               0       0
'16  Puente                                 0       0
'17  Tipo de Concepto                       0       0
'18  Imprime                                0       0
'19  Cuenta                                 0       0
'20  Modelo de asiento                      0       0

Dim I As Integer
Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim testrnomb1
Dim testrnomb2
Dim testrnomb3
Dim estrnomb1_confrep As String
Dim testrnomb1_confrep As String
Dim estrnomb2_confrep As String
Dim testrnomb2_confrep As String
Dim estrnomb3_confrep As String
Dim testrnomb3_confrep As String
Dim pliqDesc
Dim pliqFecha
Dim proNro
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsCuenta As New ADODB.Recordset
Dim lineaEncabezado As String

    On Error GoTo ME_Local

    estrnomb1 = ""
    estrnomb2 = ""
    estrnomb3 = ""
    proNro = 0

    '------------------------------------------------------------------
    'Controlo si el empleado tiene algun proceso en el periodo
    '------------------------------------------------------------------
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM proceso "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqNro
    StrSql = StrSql & " WHERE empleado = " & ternro
    StrSql = StrSql & "   AND proceso.pliqnro = " & pliqNro
    StrSql = StrSql & "   AND proceso.pronro IN (" & ListaProcesos & ")"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
       'Si el empleado no tiene procesos en el periodo paso al siguiente
       rsConsult.Close
       
       Exit Sub
    End If
    
    rsConsult.Close

    '------------------------------------------------------------------
    'Busco los datos del empleado
    '------------------------------------------------------------------
    StrSql = " SELECT empleg,terape,terape2,ternom,ternom2 "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE ternro= " & ternro
    
    Flog.writeline "Buscando datos del empleado"
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       nombre = rsConsult!ternom
       If IsNull(rsConsult!ternom2) Then
          nombre2 = ""
       Else
          nombre2 = rsConsult!ternom2
       End If
       apellido = rsConsult!terape
       If IsNull(rsConsult!terape2) Then
          apellido2 = ""
       Else
          apellido2 = rsConsult!terape2
       End If
       Legajo = rsConsult!empleg
    Else
       Flog.writeline "Error al obtener los datos del empleado"
    '   GoTo MError
    End If
    
    rsConsult.Close

    '------------------------------------------------------------------
    'Busco los datos del tipos de estructura 1
    '------------------------------------------------------------------
    
    '---LOG---
    Flog.writeline "Buscando datos estructura 1"
    
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
    
    '---LOG---
    Flog.writeline "Buscando datos estructura 2"
    
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
    
    '---LOG---
    Flog.writeline "Buscando datos estructura 3"
    
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
    'Busco los datos del tipo de estructura 1 del confrep si existe
    '------------------------------------------------------------------
    '---LOG---
    Flog.writeline "Buscando datos estructura 1 definida en el confrep: " & tenro1_confrep
    estrnomb1_confrep = ""
    If tenro1_confrep <> 0 Then
        
        StrSql = " SELECT estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro1_confrep
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           estrnomb1_confrep = rsConsult!estrdabr
           testrnomb1_confrep = rsConsult!tedabr
        End If
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del tipo de estructura 2 del confrep si existe
    '------------------------------------------------------------------
    '---LOG---
    Flog.writeline "Buscando datos estructura 2 definida en el confrep: " & tenro2_confrep
    estrnomb2_confrep = ""
    If tenro2_confrep <> 0 Then
        
        StrSql = " SELECT estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro2_confrep
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           estrnomb2_confrep = rsConsult!estrdabr
           testrnomb2_confrep = rsConsult!tedabr
        End If
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del tipo de estructura 3 del confrep si existe
    '------------------------------------------------------------------
    '---LOG---
    Flog.writeline "Buscando datos estructura 3 definida en el confrep: " & tenro3_confrep
    estrnomb3_confrep = ""
    If tenro3_confrep <> 0 Then
        
        StrSql = " SELECT estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro3_confrep
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           estrnomb3_confrep = rsConsult!estrdabr
           testrnomb3_confrep = rsConsult!tedabr
        End If
    End If
    
    
    '------------------------------------------------------------------
    'Busco los datos del periodo
    '------------------------------------------------------------------
    StrSql = " SELECT * FROM periodo WHERE pliqnro = " & pliqNro
    
    OpenRecordset StrSql, rsConsult
    
    pliqDesc = ""
    If Not rsConsult.EOF Then
       pliqDesc = rsConsult!pliqDesc
       pliqFecha = rsConsult!pliqdesde
    End If
    
    rsConsult.Close
    
    '------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores
    '------------------------------------------------------------------
    
    StrSql = " SELECT detliq.concnro, sum(detliq.dlicant) AS cant, sum(detliq.dlimonto) AS monto,  "
    StrSql = StrSql & " concepto.concimp, concepto.concpuente, concepto.tconnro, tcondesc, prodesc, "
    StrSql = StrSql & " tipoproc.tprocnro, tprocdesc, conccod, concabr, proceso.pronro, proceso.pliqnro, pliqdesc, concorden  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo   ON periodo.pliqnro = " & pliqNro
    StrSql = StrSql & " INNER JOIN proceso   ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND cabliq.pronro IN (" & ListaProcesos & ") "
    StrSql = StrSql & " INNER JOIN detliq    ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " INNER JOIN concepto  ON concepto.concnro = detliq.concnro "
    StrSql = StrSql & " INNER JOIN tipconcep ON tipconcep.tconnro = concepto.tconnro "
    StrSql = StrSql & " INNER JOIN tipoproc  ON tipoproc.tprocnro = proceso.tprocnro "
    StrSql = StrSql & " GROUP BY detliq.concnro, concepto.concimp, concepto.concpuente, "
    StrSql = StrSql & " concepto.tconNro , tconDesc, proDesc, tipoproc.tprocnro, tprocdesc, conccod, "
    StrSql = StrSql & " concabr, proceso.pronro, proceso.pliqnro, pliqdesc, concorden  "
    StrSql = StrSql & " ORDER BY concepto.tconnro, concorden  "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF

        If Encabezado = True Then
           lineaEncabezado = "Legajo;Apellido;Nombre;"
           If testrnomb1_confrep <> "" Then
              lineaEncabezado = lineaEncabezado & testrnomb1_confrep & ";"
           End If
           If testrnomb2_confrep <> "" Then
              lineaEncabezado = lineaEncabezado & testrnomb2_confrep & ";"
           End If
           If testrnomb3_confrep <> "" Then
              lineaEncabezado = lineaEncabezado & testrnomb3_confrep & ";"
           End If
           
           If testrnomb1 <> "" Then
              lineaEncabezado = lineaEncabezado & testrnomb1 & ";"
           End If
           If testrnomb2 <> "" Then
              lineaEncabezado = lineaEncabezado & testrnomb2 & ";"
           End If
           If testrnomb3 <> "" Then
              lineaEncabezado = lineaEncabezado & testrnomb3 & ";"
           End If
           
           lineaEncabezado = lineaEncabezado & "Periodo;Proceso;Modelo;Concepto;"
           lineaEncabezado = lineaEncabezado & "Descripción;Monto;Cantidad;Puente;"
           lineaEncabezado = lineaEncabezado & "Tipo de Concepto;Imprime;Cuenta;Modelo;"
        
           Call imprimirTexto(lineaEncabezado, ArchExp, 11, True)        'Encabezado
           
           'Salto de linea
           ArchExp.writeline ""
           
           Encabezado = False
        End If
        
        StrSql = "SELECT mod_linea.linacuenta, mod_asiento.masidesc FROM mod_linea "
        StrSql = StrSql & " INNER JOIN mod_asiento ON mod_linea.masinro = mod_asiento.masinro "
        StrSql = StrSql & " INNER JOIN asi_con ON mod_linea.masinro = asi_con.masinro "
        StrSql = StrSql & " AND mod_linea.linaorden = asi_con.linaorden AND asi_con.concnro = " & rsConsult!concnro

        'StrSql = "SELECT linea_asi.cuenta, mod_asiento.masidesc FROM linea_asi "
        'StrSql = StrSql & " INNER JOIN mod_asiento ON linea_asi.masinro = mod_asiento.masinro "
        'StrSql = StrSql & " INNER JOIN asi_con ON linea_asi.masinro = asi_con.masinro AND "
        'StrSql = StrSql & " linea_asi.linea = asi_con.linaorden AND asi_con.concnro = " & rsConsult!concnro
        'StrSql = StrSql & " INNER JOIN proc_vol_pl ON linea_asi.vol_cod = proc_vol_pl.vol_cod AND proc_vol_pl.pronro = " & rsConsult!proNro
        OpenRecordset StrSql, rsCuenta
        
        If rsCuenta.EOF Then
            Call imprimirTexto(Legajo, ArchExp, 11, True)                    'Legajo
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(apellido & " " & apellido2, ArchExp, 2, True) 'Apellido
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(nombre & " " & nombre2, ArchExp, 2, True)     'Nombre
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
            If estrnomb1_confrep <> "" Then
               Call imprimirTexto(estrnomb1_confrep, ArchExp, 6, False)      'Estructura 1 del confrep
               Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
            End If
            If estrnomb2_confrep <> "" Then
               Call imprimirTexto(estrnomb2_confrep, ArchExp, 6, False)      'Estructura 2 del confrep
               Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
            End If
            If estrnomb3_confrep <> "" Then
               Call imprimirTexto(estrnomb3_confrep, ArchExp, 6, False)      'Estructura 3 del confrep
               Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
            End If
            
            If estrnomb1 <> "" Then
               Call imprimirTexto(estrnomb1, ArchExp, 6, False)              'Estrucura 1
               Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
            End If
            If estrnomb2 <> "" Then
               Call imprimirTexto(estrnomb2, ArchExp, 6, False)              'Estrucura 2
               Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
            End If
            If estrnomb3 <> "" Then
               Call imprimirTexto(estrnomb3, ArchExp, 6, False)              'Estrucura 3
               Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
            End If
            Call imprimirTexto(rsConsult!pliqDesc, ArchExp, 6, False)        'Periodo
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(rsConsult!proDesc, ArchExp, 6, False)         'Proceso
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(rsConsult!tprocDesc, ArchExp, 6, False)       'Modelo
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(rsConsult!Conccod, ArchExp, 6, False)         'Cod. Concepto
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(rsConsult!concabr, ArchExp, 6, False)         'Desc. Concepto
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTextoConCeros(rsConsult!Monto, ArchExp, 2, False)  'Monto
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTextoConCeros(rsConsult!cant, ArchExp, 2, False)   'Cantidad
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            If rsConsult!concpuente = "-1" Then
               Call imprimirTexto("SI", ArchExp, 6, False)                   'Puente
            Else
               Call imprimirTexto("NO", ArchExp, 6, False)                   'Puente
            End If
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            Call imprimirTexto(rsConsult!tconDesc, ArchExp, 6, False)        'Desc. Concepto
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            If rsConsult!concimp = "-1" Then
               Call imprimirTexto("SI", ArchExp, 6, False)                   'Imprime
            Else
               Call imprimirTexto("NO", ArchExp, 6, False)                   'Imprime
            End If
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
            Call imprimirTexto("", ArchExp, 6, False)                        'Cuenta
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
            Call imprimirTexto("", ArchExp, 6, False)                        'Modelo de asiento
            Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
            
            'Salto de linea
            ArchExp.writeline ""
        
        Else
            Do Until rsCuenta.EOF
                Call imprimirTexto(Legajo, ArchExp, 11, True)                    'Legajo
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(apellido & " " & apellido2, ArchExp, 2, True) 'Apellido
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(nombre & " " & nombre2, ArchExp, 2, True)     'Nombre
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                
                If estrnomb1_confrep <> "" Then
                   Call imprimirTexto(estrnomb1_confrep, ArchExp, 6, False)      'Estructura 1 del confrep
                   Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
                End If
                If estrnomb2_confrep <> "" Then
                   Call imprimirTexto(estrnomb2_confrep, ArchExp, 6, False)      'Estructura 2 del confrep
                   Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
                End If
                If estrnomb3_confrep <> "" Then
                   Call imprimirTexto(estrnomb3_confrep, ArchExp, 6, False)      'Estructura 3 del confrep
                   Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
                End If
                
                If estrnomb1 <> "" Then
                   Call imprimirTexto(estrnomb1, ArchExp, 6, False)              'Estrucura 1
                   Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
                End If
                If estrnomb2 <> "" Then
                   Call imprimirTexto(estrnomb2, ArchExp, 6, False)              'Estrucura 2
                   Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
                End If
                If estrnomb3 <> "" Then
                   Call imprimirTexto(estrnomb3, ArchExp, 6, False)              'Estrucura 3
                   Call imprimirTexto(Sep, ArchExp, 2, True)                     'Separador
                End If
                Call imprimirTexto(rsConsult!pliqDesc, ArchExp, 6, False)        'Periodo
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(rsConsult!proDesc, ArchExp, 6, False)         'Proceso
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(rsConsult!tprocDesc, ArchExp, 6, False)       'Modelo
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(rsConsult!Conccod, ArchExp, 6, False)         'Cod. Concepto
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(rsConsult!concabr, ArchExp, 6, False)         'Desc. Concepto
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTextoConCeros(rsConsult!Monto, ArchExp, 2, False)   'Monto
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTextoConCeros(rsConsult!cant, ArchExp, 2, False)    'Cantidad
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                If rsConsult!concpuente = "-1" Then
                   Call imprimirTexto("SI", ArchExp, 6, False)                   'Puente
                Else
                   Call imprimirTexto("NO", ArchExp, 6, False)                   'Puente
                End If
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                Call imprimirTexto(rsConsult!tconDesc, ArchExp, 6, False)        'Desc. Concepto
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                If rsConsult!concimp = "-1" Then
                   Call imprimirTexto("SI", ArchExp, 6, False)                   'Imprime
                Else
                   Call imprimirTexto("NO", ArchExp, 6, False)                   'Imprime
                End If
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                
                Call imprimirTexto(rsCuenta!linacuenta, ArchExp, 6, False)       'Cuenta
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                
                Call imprimirTexto(rsCuenta!masidesc, ArchExp, 6, False)         'Modelo de asiento
                Call imprimirTexto(Sep, ArchExp, 2, True)                        'Separador
                
                'Salto de linea
                ArchExp.writeline ""
                              
                rsCuenta.MoveNext
            Loop
        End If
        
        I = I + 1
        rsConsult.MoveNext
    Loop
    rsConsult.Close
    
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    If rsConsult.State = adStateOpen Then rsConsult.Close
    
    Set rs = Nothing
    Set objRs = Nothing
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    HuboErrores = True
End Sub
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    StrEmpl = StrEmpl & " ORDER BY ternro "
    
    OpenRecordset StrEmpl, rsEmpl
End Sub

