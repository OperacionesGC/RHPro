Attribute VB_Name = "ExpDirection"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion de conceptos liquidados
' Autor      : JMH
' Fecha      : 09/12/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "20/12/2005"
Global Const UltimaModificacion = " " 'Version Inicial

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

Global errorConfrep As Boolean

Global TipoCols4(200)
Global CodCols4(200)
Global TipoCols5(200)
Global CodCols5(200)

Global mes1 As String
Global mesPorc1 As String
Global mes2 As String
Global mesPorc2 As String
Global mes3 As String
Global mes4 As String
Global mesPorc3 As String
Global mes5 As String
Global mesPorc4 As String
Global mes6 As String


Global mesPeriodo As Integer
Global anioPeriodo As Integer
Global mesAnterior1 As Integer
Global mesAnterior2 As Integer
Global anioAnterior1 As Integer
Global anioAnterior2 As Integer

Global cantColumna4
Global cantColumna5

Global estrnomb1
Global estrnomb2
Global estrnomb3
Global testrnomb1
Global testrnomb2
Global testrnomb3

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
Dim SepDec As String
Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim param
Dim listaModelos
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
    
    Nombre_Arch = PathFLog & "ExportacionDirection" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Exportación Direction : " & Now
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
    StrSql = StrSql & " AND btprcnro = 118"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       'Obtengo la lista de procesos
       listaModelos = ArrParametros(0)
       
       'Obtengo el periodo desde
       pliqdesde = CLng(ArrParametros(1))
       
       'Obtengo el periodo hasta
       'pliqhasta = CLng(ArrParametros(2))
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(2))
       estrnro1 = CInt(ArrParametros(3))
       tenro2 = CInt(ArrParametros(4))
       estrnro2 = CInt(ArrParametros(5))
       tenro3 = CInt(ArrParametros(6))
       estrnro3 = CInt(ArrParametros(7))
       fecEstr = ArrParametros(8)
       
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
       
       NroModelo = 270
    
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
             Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
          Else
             Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
          End If
       Else
          Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
       End If
                
       'Obtengo los datos del separador
       Sep = rs_Modelo!modseparador
       SepDec = rs_Modelo!modsepdec
       UsaEncabezado = rs_Modelo!modencab
       
       If UsaEncabezado = -1 Then
          Encabezado = True
       Else
          Encabezado = False
       End If
       
       Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
        
       Nombre_Arch = Directorio & "\direction" & "-" & NroProceso & ".csv"
       Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
       Set fs = CreateObject("Scripting.FileSystemObject")
       On Error Resume Next
       If Err.Number <> 0 Then
          Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
          Set Carpeta = fs1.CreateFolder(Directorio)
       End If
       'desactivo el manejador de errores
       Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
      
       'Cargo la configuracion del reporte
       Call CargarConfiguracionReporte
       If errorConfrep = True Then
          Exit Sub
       End If
       
       'Obtengo los empleados sobre los que tengo que generar los recibos
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Genero el encabezado
       Call GenerarEncabezado(pliqdesde)
       
       'seteo de las variables de progreso
       Progreso = 0
       cantRegistros = rsEmpl.RecordCount
       totalEmpleados = rsEmpl.RecordCount
       If cantRegistros = 0 Then
          cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
       End If
       IncPorc = (99 / cantRegistros)
          
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
       
          ternro = rsEmpl!ternro
                        
          Flog.writeline "Generando datos empleado " & ternro & " para el periodo " & pliqhasta
              
          'Call Generar_Archivo_Direction(listaModelos, pliqhasta, ternro, Sep, SepDec)
          Call Cargar_Tabla_Direction(listaModelos, pliqdesde, ternro)
              
            
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
       
       Call Generar_Archivo_Direction(Sep, SepDec)
       
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



Private Sub Generar_Archivo_Direction(Sep, SepDec)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion de direction
' Autor      : JMH
' Fecha      : 09/12/2005
' Ultima Mod :
' Descripcion:

Dim i As Integer
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
Dim pliqDesc
Dim pliqFecha
Dim proNro
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim lineaEncabezado As String
Dim lineaDetalle As String
Dim mesPeriodo As Integer
Dim anioPeriodo As Integer
Dim mesAnterior1 As Integer
Dim mesAnterior2 As Integer
Dim anioAnterior1 As Integer
Dim anioAnterior2 As Integer
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim porcentaje1 As Double
Dim porcentaje2 As Double
Dim valor4 As Double
Dim valor5 As Double
Dim valor6 As Double
Dim porcentaje3 As Double
Dim porcentaje4 As Double
Dim convenio As String
Dim apeynom As String

    On Error GoTo ME_Local

    estrnomb1 = ""
    estrnomb2 = ""
    estrnomb3 = ""
    proNro = 0
    
    
    '------------------------------------------------------------------
    'Busco los datos guardados del encabezado
    '------------------------------------------------------------------
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM rep_direction "
    StrSql = StrSql & " WHERE bpronro= " & NroProceso
        
    Flog.writeline "Buscando datos de la tabla"
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       testrnomb1 = rsConsult!tedabr1
       testrnomb2 = rsConsult!tedabr2
       testrnomb3 = rsConsult!tedabr3
       mes1 = rsConsult!mes_mon_1
       mesPorc1 = rsConsult!mes_porc_1
       mes2 = rsConsult!mes_mon_2
       mesPorc2 = rsConsult!mes_porc_2
       mes3 = rsConsult!mes_mon_3
       mes4 = rsConsult!mes_mon_4
       mesPorc3 = rsConsult!mes_porc_3
       mes5 = rsConsult!mes_mon_5
       mesPorc4 = rsConsult!mes_porc_4
       mes6 = rsConsult!mes_mon_6
    End If
    rsConsult.Close
    
    '------------------------------------------------------------------
    'Busco los datos guardados de los empleados en la tabla
    '------------------------------------------------------------------
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM rep_direc_det "
    StrSql = StrSql & " WHERE bpronro= " & NroProceso
    StrSql = StrSql & " ORDER BY convenio, mon_porc_2 DESC "
    
    Flog.writeline "Buscando datos de la tabla"
           
    OpenRecordset StrSql, rsConsult
    
    Do While Not rsConsult.EOF
       Legajo = rsConsult!Legajo
       apeynom = rsConsult!apeynom
       estrnomb1 = rsConsult!estrdabr1
       estrnomb2 = rsConsult!estrdabr2
       estrnomb3 = rsConsult!estrdabr3
       valor1 = rsConsult!mon_mes_1
       porcentaje1 = rsConsult!mon_porc_1
       valor2 = rsConsult!mon_mes_2
       porcentaje2 = rsConsult!mon_porc_2
       valor3 = rsConsult!mon_mes_3
       valor4 = rsConsult!mon_mes_4
       porcentaje3 = rsConsult!mon_porc_4
       valor5 = rsConsult!mon_mes_5
       porcentaje4 = rsConsult!mon_porc_5
       valor6 = rsConsult!mon_mes_6
       
       If Encabezado = True Then
    
          'Imprime el encabezado
          lineaEncabezado = "MENSUALES" & Sep & Sep
          If testrnomb1 <> "" Then
             lineaEncabezado = lineaEncabezado & Sep
          End If
          If testrnomb2 <> "" Then
             lineaEncabezado = lineaEncabezado & Sep
          End If
          If testrnomb3 <> "" Then
             lineaEncabezado = lineaEncabezado & Sep
          End If
          lineaEncabezado = lineaEncabezado & Sep & Sep & "BASICO" & Sep & Sep
          lineaEncabezado = lineaEncabezado & Sep & Sep & "VARIABLE" & Sep & Sep
       
          Call imprimirTexto(lineaEncabezado, ArchExp, 11, True)        'Encabezado
       
          'Salto de linea
          ArchExp.writeline ""
       
          lineaEncabezado = "Legajo" & Sep & "Apellido y Nombre" & Sep
          If testrnomb1 <> "" Then
             lineaEncabezado = lineaEncabezado & testrnomb1 & Sep
          End If
          If testrnomb2 <> "" Then
             lineaEncabezado = lineaEncabezado & testrnomb2 & Sep
          End If
          If testrnomb3 <> "" Then
             lineaEncabezado = lineaEncabezado & testrnomb3 & Sep
          End If
                 
          'PRIMER MES
          lineaEncabezado = lineaEncabezado & mes1 & Sep
          lineaEncabezado = lineaEncabezado & mesPorc1 & Sep
       
          'SEGUNDO MES
          lineaEncabezado = lineaEncabezado & mes2 & Sep
          lineaEncabezado = lineaEncabezado & mesPorc2 & Sep
       
          'TERCER MES
          lineaEncabezado = lineaEncabezado & mes3 & Sep
                 
          'PRIMER MES
          lineaEncabezado = lineaEncabezado & mes4 & Sep
          lineaEncabezado = lineaEncabezado & mesPorc3 & Sep
       
          'SEGUNDO MES
          lineaEncabezado = lineaEncabezado & mes5 & Sep
          lineaEncabezado = lineaEncabezado & mesPorc4 & Sep
       
          'TERCER MES
          lineaEncabezado = lineaEncabezado & mes6 & Sep
       
          Call imprimirTexto(lineaEncabezado, ArchExp, 11, True)        'Encabezado
           
          'Salto de linea
          ArchExp.writeline ""
           
          Encabezado = False
       End If
    
       'Imprime el detalle
       lineaDetalle = Legajo & Sep & apeynom & Sep
       If testrnomb1 <> "" Then
          lineaDetalle = lineaDetalle & estrnomb1 & Sep
       End If
       If testrnomb2 <> "" Then
          lineaDetalle = lineaDetalle & estrnomb2 & Sep
       End If
       If testrnomb3 <> "" Then
          lineaDetalle = lineaDetalle & estrnomb3 & Sep
       End If
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(valor1, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(porcentaje1, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(valor2, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(porcentaje2, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(valor3, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(valor4, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(porcentaje3, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(valor5, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(porcentaje4, 2), ",", ""), ".", SepDec) & Sep
       lineaDetalle = lineaDetalle & Replace(Replace(FormatNumber(valor6, 2), ",", ""), ".", SepDec) & Sep
    
       Call imprimirTexto(lineaDetalle, ArchExp, 11, True)        'Detalle
    
       'Salto de linea
       ArchExp.writeline ""
                      
       rsConsult.MoveNext
       
    Loop
    
    rsConsult.Close
    
    '-------------------------------------------------------------------------------
    'Borro los datos en la BD
    '-------------------------------------------------------------------------------

    StrSql = " DELETE FROM rep_direction WHERE bpronro = " & NroProceso
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " DELETE FROM rep_direc_det WHERE bpronro = " & NroProceso
    
    objConn.Execute StrSql, , adExecuteNoRecords
        
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
End Sub

Private Sub Cargar_Tabla_Direction(listaModelos, pliqNro, ternro)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga se almacenar en la tabla direction
' Autor      : JMH
' Fecha      : 09/12/2005
' Ultima Mod :
' Descripcion:

Dim i As Integer
Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String

Dim pliqDesc
Dim pliqFecha
Dim proNro
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim lineaEncabezado As String
Dim lineaDetalle As String
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim porcentaje1 As Double
Dim porcentaje2 As Double
Dim valor4 As Double
Dim valor5 As Double
Dim valor6 As Double
Dim porcentaje3 As Double
Dim porcentaje4 As Double
Dim convenio As String
Dim apeynom As String

    On Error GoTo ME_Local

    estrnomb1 = ""
    estrnomb2 = ""
    estrnomb3 = ""
    proNro = 0
    
    '------------------------------------------------------------------
    'Busco los datos del empleado
    '------------------------------------------------------------------
    StrSql = " SELECT empleg,terape,terape2,ternom,ternom2 "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE ternro= " & ternro
    
    Flog.writeline "Buscando datos del empleado"
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       apeynom = rsConsult!terape
       If Not IsNull(rsConsult!terape2) Then
          apeynom = apeynom & " " & rsConsult!terape2
       End If
       'apeynom = apeynom & ", " & rsConsult!ternom
       apeynom = apeynom & " " & rsConsult!ternom
       If Not IsNull(rsConsult!ternom2) Then
          apeynom = apeynom & " " & rsConsult!ternom2
       End If
       
       Legajo = rsConsult!empleg
    Else
       Flog.writeline "Error al obtener los datos del empleado"
    '   GoTo MError
    End If
    
    rsConsult.Close
    
    '------------------------------------------------------------------
    'Busco el convenio del empleado
    '------------------------------------------------------------------
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
    StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 19 "
    StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"

    OpenRecordset StrSql, rsConsult
        
    If Not rsConsult.EOF Then
       convenio = rsConsult!estrdabr
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
           'testrnomb1 = rsConsult!tedabr
        End If
        
        rsConsult.Close
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del tipos de estructura 2
    '------------------------------------------------------------------
    
    '---LOG---
    Flog.writeline "Buscando datos estructura 2"
    
    If tenro2 <> 0 Then
        
        StrSql = " SELECT tedabr "
        StrSql = StrSql & " FROM tipoestructura  "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro2
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           'estrnomb2 = rsConsult!estrdabr
           testrnomb2 = rsConsult!tedabr
        End If
        
        rsConsult.Close
        
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
           'testrnomb2 = rsConsult!tedabr
        End If
        rsConsult.Close
    End If
    
    '------------------------------------------------------------------
    'Busco los datos del tipos de estructura 3
    '------------------------------------------------------------------
    
    '---LOG---
    Flog.writeline "Buscando datos estructura 3"
    
    If tenro3 <> 0 Then
        
        StrSql = " SELECT tedabr "
        StrSql = StrSql & " FROM tipoestructura  "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro3
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           'estrnomb2 = rsConsult!estrdabr
           testrnomb3 = rsConsult!tedabr
        End If
        
        rsConsult.Close
        
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
           'testrnomb3 = rsConsult!tedabr
        End If
        rsConsult.Close
    End If
    
    
    '------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores - COLUMNA 4
    '------------------------------------------------------------------

    StrSql = " SELECT 'CO', sum(detliq.dlimonto) AS monto  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
  
    For i = 1 To cantColumna4
      If TipoCols4(i) = "CO" Then
         StrSql = StrSql & " OR detliq.concnro = " & CodCols4(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY pliqnro "
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor1 = rsConsult!Monto
    Else
       valor1 = 0
    End If
    rsConsult.Close
    
    'StrSql = StrSql & " UNION "
    
    StrSql = " SELECT 'AC', sum(acu_liq.almonto) AS monto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
    
    For i = 1 To cantColumna4
      If TipoCols4(i) = "AC" Then
         StrSql = StrSql & " OR acu_liq.acunro = " & CodCols4(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY pliqnro "

    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor1 = valor1 + rsConsult!Monto
    End If
    
    rsConsult.Close
    
    
    '---------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores de un mes anterior
    '---------------------------------------------------------------------
    
    StrSql = " SELECT 'CO', sum(detliq.dlimonto) AS monto  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior1 & " AND periodo.pliqanio = " & anioAnterior1
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
  
    For i = 1 To cantColumna4
      If TipoCols4(i) = "CO" Then
         StrSql = StrSql & " OR detliq.concnro = " & CodCols4(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor2 = rsConsult!Monto
       'porcentaje1 = ((valor1 - valor2) * 100) / valor2
    Else
       valor2 = 0
    End If
    rsConsult.Close
    
    'StrSql = StrSql & " UNION "
    
    StrSql = " SELECT 'AC', sum(acu_liq.almonto) AS monto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior1 & " AND periodo.pliqanio = " & anioAnterior1
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
    
    For i = 1 To cantColumna4
      If TipoCols4(i) = "AC" Then
         StrSql = StrSql & " OR acu_liq.acunro = " & CodCols4(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "

    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor2 = valor2 + rsConsult!Monto
       If valor2 <> 0 Then
          porcentaje1 = ((valor1 - valor2) * 100) / valor2
       Else
          porcentaje1 = valor1 * 100
       End If
    Else
       If valor2 <> 0 Then
          porcentaje1 = ((valor1 - valor2) * 100) / valor2
       Else
          porcentaje1 = valor1 * 100
       End If
    End If
    rsConsult.Close
    
    
    '---------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores de dos mes anterior
    '---------------------------------------------------------------------
    
    StrSql = " SELECT 'CO', sum(detliq.dlimonto) AS monto  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior2 & " AND periodo.pliqanio = " & anioAnterior2
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
  
    For i = 1 To cantColumna4
      If TipoCols4(i) = "CO" Then
         StrSql = StrSql & " OR detliq.concnro = " & CodCols4(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor3 = rsConsult!Monto
       'porcentaje2 = ((valor2 - valor3) * 100) / valor3
    Else
       valor3 = 0
       'porcentaje2 = (valor2 * 100)
    End If
    rsConsult.Close
    
    'StrSql = StrSql & " UNION "
    
    StrSql = " SELECT 'AC', sum(acu_liq.almonto) AS monto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior2 & " AND periodo.pliqanio = " & anioAnterior2
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
    
    For i = 1 To cantColumna4
      If TipoCols4(i) = "AC" Then
         StrSql = StrSql & " OR acu_liq.acunro = " & CodCols4(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "

    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor3 = valor3 + rsConsult!Monto
       If valor3 <> 0 Then
          porcentaje2 = ((valor2 - valor3) * 100) / valor3
       Else
          porcentaje2 = (valor2 * 100)
       End If
    Else
       If valor3 <> 0 Then
          porcentaje2 = ((valor2 - valor3) * 100) / valor3
       Else
          porcentaje2 = (valor2 * 100)
       End If
    End If
    rsConsult.Close
    
    '------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores - COLUMNA 5
    '------------------------------------------------------------------

    StrSql = " SELECT 'CO', sum(detliq.dlimonto) AS monto  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
  
    For i = 1 To cantColumna5
      If TipoCols5(i) = "CO" Then
         StrSql = StrSql & " OR detliq.concnro = " & CodCols5(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY pliqnro "
    
    'StrSql = StrSql & " UNION "
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor4 = rsConsult!Monto
    Else
       valor4 = 0
    End If
    
    rsConsult.Close
    
    StrSql = " SELECT 'AC', sum(acu_liq.almonto) AS monto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
    
    For i = 1 To cantColumna5
      If TipoCols5(i) = "AC" Then
         StrSql = StrSql & " OR acu_liq.acunro = " & CodCols5(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY pliqnro "

    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor4 = valor4 + rsConsult!Monto
    End If
    
    rsConsult.Close
    
    
    '---------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores de un mes anterior
    '---------------------------------------------------------------------
    
    StrSql = " SELECT 'CO', sum(detliq.dlimonto) AS monto  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior1 & " AND periodo.pliqanio = " & anioAnterior1
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
  
    For i = 1 To cantColumna5
      If TipoCols5(i) = "CO" Then
         StrSql = StrSql & " OR detliq.concnro = " & CodCols5(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor5 = rsConsult!Monto
       'porcentaje3 = ((valor4 - valor5) * 100) / valor5
    Else
       valor5 = 0
    End If
    rsConsult.Close
    
    'StrSql = StrSql & " UNION "
    
    StrSql = " SELECT 'AC', sum(acu_liq.almonto) AS monto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior1 & " AND periodo.pliqanio = " & anioAnterior1
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
    
    For i = 1 To cantColumna5
      If TipoCols5(i) = "AC" Then
         StrSql = StrSql & " OR acu_liq.acunro = " & CodCols5(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "

    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor5 = valor5 + rsConsult!Monto
       If valor5 <> 0 Then
          porcentaje3 = ((valor4 - valor5) * 100) / valor5
       Else
          porcentaje3 = valor4 * 100
       End If
    Else
       If valor5 <> 0 Then
          porcentaje3 = ((valor4 - valor5) * 100) / valor5
       Else
          porcentaje3 = valor4 * 100
       End If
    End If
    rsConsult.Close
    
    
    '---------------------------------------------------------------------
    'Busco los valores de los conceptos y acumuladores de dos mes anterior
    '---------------------------------------------------------------------
    
    StrSql = " SELECT 'CO', sum(detliq.dlimonto) AS monto  "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior2 & " AND periodo.pliqanio = " & anioAnterior2
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
  
    For i = 1 To cantColumna5
      If TipoCols5(i) = "CO" Then
         StrSql = StrSql & " OR detliq.concnro = " & CodCols5(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor6 = rsConsult!Monto
       'porcentaje4 = ((valor5 - valor6) * 100) / valor6
    Else
       valor6 = 0
       'porcentaje4 = (valor5 * 100)
    End If
    rsConsult.Close
    
    'StrSql = StrSql & " UNION "
    
    StrSql = " SELECT 'AC', sum(acu_liq.almonto) AS monto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN periodo  ON periodo.pliqmes = " & mesAnterior2 & " AND periodo.pliqanio = " & anioAnterior2
    StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = periodo.pliqnro AND proceso.tprocnro IN (" & listaModelos & ") "
    StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & ternro & " AND cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " AND ( 1=0 "
    
    For i = 1 To cantColumna5
      If TipoCols5(i) = "AC" Then
         StrSql = StrSql & " OR acu_liq.acunro = " & CodCols5(i)
      End If
    Next
    
    StrSql = StrSql & " ) "
    StrSql = StrSql & " GROUP BY periodo.pliqnro "

    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       valor6 = valor6 + rsConsult!Monto
       If valor6 <> 0 Then
          porcentaje4 = ((valor5 - valor6) * 100) / valor6
       Else
          porcentaje4 = ((valor5 - valor6) * 100)
       End If
    Else
       If valor6 <> 0 Then
          porcentaje4 = ((valor5 - valor6) * 100) / valor6
       Else
          porcentaje4 = (valor5 * 100)
       End If
    End If
    rsConsult.Close
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------

    StrSql = " INSERT INTO rep_direc_det ( "
    StrSql = StrSql & " bpronro , legajo, ternro, apeynom, convenio, "
    StrSql = StrSql & " estrdabr1, estrdabr2, estrdabr3, mon_mes_1, "
    StrSql = StrSql & " mon_porc_1 , mon_mes_2, mon_porc_2, mon_mes_3, "
    StrSql = StrSql & " mon_mes_4 , mon_porc_4, mon_mes_5, mon_porc_5, "
    StrSql = StrSql & " mon_mes_6 ) VALUES ( "
    StrSql = StrSql & NroProceso & ","
    StrSql = StrSql & Legajo & ","
    StrSql = StrSql & ternro & ","
    StrSql = StrSql & "'" & apeynom & "','"
    StrSql = StrSql & convenio & "','"
    StrSql = StrSql & estrnomb1 & "','"
    StrSql = StrSql & estrnomb2 & "','"
    StrSql = StrSql & estrnomb3 & "',"
    StrSql = StrSql & numberForSQL(valor3) & ","
    StrSql = StrSql & numberForSQL(Fix(porcentaje2)) & ","
    StrSql = StrSql & numberForSQL(valor2) & ","
    StrSql = StrSql & numberForSQL(Fix(porcentaje1)) & ","
    StrSql = StrSql & numberForSQL(valor1) & ","
    StrSql = StrSql & numberForSQL(valor6) & ","
    StrSql = StrSql & numberForSQL(Fix(porcentaje4)) & ","
    StrSql = StrSql & numberForSQL(valor5) & ","
    StrSql = StrSql & numberForSQL(Fix(porcentaje3)) & ","
    StrSql = StrSql & numberForSQL(valor4) & ")"

    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub

Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String


    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    StrEmpl = StrEmpl & " ORDER BY ternro "
    
    OpenRecordset StrEmpl, rsEmpl
End Sub
Sub CargarConfiguracionReporte()

    Dim objRs As New ADODB.Recordset
    Dim objRs2 As New ADODB.Recordset
    Dim StrSql As String
    Dim i
    Dim columnaActual
    Dim Nro_col
    Dim Valor As Long
    
    tenro1 = 0
    tenro2 = 0
    tenro3 = 0
    
    StrSql = " SELECT * FROM confrep WHERE confrep.repnro= 152 "
    StrSql = StrSql & "ORDER BY confnrocol "
    
    OpenRecordset StrSql, objRs
    
    If objRs.EOF Then
       Flog.writeline Espacios(Tabulador * 1) & "No se encontró la Configuración del Reporte 152 "
       errorConfrep = True
       Exit Sub
    Else
        objRs.Close
        errorConfrep = False
        
        StrSql = " SELECT * FROM confrep WHERE confrep.repnro= 152 AND confnrocol <= 3 "
        StrSql = StrSql & "ORDER BY confnrocol "
    
        OpenRecordset StrSql, objRs
        Do Until objRs.EOF
           
           columnaActual = CLng(objRs!confnrocol)
           
           If columnaActual = 1 Then     'Tipo Estructura 1
              tenro1 = objRs!confval
           ElseIf columnaActual = 2 Then 'Tipo Estructura 2
              tenro2 = objRs!confval
           ElseIf columnaActual = 3 Then 'Tipo Estructura 3
              tenro3 = objRs!confval
           End If
           
           objRs.MoveNext
           
        Loop
        
        objRs.Close
        
        Nro_col = 0
        'cantColumnas = 0
        cantColumna4 = 0
        cantColumna5 = 0
        
        StrSql = " SELECT * FROM confrep WHERE confrep.repnro= 152 AND confnrocol > 3 "
        StrSql = StrSql & "ORDER BY confnrocol "
        
        OpenRecordset StrSql, objRs
        
        Do Until objRs.EOF
           
           Nro_col = Nro_col + 1
           
           columnaActual = CLng(objRs!confnrocol)
                  
           If columnaActual = 4 And (objRs!conftipo = "CO" Or objRs!conftipo = "AC") Then 'Conceptos o Acumuladores de la columna 4
              If objRs!conftipo = "CO" Then
                 StrSql = " SELECT concnro FROM concepto WHERE concepto.conccod= " & objRs!confval2
                 
                 OpenRecordset StrSql, objRs2
                 
                 If Not objRs2.EOF Then
                    Valor = objRs2!concnro
                 End If
              Else
                 Valor = objRs!confval
              End If
                 
              cantColumna4 = cantColumna4 + 1
              TipoCols4(cantColumna4) = objRs!conftipo
              CodCols4(cantColumna4) = Valor
              
           ElseIf columnaActual = 5 And (objRs!conftipo = "CO" Or objRs!conftipo = "AC") Then 'Conceptos o Acumuladores de la columna 5
                  If objRs!conftipo = "CO" Then
                     StrSql = " SELECT concnro FROM concepto WHERE concepto.conccod= " & objRs!confval2
                 
                     OpenRecordset StrSql, objRs2
                 
                     If Not objRs2.EOF Then
                        Valor = objRs2!concnro
                     End If
                  Else
                     Valor = objRs!confval
                  End If
                  
                  cantColumna5 = cantColumna5 + 1
                  TipoCols5(cantColumna5) = objRs!conftipo
                  CodCols5(cantColumna5) = Valor
           End If
           
           objRs.MoveNext
        Loop
        
        objRs.Close
    End If
End Sub

Function cambiarMesALetras(mes)
Dim MesLetra As String

Select Case mes
   Case 1:  MesLetra = "Ene"
   Case 2:  MesLetra = "Feb"
   Case 3:  MesLetra = "Mar"
   Case 4:  MesLetra = "Abr"
   Case 5:  MesLetra = "May"
   Case 6:  MesLetra = "Jun"
   Case 7:  MesLetra = "Jul"
   Case 8:  MesLetra = "Ago"
   Case 9:  MesLetra = "Sep"
   Case 10: MesLetra = "Oct"
   Case 11: MesLetra = "Nov"
   Case 12: MesLetra = "Dic"
End Select

cambiarMesALetras = MesLetra

End Function
Function numberForSQL(Str)
     
  If Not IsNull(Str) Then
     If Len(Str) = 0 Then
        numberForSQL = 0
     Else
        numberForSQL = Replace(Str, ",", ".")
     End If
  End If

End Function
Sub GenerarEncabezado(pliqNro)

Dim rsConsult As New ADODB.Recordset

On Error GoTo MError

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
StrSql = " SELECT pliqmes, pliqanio "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro= " & pliqNro
    
Flog.writeline "Buscando datos del periodo"
           
OpenRecordset StrSql, rsConsult
    
If Not rsConsult.EOF Then
   mesPeriodo = rsConsult!pliqmes
   anioPeriodo = rsConsult!pliqanio
       
   If mesPeriodo = 1 Then
      mesAnterior1 = 12
      mesAnterior2 = 11
      anioAnterior1 = anioPeriodo - 1
      anioAnterior2 = anioPeriodo - 1
   ElseIf mesPeriodo = 2 Then
          mesAnterior1 = 1
          mesAnterior2 = 12
          anioAnterior1 = anioPeriodo
          anioAnterior2 = anioPeriodo - 1
   Else
          mesAnterior1 = mesPeriodo - 1
          mesAnterior2 = mesPeriodo - 2
          anioAnterior1 = anioPeriodo
          anioAnterior2 = anioPeriodo
    End If
End If
    
rsConsult.Close

If tenro1 <> 0 Then
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura  "
    StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro1
        
    OpenRecordset StrSql, rsConsult
        
    If Not rsConsult.EOF Then
       testrnomb1 = rsConsult!tedabr
    Else
       testrnomb1 = ""
    End If
    rsConsult.Close
End If

If tenro2 <> 0 Then
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura  "
    StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro2
        
    OpenRecordset StrSql, rsConsult
        
    If Not rsConsult.EOF Then
       testrnomb2 = rsConsult!tedabr
    Else
       testrnomb2 = ""
    End If
    rsConsult.Close
End If

If tenro3 <> 0 Then
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura  "
    StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro3
        
    OpenRecordset StrSql, rsConsult
        
    If Not rsConsult.EOF Then
       testrnomb3 = rsConsult!tedabr
    Else
       testrnomb3 = ""
    End If
    rsConsult.Close
End If


mes1 = cambiarMesALetras(mesAnterior2) & "-" & Right(CStr(anioAnterior2), 2)
mesPorc1 = "%" & " " & cambiarMesALetras(mesAnterior2) & "/" & cambiarMesALetras(mesAnterior1)
mes2 = cambiarMesALetras(mesAnterior1) & "-" & Right(CStr(anioAnterior1), 2)
mesPorc2 = "%" & " " & cambiarMesALetras(mesAnterior1) & "/" & cambiarMesALetras(mesPeriodo)
mes3 = cambiarMesALetras(mesPeriodo) & "-" & Right(CStr(anioPeriodo), 2)
mes4 = cambiarMesALetras(mesAnterior2) & "-" & Right(CStr(anioAnterior2), 2)
mesPorc3 = "%" & " dif mes"
mes5 = cambiarMesALetras(mesAnterior1) & "-" & Right(CStr(anioAnterior1), 2)
mesPorc4 = "%" & " dif mes"
mes6 = cambiarMesALetras(mesPeriodo) & "-" & Right(CStr(anioPeriodo), 2)
          
StrSql = " INSERT INTO rep_direction ( "
StrSql = StrSql & " bpronro , tedabr1 , tedabr2, tedabr3, "
StrSql = StrSql & " mes_mon_1 , mes_porc_1, mes_mon_2, mes_porc_2, mes_mon_3, "
StrSql = StrSql & " mes_mon_4 , mes_porc_3, mes_mon_5, mes_porc_4, mes_mon_6 "
StrSql = StrSql & " ) VALUES ( "
StrSql = StrSql & NroProceso & ","
StrSql = StrSql & "'" & testrnomb1 & "',"
StrSql = StrSql & "'" & testrnomb2 & "',"
StrSql = StrSql & "'" & testrnomb3 & "',"
StrSql = StrSql & "'" & mes1 & "',"
StrSql = StrSql & "'" & mesPorc1 & "',"
StrSql = StrSql & "'" & mes2 & "',"
StrSql = StrSql & "'" & mesPorc2 & "',"
StrSql = StrSql & "'" & mes3 & "',"
StrSql = StrSql & "'" & mes4 & "',"
StrSql = StrSql & "'" & mesPorc3 & "',"
StrSql = StrSql & "'" & mes5 & "',"
StrSql = StrSql & "'" & mesPorc4 & "',"
StrSql = StrSql & "'" & mes6 & "')"

objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "Error al cargar los datos del reporte. Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub

