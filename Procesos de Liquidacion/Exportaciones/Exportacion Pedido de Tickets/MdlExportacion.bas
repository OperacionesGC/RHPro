Attribute VB_Name = "MdlExportacion"
Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : JMH
' Fecha      : 03/02/2005
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
    
    Nombre_Arch = PathFLog & "Exp_Ped_Tick" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 71 AND bpronro =" & NroProcesoBatch
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
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub
Public Sub Generacion(ByVal PedidoTicket As Long, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Pedido de Ticket
' Autor      : JMH
' Fecha      : 03/02/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim Cabecera As String
Dim Detalle As String
Dim Periodo As String
Dim FechaValidez As String
Dim nombre As String
Dim NroItemPedido As Integer
Dim mes As Integer

Dim rs_CabPedido As New ADODB.Recordset
Dim rs_DetPedido As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Rep_jub_mov As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Const ForReading = 1
Const TristateFalse = 0
Dim fExportCab
Dim fExportDet
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 214"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If




'Activo el manejador de errores
On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\tm_cab_pedidos.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExportCab = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExportCab = fs.CreateTextFile(Archivo, True)
End If

'Archivo para el detalle del Pedido de Pago
Archivo = Directorio & "\tm_det_pedidos.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExportDet = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExportDet = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores

On Error GoTo 0

' Comienzo la transaccion
MyBeginTrans

StrSql = "SELECT emp_ticket.*, tikpedido.*, ticket.tiksigla, ticket.tikdesc, ternro, empleg, terape, terape2, ternom, ternom2, pliqmes, pliqanio "
StrSql = StrSql & " FROM  tikpedido "
StrSql = StrSql & " INNER JOIN emp_ticket ON emp_ticket.tikpednro = tikpedido.tikpednro "
StrSql = StrSql & " INNER JOIN empleado  ON empleado.ternro = emp_ticket.empleado "
StrSql = StrSql & " INNER JOIN ticket ON ticket.tiknro = emp_ticket.tiknro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = tikpedido.pliqnro "
StrSql = StrSql & " WHERE tikpedido.tikpednro = " & PedidoTicket
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_CabPedido

Cantidad = rs_CabPedido.RecordCount
cantidadProcesada = Cantidad

'inicializo
Cabecera = "cus_code;vou_code;dad_code;ord_creation_date;ord_cus_deliv_date;ord_delivery_time;ord_type;ord_period;ord_validity;ord_active;ord_number"
fExportCab.writeline Cabecera

Detalle = "dre_name;per_name;ode_vou_qty;ode_vou_fv;ode_total_amt;ode_booklet;ode_presentation;ped_number;ped_item"
fExportDet.writeline Detalle

Dim Error As Boolean

Do While Not rs_CabPedido.EOF
           
   Error = False
   
   'Valido que la sigla del ticket no sea vacia
   If rs_CabPedido!tiksigla = "" Then
      Flog.writeline "La Sigla del Ticket '" & rs_CabPedido!tikdesc & "' es nula. Por lo tanto no será creada este registro en los archivos de la cabecera y detalle del pedido. "
      Error = True
   End If
   
   If Error = False Then
   
        Periodo = Format_StrNro(rs_CabPedido!pliqmes, 2, True, 0) & Mid(rs_CabPedido!pliqanio, 3, 2)
        
        If Month(rs_CabPedido!tikfechaent) = 12 Then
           mes = 0
           Else: mes = Month(rs_CabPedido!tikfechaent)
        End If
        
        FechaValidez = Format(Day(rs_CabPedido!tikfechaent) & "/" & (mes + 1) & "/" & Year(rs_CabPedido!tikfechaent), "dd/mm/yyyy")
        
        Cabecera = Format_StrNro(rs_CabPedido!empleg, 5, False, "") & ";" & Format_Str(rs_CabPedido!tiksigla, 2, False, "") & ";1;"
        Cabecera = Cabecera & Format(rs_CabPedido!tikpedfecha, "dd/mm/yyyy") & ";" & Format(rs_CabPedido!tikfechaent, "dd/mm/yyyy") & ";"
        Cabecera = Cabecera & "T;N;" & Periodo & ";" & FechaValidez & ";0;1"
        
        fExportCab.writeline Cabecera
        
        ' Buscar el detalle para esa distribucion de ticket
        StrSql = " SELECT * "
        StrSql = StrSql & " FROM emp_tikdist "
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = " & rs_CabPedido!ternro & " AND "
        StrSql = StrSql & " his_estructura.tenro = 2 AND his_estructura.htethasta is null"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " WHERE emp_tikdist.etiknro= " & rs_CabPedido!etiknro
           
        OpenRecordset StrSql, rs_DetPedido
        
        Do While Not rs_DetPedido.EOF
                 
           Error = False
           NroItemPedido = 1
            
           nombre = rs_CabPedido!terape & " " & rs_CabPedido!terape2 & "," & rs_CabPedido!ternom & " " & rs_CabPedido!ternom2
           'Valido que el sector delcliente no sea vacio
           If rs_DetPedido!estrdabr = "" Then
              Flog.writeline "El Sector del Cliente '" & nombre & "' es nulo. Por lo tanto no será creada este registro en los archivos de la cabecera y detalle del pedido. "
              Error = True
           End If
           
           If Error = False Then
                                   
                Detalle = Format_Str(rs_DetPedido!estrdabr, 20, False, "") & ";" & Format_Str(nombre, 30, False, "") & ";" & Format_StrNro(rs_DetPedido!etikdcant, 6, False, "") & ";"
                Detalle = Detalle & rs_DetPedido!etikdmontouni & ";"
                Detalle = Detalle & rs_CabPedido!etikmonto & ";T;S;1;" & NroItemPedido
                fExportDet.writeline Detalle
                
                NroItemPedido = NroItemPedido + 1
           End If
           
           rs_DetPedido.MoveNext
        Loop
             
        rs_DetPedido.Close
        
    End If
    TiempoAcumulado = GetTickCount
          
    cantidadProcesada = cantidadProcesada - 1
          
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((Cantidad - cantidadProcesada) * 100) / Cantidad) & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
   rs_CabPedido.MoveNext
Loop

rs_CabPedido.Close
fExportCab.Close
fExportDet.Close

MyCommitTrans

Set rs_CabPedido = Nothing
Set rs_DetPedido = Nothing

End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : JMH
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim ArrParametros
Dim PedidoTicket As Long

'Orden de los parametros
'Pedido de Ticket

ArrParametros = Split(parametros, "@")
' Levanto cada parametro por separado
PedidoTicket = ArrParametros(0)

Call Generacion(PedidoTicket, bpronro)

End Sub


