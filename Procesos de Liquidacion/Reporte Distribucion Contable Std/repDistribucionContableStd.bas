Attribute VB_Name = "repDistribucionContable"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "01/03/2013"
'Global Const UltimaModificacion = " "   'LED - Inicial - CAS-13764 - H&A - Imputacion Contable

Global Const Version = "1.01"
Global Const FechaModificacion = "15/04/2013"
Global Const UltimaModificacion = " "   'LED - CAS-13764 - Correccion en el caso que las horas cargadas son mayores que las que trabaja mensualmente.

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Dim fs, f

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global tenro1 As Long
Global estrnro1 As Long
Global tenro2 As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global fecEstr As String
Global fecEstr2 As String
Global Formato As Long
Global TipoProyecto As Long
Global Modelo As Long
Global ModeloDesc As String
Global CantColumnas As Long
Global descDesde
Global descHasta
Global FechaHasta
Global FechaDesde
Global Nro_Col
Global listaPer
Global concAnt
Global Desde
Global Hasta
Global nomape
Global prog As Double



Private Sub Main()
Dim NombreArchivo As String
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim objRs As New ADODB.Recordset
Dim Parametros As String

Dim rep_DC_IDuser As String
Dim rep_DC_Fecha As String
Dim rep_DC_Hora As String

Dim strTempo As String
Dim orden
Dim rs_confrep

Dim ArrParametros
Dim PID As String


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


    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo CE
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "Distribucion_Contable_Std" & "-" & NroProceso & ".log"
    
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
    
    Flog.writeline "Inicio Proceso: " & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID
    StrSql = StrSql & " WHERE btprcnro = 388 AND bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 388 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = objRs!bprcparam
       rep_DC_IDuser = objRs!Iduser
       rep_DC_Fecha = objRs!bprcfecha
       rep_DC_Hora = objRs!bprchora
       Call GenerarReporte(NroProceso, Parametros, rep_DC_IDuser, rep_DC_Fecha, rep_DC_Hora)
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
    Flog.writeline
    MyRollbackTransliq
End Sub


Public Sub GenerarReporte(ByVal bpronro As Long, ByVal Parametros As String, ByVal Iduser As String, ByVal Fecha As String, ByVal hora As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : LED
' Fecha      : 19/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pliqnro
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim proNro
Dim Ternro
Dim arrpronro
Dim Periodos
Dim I
Dim TotalEmpleados
Dim TotalRubros
Dim CantRegistros
Dim PID As String
Dim TituloReporte As String
Dim ArrParametros
Dim Columna
Dim Etiqueta
Dim tipo
Dim Valor
Dim repnro
Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim rs As New ADODB.Recordset
Dim rsEmpl As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim rs_confrep As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset

    On Error GoTo MError
    
    Flog.writeline "Lista de Parametros = " & Parametros
    ArrParametros = Split(Parametros, "@")
           
    'Obtengo el periodo
    Flog.writeline "Obtengo el Período"
    pliqnro = CLng(ArrParametros(0))
    Flog.writeline "Período = " & pliqnro
    
    
    'Obtengo los cortes de estructura
    Flog.writeline "Obtengo los cortes de estructuras"
    
    Flog.writeline "Obtengo estructura 1"
    tenro1 = CInt(ArrParametros(1))
    estrnro1 = CInt(ArrParametros(2))
    Flog.writeline "Corte 1 = " & tenro1 & " - " & estrnro1
    
    Flog.writeline "Obtengo estructura 2"
    tenro2 = CInt(ArrParametros(3))
    estrnro2 = CInt(ArrParametros(4))
    Flog.writeline "Corte 2 = " & tenro2 & " - " & estrnro2
    
    Flog.writeline "Obtengo estructura 3"
    tenro3 = CInt(ArrParametros(5))
    estrnro3 = CInt(ArrParametros(6))
    Flog.writeline "Corte 3 = " & tenro3 & " - " & estrnro3
    
    
    Flog.writeline "Obtengo las Fechas Desde y Hasta"
    fecEstr = ArrParametros(7)
    fecEstr2 = ArrParametros(8)
    Flog.writeline "Fecha Desde = " & fecEstr
    Flog.writeline "Fecha Hasta = " & fecEstr2
    
    'Tipo de Proyecto
    Flog.writeline "Obtengo Tipo de Proyecto"
    TipoProyecto = CLng(ArrParametros(9))
    Flog.writeline "Tipo de Proyecto = " & TipoProyecto
    
    
    '============================================================================================
    'EMPIEZA EL PROCESO
    
    'Cargo la configuracion del reporte
    Flog.writeline "Cargo la Configuración del Reporte"
        
    'Obtengo los empleados sobre los que tengo que generar los recibos
    Flog.writeline "Cargo los Empleados "
    Call CargarEmpleados(NroProceso, rsEmpl)
    
    'Borro todos los registros (Para Reprocesamiento)--------------------------
    MyBeginTrans
        'Detalles
        StrSql = "DELETE rep_dist_cont_std_det WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                      
        'Cabecera
        StrSql = "DELETE rep_dist_cont_std WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    'Borro todos los registros -------------------------------------------------
    
    'MyBeginTrans
    'Guardo en la BD el encabezado
    Flog.writeline "Genero el encabezado del Reporte"
    Call GenerarEncabezadoReporte(NroProceso, pliqnro, Iduser, Fecha, hora, tenro1, tenro2, tenro3)
    repnro = getLastIdentity(objConn, "rep_dist_cont_std")
    
    Call EstablecerFirmas
    
    
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                ", bprcempleados ='" & CStr(CantRegistros) & "' WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    prog = 0
    
    If (rsEmpl.RecordCount <> 0) Then
        Progreso = 100 / (rsEmpl.RecordCount)
        TotalEmpleados = rsEmpl.RecordCount
        CantRegistros = rsEmpl.RecordCount
        rsEmpl.MoveFirst
    Else
        TotalEmpleados = 1
        CantRegistros = 1
    End If
         
    
    Do Until rsEmpl.EOF
        prog = prog + Progreso
        EmpErrores = False
        Ternro = rsEmpl!Ternro
        
        
       '------------------------------------------------------------------
       'Busco los datos del empleado
       '------------------------------------------------------------------
       StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
       StrSql = StrSql & " FROM empleado "
       StrSql = StrSql & " WHERE ternro= " & Ternro
       Flog.writeline "Buscando datos del empleado"

       OpenRecordset StrSql, rsConsult
       nomape = ""
       If Not rsConsult.EOF Then
          nombre = rsConsult!ternom
          nomape = nombre
          If IsNull(rsConsult!ternom2) Then
             nombre2 = ""
          Else
             nombre2 = rsConsult!ternom2
             nomape = nomape & " " & nombre2
          End If
          apellido = rsConsult!terape
          nomape = nomape & " " & apellido
          If IsNull(rsConsult!terape2) Then
             apellido2 = ""
          Else
             apellido2 = rsConsult!terape2
             nomape = nomape & " " & apellido2
          End If
          Legajo = rsConsult!empleg
       Else
          Flog.writeline "Error al obtener los datos del empleado"
       '   GoTo MError
       End If
       rsConsult.Close
        
        Call DistribuirHoras(fecEstr, fecEstr2, Ternro, Legajo, NroProceso, nomape, repnro)
    
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
    
        CantRegistros = CantRegistros - 1
    
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & prog
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(CantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

      'Si se generaron todos los datos del empleado correctamente lo borro
      If Not EmpErrores Then
          StrSql = " DELETE FROM batch_empleado "
          StrSql = StrSql & " WHERE bpronro = " & NroProceso
          StrSql = StrSql & " AND ternro = " & Ternro
          'objConn.Execute StrSql, , adExecuteNoRecords
      End If

       
        rsEmpl.MoveNext
    Loop
    
Exit Sub

MError:
    Flog.writeline "Error al generando el reporte. Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    MyRollbackTrans
    Exit Sub
End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function


'--------------------------------------------------------------------
' Se encarga de generar los datos para el empleado por cada proceso
'--------------------------------------------------------------------


Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)
'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc & " ORDER BY ternro "
    OpenRecordset StrEmpl, rsEmpl
End Sub

Function numberForSQL(Str)
     
  If Not IsNull(Str) Then
     If Len(Str) = 0 Then
        numberForSQL = 0
     Else
        numberForSQL = Replace(Str, ",", ".")
     End If
  End If

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function


Function sinDatos(Str)
  If IsNull(Str) Then
     sinDatos = True
  Else
     If Trim(Str) = "" Then
        sinDatos = True
     Else
        sinDatos = False
     End If
  End If

End Function




Sub GenerarEncabezadoReporte(ByVal bpronro As Long, ByVal pliqnro As Long, ByVal Iduser As String, ByVal Fecha As String, ByVal hora As String, ByVal tenro1 As Long, ByVal tenro2 As Long, ByVal tenro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : LED
' Fecha      : 19/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim teNomb1
Dim teNomb2
Dim teNomb3
Dim pliqmes
Dim pliqanio

Dim I
Dim TituloRep As String

Dim rsConsult As New ADODB.Recordset


teNomb1 = ""
teNomb2 = ""
teNomb3 = ""

If tenro1 <> 0 Then
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro1
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb1 = rsConsult!tedabr
    Else
       teNomb1 = ""
    End If
End If

If tenro2 <> 0 Then
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro2
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb2 = rsConsult!tedabr
    Else
       teNomb2 = ""
    End If
End If

If tenro3 <> 0 Then
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro3
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb3 = rsConsult!tedabr
    Else
       teNomb3 = ""
    End If
End If

'Descripcion del historico del reporte
    TituloRep = ""
    TituloRep = TituloRep & bpronro & "-"
    
    StrSql = " SELECT pliqdesc, pliqmes, pliqanio FROM periodo "
    StrSql = StrSql & "  WHERE pliqnro = " & pliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       TituloRep = TituloRep & rsConsult!PliqDesc
       pliqmes = rsConsult!pliqmes
       pliqanio = rsConsult!pliqanio
    End If
    TituloRep = TituloRep & " - " & Fecha
    TituloRep = TituloRep & " " & hora


StrSql = " INSERT INTO rep_dist_cont_std (bpronro,formato,repdesc,rep_user,pliqnro,pliqmes,pliqanio,tedabr1,tedabr2,tedabr3) VALUES ( "
StrSql = StrSql & NroProceso
StrSql = StrSql & ",1"
StrSql = StrSql & ",'" & TituloRep & "'"
StrSql = StrSql & ",'" & Iduser & "'"
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & teNomb1 & "'"
StrSql = StrSql & ",'" & teNomb2 & "'"
StrSql = StrSql & ",'" & teNomb3 & "'"
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords



End Sub

Public Sub EstablecerFirmas()
Dim rs_cystipo As New ADODB.Recordset

    
    FirmaActiva5 = False
    FirmaActiva15 = False
    FirmaActiva19 = False
    FirmaActiva20 = False
    FirmaActiva165 = False
    
    StrSql = "select cystipnro from cystipo where (cystipnro = 5 or cystipnro = 15 OR cystipnro = 19 or cystipnro = 20 or cystipnro = 165) AND cystipact = -1"
    OpenRecordset StrSql, rs_cystipo
    
    Do While Not rs_cystipo.EOF
    Select Case rs_cystipo!cystipnro
    Case 5:
        FirmaActiva5 = True
    Case 15:
        FirmaActiva15 = True
    Case 19:
        FirmaActiva19 = True
    Case 20:
        FirmaActiva20 = True
    Case 165:
        FirmaActiva165 = True
    
    Case Else
    End Select
        
        rs_cystipo.MoveNext
    Loop
    
If rs_cystipo.State = adStateOpen Then rs_cystipo.Close
Set rs_cystipo = Nothing

End Sub



Public Sub DistribuirHoras(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal empleg, ByVal bpronro, ByVal nombre, ByVal repnro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim rs_Horas As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset

Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothorasSinCC As Double
Dim tothorasCC As Double
Dim tothorasLic As Double
Dim tothorasSinCargar As Double
Dim AuxHoraDec As Double
Dim porcentaje As Double
Dim Nominal As Double
Dim cantDiasLaborables As Double
Dim cantHorasLaborables As Double
Dim cantHorasLaborablesCargadas As Double
Dim cantHorasLaborablesxDia As Double
Dim confactivo As Integer
Dim centroCosto As String
Dim estrnroCentroCosto As Long
Dim porcentajeAcumulado As Double
Dim horasAcumulado As Double
    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothorasSinCC = 0
    tothorasCC = 0
    porcentajeAcumulado = 0
    horasAcumulado = 0
    
    'Obtengo los dias laborables en el periodo
    Call bus_DiasHabiles(Desde, Hasta, Ternro, cantDiasLaborables)
     
    'Obtengo la cantidad de horas que trabaja el empleado por dia
    StrSql = " SELECT horasdia From his_estructura " & _
             " INNER JOIN regimenHorario ON regimenHorario.estrnro = his_estructura.estrnro " & _
             " WHERE Ternro = " & Ternro & " AND his_estructura.Tenro = 21 " & _
             " AND ((his_estructura.htetdesde <= " & ConvFecha(Desde) & "  AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(Hasta) & " " & _
             " OR his_estructura.htethasta >= " & ConvFecha(Desde) & ")) OR " & _
             " (his_estructura.htetdesde >= " & ConvFecha(Desde) & " AND (his_estructura.htetdesde <= " & ConvFecha(Hasta) & "))) "
    OpenRecordset StrSql, rs_Horas
    'por defecto asumo 8 horas
    cantHorasLaborablesxDia = 8
    If Not rs_Horas.EOF Then
        cantHorasLaborablesxDia = rs_Horas!horasdia
    End If
    
    'Calculo la cantidad de horas por periodo que debe trabajar
    cantHorasLaborables = cantDiasLaborables * cantHorasLaborablesxDia
    
    
    'Chequeo si tengo que buscar licencias o si no es necesario
    StrSql = " SELECT confactivo FROM confper WHERE confper.confnro = 9 "
    OpenRecordset StrSql, rs_Aux
    
    If Not rs_Aux.EOF Then
        confactivo = rs_Aux!confactivo
    Else
        confactivo = 0
    End If
    
    If Not confactivo Then
        Flog.writeline "Busqueda de horas de licencias para ternro: " & Ternro
        'Si esta activo ya se resolvio, sino esta activo tengo que buscar que centro de costo corresponde
        Call horasLicencia(Desde, Hasta, Ternro, tothorasLic, cantHorasLaborablesxDia, centroCosto, estrnroCentroCosto)
    End If
    
    'calculo el total de horas cargadas (sin cc, con cc y licencias) para ver si supera cantHorasLaborables
    StrSql = " SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h  " & _
             " INNER JOIN tarea t ON t.tareanro = H.tareanro  " & _
             " INNER JOIN etapas e ON e.etapanro = t.etapanro  " & _
             " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro " & _
             " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro " & _
             " INNER JOIN estructura es ON es.estrnro = p.ccosto  " & _
             " WHERE h.ternro = " & Ternro & _
             " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    OpenRecordset StrSql, rs_Horas
    cantHorasLaborablesCargadas = IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
    cantHorasLaborablesCargadas = cantHorasLaborablesCargadas + tothorasLic
    If cantHorasLaborables < cantHorasLaborablesCargadas Then
        cantHorasLaborables = cantHorasLaborablesCargadas
        Flog.writeline "La cantidad de Horas cargadas para el: " & Ternro & " superan la cantidad de horas laborables en el mes."
    End If
    'calculo el porcentaje de licencias una vez que ya se cual es la cantidad de horas que representa el 100%
    If Not confactivo Then
        If tothorasLic <> 0 Then
            Call calcularPorcentaje(cantHorasLaborables, tothorasLic, porcentaje)
            porcentajeAcumulado = porcentajeAcumulado + porcentaje
            horasAcumulado = horasAcumulado + tothorasLic
            StrSql = " INSERT INTO rep_dist_cont_std_det (repnro, bpronro, ternro, empleg, nombre, centro_costo, estrnro_cc, horas, porcasig) " & _
                     " VALUES (" & repnro & "," & bpronro & "," & Ternro & "," & empleg & ",'" & nombre & "','" & centroCosto & "'," & estrnroCentroCosto & ",'" & tothorasLic & ":00'," & porcentaje & ") "
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    'sumo las horas de los proyectos que NO tienen asociado ningun centro de costo
    StrSql = " SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h " & _
             " INNER JOIN tarea t ON t.tareanro = H.tareanro " & _
             " INNER JOIN etapas e ON e.etapanro = t.etapanro " & _
             " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro " & _
             " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro " & _
             " WHERE h.ternro = " & Ternro & _
             " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde) & " AND p.ccosto = 0 "
    Select Case TipoProyecto
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 2:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    
    
    If Not rs_Horas.EOF Then
        Flog.writeline "Busqueda de horas sin centro de costo para ternro: " & Ternro
        tothorasSinCC = IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If tothorasSinCC <> 0 Then
            Call calcularPorcentaje(cantHorasLaborables, tothorasSinCC, porcentaje)
            porcentajeAcumulado = porcentajeAcumulado + porcentaje
            horasAcumulado = horasAcumulado + tothorasSinCC
            StrSql = " INSERT INTO rep_dist_cont_std_det (repnro, bpronro, ternro, empleg, nombre, centro_costo, estrnro_cc, horas, porcasig) " & _
                     " VALUES (" & repnro & "," & bpronro & "," & Ternro & "," & empleg & ",'" & nombre & "','**** Sin Centro Costo ****',0,'" & Format(IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min, 0) / 60), "#####00") & ":" & Format(IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min, 0) Mod 60, "00") & "'," & porcentaje & ") "
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    'sumo las horas de los proyectos que tienen asociado centro de costo
    StrSql = " SELECT sum(h.CantHoras) horas, sum(h.cantmin) min, p.ccosto, es.estrdabr FROM Horas h " & _
             " INNER JOIN tarea t ON t.tareanro = H.tareanro " & _
             " INNER JOIN etapas e ON e.etapanro = t.etapanro " & _
             " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro " & _
             " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro " & _
             " INNER JOIN estructura es ON es.estrnro = p.ccosto " & _
             " WHERE h.ternro = " & Ternro & _
             " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde) & " AND p.ccosto <> 0 "
    Select Case TipoProyecto
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 2:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    StrSql = StrSql & " GROUP BY p.CCosto, es.estrdabr "
    OpenRecordset StrSql, rs_Horas
    
    Do While Not rs_Horas.EOF
        Flog.writeline "Busqueda de centro de costo: " & rs_Horas!estrdabr & " - " & rs_Horas!CCosto & " para ternro: " & Ternro
        tothorasCC = IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        Call calcularPorcentaje(cantHorasLaborables, tothorasCC, porcentaje)
        porcentajeAcumulado = porcentajeAcumulado + porcentaje
        horasAcumulado = horasAcumulado + tothorasCC
        StrSql = " INSERT INTO rep_dist_cont_std_det (repnro, bpronro, ternro, empleg, nombre, centro_costo, estrnro_cc, horas, porcasig) " & _
                 " VALUES (" & repnro & "," & bpronro & "," & Ternro & "," & empleg & ",'" & nombre & "','" & rs_Horas!estrdabr & "'," & rs_Horas!CCosto & ",'" & Format(rs_Horas!Horas + (rs_Horas!Min / 60), "#####00") & ":" & Format(rs_Horas!Min Mod 60, "00") & "'," & porcentaje & ") "
        objConn.Execute StrSql, , adExecuteNoRecords
        rs_Horas.MoveNext
    Loop
    

    
    'horas sin cargar
    Flog.writeline "Busqueda de Horas Sin cargar para ternro: " & Ternro
    tothorasSinCargar = cantHorasLaborables - horasAcumulado
    If tothorasSinCargar > 0 Then
        'Call calcularPorcentaje(cantHorasLaborables, tothorasSinCargar, porcentaje)
        porcentaje = CDbl(FormatNumber(100, 2)) - CDbl(FormatNumber(porcentajeAcumulado, 2))
        Call centroCostoHorasSinCargar(Desde, Hasta, Ternro, centroCosto, estrnroCentroCosto)
        If estrnroCentroCosto <> 0 Then
        StrSql = " INSERT INTO rep_dist_cont_std_det (repnro, bpronro, ternro, empleg, nombre, centro_costo, estrnro_cc, horas, porcasig) " & _
                 " VALUES (" & repnro & "," & bpronro & "," & Ternro & "," & empleg & ",'" & nombre & "','" & centroCosto & "'," & estrnroCentroCosto & ",'" & Format(Fix(tothorasSinCargar), "#####00") & ":" & Format((tothorasSinCargar - Fix(tothorasSinCargar)) * 60, "00") & "'," & porcentaje & ") "
        objConn.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline "Ocurrio un error en la busqueda del centro de costo asociado a la falta de carga de horas para el ternro: " & Ternro
        End If
    Else
        Flog.writeline "No hay horas sin cargar para el ternro: " & Ternro
    End If
'Cierro y libero
If rs_Horas.State = adStateOpen Then rs_Horas.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close
    
End Sub


Public Sub bus_DiasHabiles(ByVal Desde As Date, ByVal Hasta As Date, ByVal Ternro As Long, ByRef Valor As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias Habiles entre dos fechas
' Autor      : FGZ
' Fecha      : 05/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim DiasHabiles As Double
Dim Dia As Date
Dim EsFeriado As Boolean
Dim objFeriado As New Feriado

' inicializacion de variables
Set objFeriado.Conexion = objConn

    
Dia = Desde
Do While Dia <= Hasta
    
    EsFeriado = objFeriado.Feriado(Dia, Ternro, False)
    
    If Not EsFeriado And Not Weekday(Dia) = 1 And Not Weekday(Dia) = 7 Then
        ' No es feriado no Domingo, no Sabado
        DiasHabiles = DiasHabiles + 1
    End If
    Dia = Dia + 1
Loop
Valor = DiasHabiles
End Sub

Public Sub calcularPorcentaje(ByVal total As Double, ByVal parcial As Double, ByRef porcentaje As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula el porcentaje segun cantidad total y cantidad parcial
' Autor      : LED
' Fecha      : 21/02/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If total <> 0 Then
        porcentaje = FormatNumber(((parcial * 100) / total), 2)
    Else
        porcentaje = 100
    End If
End Sub

Public Sub horasLicencia(ByVal Desde As Date, ByVal Hasta As Date, ByVal Ternro As Long, ByRef tothorasLic As Double, ByVal horasXDia As Double, ByRef centroCosto As String, ByRef estrnroCentroCosto As Long)
Dim rs_estr As New ADODB.Recordset
Dim rs_confrep As New ADODB.Recordset
Dim rs_proy As New ADODB.Recordset

    'busco estructura gerencia (6), que posee el empleado
    StrSql = " SELECT his_estructura.estrnro, estrdabr " & _
             " From estructura " & _
             " INNER JOIN his_estructura ON his_estructura.estrnro = estructura.estrnro AND his_estructura.tenro = 6 " & _
             " WHERE ternro = " & Ternro & " AND ((his_estructura.htetdesde <= " & ConvFecha(Desde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(Hasta) & _
             " or his_estructura.htethasta >= " & ConvFecha(Desde) & ")) OR " & _
             " (his_estructura.htetdesde >= " & ConvFecha(Desde) & " AND (his_estructura.htetdesde <= " & ConvFecha(Hasta) & "))) "
    OpenRecordset StrSql, rs_estr
    
    If Not rs_estr.EOF Then
        'busco el confrep el proyecto asociado a la estructura
        StrSql = " SELECT confval2 FROM confrep WHERE confval = " & rs_estr!Estrnro & " AND conftipo = 'HLC' AND repnro = 295 "
        OpenRecordset StrSql, rs_confrep
        
        If Not rs_confrep.EOF Then
            StrSql = " SELECT CCosto, estrdabr FROM proyecto " & _
                     " INNER JOIN estructura ON proyecto.CCosto = estructura.estrnro " & _
                     " WHERE proyecnro = " & rs_confrep!confval2
            OpenRecordset StrSql, rs_proy
            If Not rs_proy.EOF Then
                centroCosto = rs_proy!estrdabr
                estrnroCentroCosto = rs_proy!CCosto
                'una vez resuelto el centro de costo, se debe calcular la horas
                'Busco las licencias segun sean parcial o completas
                StrSql = " SELECT elfechadesde, elfechahasta,elhoradesde,elhorahasta,elcanthrs, eltipo " & _
                         " FROM emp_lic WHERE licestnro = 2 AND empleado = " & Ternro & _
                         " AND ((elfechadesde <= " & ConvFecha(Desde) & " AND (elfechahasta is null or elfechahasta >= " & ConvFecha(Hasta) & _
                         " OR elfechahasta >= " & ConvFecha(Desde) & ")) OR (elfechadesde >= " & ConvFecha(Desde) & " AND (elfechadesde <= " & ConvFecha(Hasta) & ")) ) "
                OpenRecordset StrSql, rs_proy
                tothorasLic = 0
                Do While Not rs_proy.EOF
                    Select Case rs_proy!eltipo
                        Case 1:
                            Call bus_DiasHabiles(rs_proy!elfechadesde, rs_proy!elfechahasta, Ternro, tothorasLic)
                            tothorasLic = tothorasLic * horasXDia
                        Case 2:
                            tothorasLic = tothorasLic + (rs_proy!elhorahasta - rs_proy!elhoradesde)
                        Case 3:
                            tothorasLic = tothorasLic + rs_proy!elcanthrs
                    End Select
                    rs_proy.MoveNext
                Loop
            Else
                Flog.writeline "No existe el proyecto configurado en reporte."
            End If
        Else
            Flog.writeline "No existe la estructura de tipo Gerencia del tercero: " & Ternro & " configurado en reporte."
        End If
    Else
        Flog.writeline "El tercero: " & Ternro & " no tiene estructura de tipo Gerencia cargada."
    End If

'libero
If rs_proy.State = adStateOpen Then rs_proy.Close
If rs_confrep.State = adStateOpen Then rs_confrep.Close
If rs_estr.State = adStateOpen Then rs_estr.Close
    
End Sub

Public Sub centroCostoHorasSinCargar(ByVal Desde As Date, ByVal Hasta As Date, ByVal Ternro As Long, ByRef centroCosto As String, ByRef estrnroCentroCosto As Long)
Dim rs_estr As New ADODB.Recordset
Dim rs_confrep As New ADODB.Recordset
Dim rs_proy As New ADODB.Recordset

    
    estrnroCentroCosto = 0
    'busco estructura gerencia (6), que posee el empleado
    StrSql = " SELECT his_estructura.estrnro, estrdabr " & _
             " From estructura " & _
             " INNER JOIN his_estructura ON his_estructura.estrnro = estructura.estrnro AND his_estructura.tenro = 6 " & _
             " WHERE ternro = " & Ternro & " AND ((his_estructura.htetdesde <= " & ConvFecha(Desde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(Hasta) & _
             " or his_estructura.htethasta >= " & ConvFecha(Desde) & ")) OR " & _
             " (his_estructura.htetdesde >= " & ConvFecha(Desde) & " AND (his_estructura.htetdesde <= " & ConvFecha(Hasta) & "))) "
    OpenRecordset StrSql, rs_estr
    If Not rs_estr.EOF Then
        'busco el confrep el proyecto asociado a la estructura
        StrSql = " SELECT confval2 FROM confrep WHERE confval = " & rs_estr!Estrnro & " AND conftipo = 'HNC' AND repnro = 295 "
        OpenRecordset StrSql, rs_confrep
        
        If Not rs_confrep.EOF Then
            StrSql = " SELECT CCosto, estrdabr FROM proyecto " & _
                     " INNER JOIN estructura ON proyecto.CCosto = estructura.estrnro " & _
                     " WHERE proyecnro = " & rs_confrep!confval2
            OpenRecordset StrSql, rs_proy
            If Not rs_proy.EOF Then
                centroCosto = rs_proy!estrdabr
                estrnroCentroCosto = rs_proy!CCosto
            Else
                Flog.writeline "No existe el proyecto configurado en reporte."
            End If
        Else
            Flog.writeline "No existe la estructura de tipo Gerencia del tercero: " & Ternro & " configurado en reporte."
        End If
    Else
        Flog.writeline "El tercero: " & Ternro & " no tiene estructura de tipo Gerencia cargada."
    End If

'libero
If rs_proy.State = adStateOpen Then rs_proy.Close
If rs_confrep.State = adStateOpen Then rs_confrep.Close
If rs_estr.State = adStateOpen Then rs_estr.Close

End Sub
