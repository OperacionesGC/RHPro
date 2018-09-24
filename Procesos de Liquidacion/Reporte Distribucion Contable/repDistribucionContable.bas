Attribute VB_Name = "repDistribucionContable"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "15/04/2013"
'Global Const UltimaModificacion = " "   'FGZ - Inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "10/05/2013"
'17686 - Deluchi Ezequiel
'Global Const UltimaModificacion = " Se corrigio el objeto de conexion en los borrado de las tablas y se corregio el insert cuando la hora es adicional."
    
'Global Const Version = "1.02"
'Global Const FechaModificacion = "03/06/2013"
'17686 - Deluchi Ezequiel
'Global Const UltimaModificacion = " Se corrigio el codigo de rubro para los gasto."

'Global Const Version = "1.03"
'Global Const FechaModificacion = "07/06/2013"
'17686 - Deluchi Ezequiel
'Global Const UltimaModificacion = " Correcion en el tipo de rubro solo se muestran CON y ACU."

'Global Const Version = "1.04"
'Global Const FechaModificacion = "20/08/2014"
''CAS-25028 - Punto Farma - Modificación de Reporte de Distribución contable CDA - LED
'Global Const UltimaModificacion = "Se agrego distribucion default por horas que no estan cargadas, solo para punto farma"

Global Const Version = "1.05"
Global Const FechaModificacion = "26/11/2014"
Global Const UltimaModificacion = "LM - CAS-21757 - CDA - DIFERENCIA EN CANTIDAD DE REPORTES - se aplico la conversio de hs segun cda para su modelso -> DistribuirHoras "

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Dim fs, f

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
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
Global CodCols(200)
Global CodNovCols(200)
Global CodNov
Global CodDirCols(200)
Global CodDir
Global TitCols(200)
Global TipoCols(200)
Global EsMonto(200)
Global TipoColumna(200)
Global NroCols(200)
Global ValCols(200)
Global CharCols(200)
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
Global prog As Long



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
    
    Nombre_Arch = PathFLog & "Reporte_Dist_Contable" & "-" & NroProceso & ".log"
    
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
    StrSql = StrSql & " WHERE btprcnro = 387 AND bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 387 AND bpronro = " & NroProceso
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


Public Sub GenerarReporte(ByVal Bpronro As Long, ByVal Parametros As String, ByVal Iduser As String, ByVal Fecha As String, ByVal hora As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Pliqnro
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
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
Dim Valor2
Dim cliente As Long

Dim rs As New ADODB.Recordset
Dim rsEmpl As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim rs_confrep As New ADODB.Recordset

    On Error GoTo MError
    
    Flog.writeline "Lista de Parametros = " & Parametros
    ArrParametros = Split(Parametros, "@")
           
    'Obtengo la lista de procesos
    Flog.writeline "Obtengo la Lista de Procesos"
    listapronro = ArrParametros(0)
    Flog.writeline "Lista de Procesos = " & listapronro
    
    
    Flog.writeline "Obtengo el Formato a Usar"
    Formato = ArrParametros(1)
    Flog.writeline "Modelo = " & Formato
    
    'Obtengo el periodo
    Flog.writeline "Obtengo el Período"
    Pliqnro = CLng(ArrParametros(2))
    Flog.writeline "Período = " & Pliqnro
    
    'Obtengo el tipo de Proceso (modelo)
    Flog.writeline "Obtengo el Modelo"
    Modelo = CLng(ArrParametros(3))
    Flog.writeline "Modelo = " & Modelo
        
    StrSql = "SELECT tprocdesc FROM tipoproc WHERE tprocnro = " & Modelo
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        ModeloDesc = rs!tprocdesc
    Else
        ModeloDesc = "Desconocido"
    End If
    
    'Obtengo los cortes de estructura
    Flog.writeline "Obtengo los cortes de estructuras"
    
    Flog.writeline "Obtengo estructura 1"
    tenro1 = CInt(ArrParametros(4))
    estrnro1 = CInt(ArrParametros(5))
    Flog.writeline "Corte 1 = " & tenro1 & " - " & estrnro1
    
    Flog.writeline "Obtengo estructura 2"
    tenro2 = CInt(ArrParametros(6))
    estrnro2 = CInt(ArrParametros(7))
    Flog.writeline "Corte 2 = " & tenro2 & " - " & estrnro2
    
    Flog.writeline "Obtengo estructura 3"
    tenro3 = CInt(ArrParametros(8))
    estrnro3 = CInt(ArrParametros(9))
    Flog.writeline "Corte 3 = " & tenro3 & " - " & estrnro3
    
    
    Flog.writeline "Obtengo las Fechas Desde y Hasta"
    fecEstr = ArrParametros(10)
    fecEstr2 = ArrParametros(11)
    Flog.writeline "Fecha Desde = " & fecEstr
    Flog.writeline "Fecha Hasta = " & fecEstr2
    
    'Tipo de Proyecto
    Flog.writeline "Obtengo Tipo de Proyecto"
    TipoProyecto = CLng(ArrParametros(12))
    Flog.writeline "Tipo de Proyecto = " & TipoProyecto
    
    'chequeo si el proceso corresponse a punto farma o cda, si es 0 es cda
    cliente = 0
    Flog.writeline "Chequeo el cliente"
    If UBound(ArrParametros) = 13 Then
        cliente = CLng(ArrParametros(13))
    End If
    Flog.writeline "cliente = " & cliente
    

    '============================================================================================
    'EMPIEZA EL PROCESO
    
    'Cargo la configuracion del reporte
    Flog.writeline "Cargo la Configuración del Reporte"
    
     StrSql = "SELECT distinct confnrocol,confetiq,conftipo,confval,confval2 FROM confrep "
     StrSql = StrSql & " WHERE repnro = 394"
     StrSql = StrSql & " ORDER BY confnrocol "
     OpenRecordset StrSql, rs_confrep
            
    
    'Obtengo los empleados sobre los que tengo que generar los recibos
    Flog.writeline "Cargo los Empleados "
    Call CargarEmpleados(NroProceso, rsEmpl, Pliqnro, listapronro)
    
    'Borro todos los registros (Para Reprocesamiento)--------------------------
    MyBeginTrans
        'Detalles del detalle
        StrSql = "DELETE rep_dist_cont_det_det WHERE rep_detnro IN ("
        StrSql = StrSql & "SELECT rep_detnro FROM rep_dist_cont_det WHERE bpronro = " & NroProceso & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
              
        'Detalle
        StrSql = "DELETE rep_dist_cont_det WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Cabecera
        StrSql = "DELETE rep_dist_cont WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    'Borro todos los registros -------------------------------------------------
    
    MyBeginTrans
    'Guardo en la BD el encabezado
    Flog.writeline "Genero el encabezado del Reporte"
    Call GenerarEncabezadoReporte(NroProceso, Pliqnro, Iduser, Fecha, hora, tenro1, tenro2, tenro3)
    
    
    Call EstablecerFirmas
    
    

    Progreso = 0
    TotalRubros = 1
    If (rs_confrep.RecordCount <> 0) Then
        TotalRubros = rs_confrep.RecordCount
    End If
    If (rsEmpl.RecordCount <> 0) Then
        'IncPorc = 100 / (rsEmpl.RecordCount * TotalRubros)
        IncPorc = 100 / (rsEmpl.RecordCount)
        TotalEmpleados = rsEmpl.RecordCount
        CantRegistros = rsEmpl.RecordCount
    Else
        TotalEmpleados = 1
        CantRegistros = 1
    End If
     
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                ", bprcempleados ='" & CStr(TotalEmpleados) & "' WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
     
    '-------------------------------------------
    If Not rsEmpl.EOF Then
'        Do While Not rs_Confrep.EOF
'             Columna = rs_Confrep!confnrocol
'             Etiqueta = rs_Confrep!confetiq
'             tipo = rs_Confrep!conftipo
'             Valor = IIf(Not EsNulo(rs_Confrep!confval2), rs_Confrep!confval2, rs_Confrep!confval)
'             Valor2 = rs_Confrep!confval
             
             'rsEmpl.MoveFirst
             Do Until rsEmpl.EOF
         
                 EmpErrores = False
                 Ternro = rsEmpl!Ternro
         
                 'Call GenerarDatosEmpleadoPeriodo(fecEstr, fecEstr2, listapronro, Pliqnro, Ternro, Columna, Etiqueta, tipo, Valor, Valor2)
                 Call GenerarDatosEmpleadoPeriodo(fecEstr, fecEstr2, listapronro, Pliqnro, Ternro, rs_confrep, cliente)
                 rs_confrep.MoveFirst
                
                 'Actualizo el estado del proceso
                 TiempoAcumulado = GetTickCount
             
                 CantRegistros = CantRegistros - 1
             
                Progreso = Progreso + IncPorc
                'bprcprogreso = " & Progreso
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((TotalEmpleados - CantRegistros) * 100) / TotalEmpleados)
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
             
'             rs_Confrep.MoveNext
'        Loop
    End If
    MyCommitTrans
    
Exit Sub

MError:
    Flog.writeline "Error al generando el reporte. Error: " & Err.Description
    Flog.writeline " Ultimo SQL: " & StrSql
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


Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal Pliqnro, ByVal ListaProcesos)
'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Dim StrEmpl As String

    'StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    StrEmpl = " SELECT distinct ternro FROM proceso  " & _
              " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND proceso.pliqnro IN (" & Pliqnro & ")" & _
              " INNER JOIN batch_empleado ON batch_empleado.ternro = cabliq.empleado AND batch_empleado.bpronro = " & NroProc & _
              " WHERE proceso.pronro IN (" & ListaProcesos & ")"

'StrSql = " SELECT proceso.pronro "
'StrSql = StrSql & " FROM proceso "
'StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND proceso.pliqnro IN (" & Pliqnro & ")"
'StrSql = StrSql & " WHERE empleado = " & Ternro
'StrSql = StrSql & "   AND proceso.pliqnro IN (" & Pliqnro & ")"
'StrSql = StrSql & "   AND proceso.pronro IN (" & ListaProcesos & ")"

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




Sub GenerarEncabezadoReporte(ByVal Bpronro As Long, ByVal Pliqnro As Long, ByVal Iduser As String, ByVal Fecha As String, ByVal hora As String, ByVal tenro1 As Long, ByVal tenro2 As Long, ByVal tenro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim teNomb1
Dim teNomb2
Dim teNomb3
Dim I
Dim TituloRep As String

Dim rsConsult As New ADODB.Recordset


teNomb1 = ""
teNomb2 = ""
teNomb3 = ""

If tenro1 <> 0 Then
    StrSql = " SELECT * "
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
    StrSql = " SELECT * "
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
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro3
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb3 = rsConsult!tedabr
    Else
       teNomb3 = ""
    End If
End If


'CREATE TABLE [dbo].[rep_dist_cont](
'    [bpronro] [int] NULL,
'    [formato] [int] NULL,
'    [repdesc] [varchar](100) NULL,
'    [rep_usuario] [varchar](60) NOT NULL,
'    [rep_fecha] [datetime] NOT NULL,
'    [rep_hora] [varchar](8) NULL
') ON [PRIMARY]
'GO


'Descripcion del historico del reporte
    TituloRep = ""
    TituloRep = TituloRep & Bpronro & "-"
    
    StrSql = " SELECT pliqdesc FROM periodo "
    StrSql = StrSql & "  WHERE pliqnro = " & Pliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       TituloRep = TituloRep & rsConsult!PliqDesc
    End If
    TituloRep = TituloRep & " - " & Fecha
    TituloRep = TituloRep & " " & hora


StrSql = " INSERT INTO rep_dist_cont (bpronro ,formato, repdesc, rep_usuario, rep_fecha, rep_hora) VALUES ( "
StrSql = StrSql & NroProceso
StrSql = StrSql & "," & Formato
StrSql = StrSql & ",'" & TituloRep & "'"
StrSql = StrSql & ",'" & Iduser & "'"
StrSql = StrSql & ",'" & Fecha & "'"
StrSql = StrSql & ",'" & Left(hora, 8) & "'"
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords



End Sub



Sub GenerarDatosEmpleadoPeriodo(ByVal Desde As Date, ByVal Hasta As Date, ByVal ListaProcesos, ByVal Pliqnro, ByVal Ternro, ByVal rs_confrep, ByVal cliente As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Legajo As Long
Dim Apellido As String
Dim apellido2 As String
Dim Nombre As String
Dim nombre2 As String
Dim AuxTipo As String
Dim Busca

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3

Dim I
Dim Proceso

Dim PliqDesc
Dim PliqMes
Dim PliqAnio
Dim FPerDesde
Dim FPerHasta
Dim ProDesc
Dim FProDesde
Dim FProHasta
Dim RubroVal As Double
Dim RubroVal2 As Double
Dim Inserta As Boolean
Dim rep_detnro As Long
Dim RubroCod As String
Dim RubroCod2 As String
Dim RubroTipo As String
Dim TotalGasto As Double

Dim Col
Dim Etiqueta
Dim tipo
Dim Valor
Dim Valor2

Dim rs_procesos  As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
   
On Error GoTo MError
Flog.writeline "Entrando a generar dato empleado, ternro: " & Ternro
estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
Proceso = 0
'RubroCod = Valor
'RubroCod2 = Valor2
'RubroTipo = tipo
'------------------------------------------------------------------
'Controlo si el empleado tiene algun proceso en el periodo
'------------------------------------------------------------------
'StrSql = " SELECT proceso.pronro "
'StrSql = StrSql & " FROM proceso "
'StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND proceso.pliqnro IN (" & Pliqnro & ")"
'StrSql = StrSql & " WHERE empleado = " & Ternro
'StrSql = StrSql & "   AND proceso.pliqnro IN (" & Pliqnro & ")"
'StrSql = StrSql & "   AND proceso.pronro IN (" & ListaProcesos & ")"
'OpenRecordset StrSql, rsConsult
'If rsConsult.EOF Then
'   'Si el empleado no tiene procesos en el periodo paso al siguiente
'   rsConsult.Close
'   Flog.writeline "No hay empleados en los procesos seleccionados"
'   Exit Sub
'End If
'rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
'Flog.writeline "Buscando datos del empleado"
OpenRecordset StrSql, rsConsult
nomape = ""
If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom
   nomape = Nombre
   If IsNull(rsConsult!ternom2) Then
      nombre2 = ""
   Else
      nombre2 = rsConsult!ternom2
      nomape = nomape & " " & nombre2
   End If
   Apellido = rsConsult!terape
   nomape = nomape & " " & Apellido
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


'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------
'---LOG---
'Flog.writeline "Buscando datos estructura 1"

If tenro1 <> 0 Then
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro1
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro1 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
    End If
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb1 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

'---LOG---
'Flog.writeline "Buscando datos estructura 2"

If tenro2 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro2
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro2 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
    End If
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb2 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------

'---LOG---
'Flog.writeline "Buscando datos estructura 3"

If tenro3 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro3
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    If estrnro3 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
    End If
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb3 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
StrSql = " SELECT pliqdesc, pliqmes, pliqanio, pliqdesde, pliqhasta, proceso.pronro, tprocdesc, proceso.prodesc, proceso.profecini, proceso.profecfin FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
StrSql = StrSql & " WHERE periodo.pliqnro IN(" & Pliqnro & ")"
StrSql = StrSql & " AND proceso.pronro IN (" & ListaProcesos & ")"
OpenRecordset StrSql, rs_procesos
PliqDesc = ""
Do While Not rs_procesos.EOF
   PliqDesc = rs_procesos!PliqDesc
   PliqMes = rs_procesos!PliqMes
   PliqAnio = rs_procesos!PliqAnio
   FPerDesde = rs_procesos!pliqdesde
   FPerHasta = rs_procesos!pliqhasta
   
   Proceso = rs_procesos!proNro
   ProDesc = rs_procesos!ProDesc
   FProDesde = rs_procesos!profecini
   FProHasta = rs_procesos!profecfin
   ModeloDesc = rs_procesos!tprocdesc
   
   
rs_confrep.MoveFirst
Do While Not rs_confrep.EOF
    Col = rs_confrep!confnrocol
    Etiqueta = rs_confrep!confetiq
    tipo = rs_confrep!conftipo
    

    Valor = IIf(Not EsNulo(rs_confrep!confval2), rs_confrep!confval2, rs_confrep!confval)
    Valor2 = rs_confrep!confval
   
    RubroCod = Valor
    RubroCod2 = Valor2
    RubroTipo = tipo
   
        RubroVal = 0
        TotalGasto = 0
        Inserta = True
        'Busco rubro
        AuxTipo = tipo
        Select Case tipo
         Case "CAM": 'Monto de acumulador de Mensual
             'Flog.writeline "Tipo " & "Monto de acumulador de Mensual"
             RubroTipo = "ACU"
             Busca = "Cantidad"
             RubroVal = bus_ACM(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca, PliqMes, PliqAnio)
         Case "MAM": 'Monto de acumulador Mensual
             'Flog.writeline "Tipo " & "Monto de acumulador Mensual"
             RubroTipo = "ACU"
             Busca = "Monto"
             RubroVal = bus_ACM(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca, PliqMes, PliqAnio)
         Case "CAL": 'Cantidad de acumulador de liquidacion
             'Flog.writeline "Tipo " & "Cantidad de acumulador de liquidacion"
             RubroTipo = "ACU"
             Busca = "Cantidad"
             RubroVal = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "MAL": 'Monto de acumulador de liquidacion
             'Flog.writeline "Tipo " & "Monto de acumulador de liquidacion"
             RubroTipo = "ACU"
             Busca = "Monto"
             RubroVal = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "CCO": 'Cantidad de un concepto
             'Flog.writeline "Tipo " & "Cantidad de un concepto"
             RubroTipo = "CON"
             Busca = "Cantidad"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "MCO": 'Monto de Concepto
             'Flog.writeline "Tipo " & "Monto de Concepto"
             RubroTipo = "CON"
             Busca = "Monto"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "GAS": 'Gasto
             'Flog.writeline "Tipo " & "Gasto"
             RubroTipo = "GAS"
             Busca = ""
             RubroVal = bus_Gastos(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             RubroCod = tipo
         Case "CGA": 'Cantidad de Concepto de Gasto
             'Flog.writeline "Tipo " & "Cantidad de Concepto de Gasto"
             RubroTipo = "CON"
             Busca = "Monto"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             Busca = ""
             TotalGasto = bus_Gastos(Desde, Hasta, Pliqnro, Proceso, Ternro, RubroCod2, Busca)
             'RubroCod = tipo
         Case "MGA": 'Monto de Concepto de Gasto
             'Flog.writeline "Tipo " & "Monto de Concepto de Gasto"
             RubroTipo = "CON"
             Busca = "Monto"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             'RubroCod = tipo
         Case "AGA": 'Acumulador de Gasto
             'Flog.writeline "Tipo " & "Acumulador de Gasto"
             RubroTipo = "AGA"
             Busca = "Monto"
             RubroVal = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             Busca = ""
             TotalGasto = bus_Gastos(Desde, Hasta, Pliqnro, Proceso, Ternro, RubroCod2, Busca)
             'RubroCod = tipo
         Case "CSD": 'Monto de Concepto sin distribucion
             'Flog.writeline "Tipo " & "Monto de Concepto sin distribucion"
             RubroTipo = "CON"
             Busca = "Monto"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             
             Busca = "Cantidad"
             RubroVal2 = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "CDC": 'Monto de Concepto con distribucion sobre cantidad del concepto por % de hs cargadas
             'Flog.writeline "Tipo " & "Monto de Concepto con distribucion sobre cantidad del concepto por % de hs cargadas"
             RubroTipo = "CON"
             Busca = "Monto"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             
             Busca = "Cantidad"
             RubroVal2 = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "CDP": 'Monto de Concepto con distribucion a proyecto especifico
             'Flog.writeline "Tipo " & "Monto de Concepto con distribucion a proyecto especifico"
             RubroTipo = "CON"
             Busca = "Monto"
             RubroVal = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             
             Busca = "Cantidad"
             RubroVal2 = bus_CON(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "ADP": 'Monto de Acumulador de liquidacion con distribucion a proyecto especifico
             'Flog.writeline "Tipo " & "Monto de Acumulador de liquidacion con distribucion a proyecto especifico"
             RubroTipo = "ACU"
             Busca = "Monto"
             RubroVal = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             
             Busca = "Cantidad"
             RubroVal2 = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "ASD": 'Monto de Acumulador de liquidacion sin distribucion de hs
             'Flog.writeline "Tipo " & "Monto de Acumulador de liquidacion sin distribucion de hs"
             RubroTipo = "ACU"
             Busca = "Monto"
             RubroVal = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             
             Busca = "Cantidad"
             RubroVal2 = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
         Case "ADC": 'Monto de Acumulador de liquidacion con distribucion sobre cantidad del acumulador por % de hs cargadas
             'Flog.writeline "Tipo " & "Monto de Acumulador de liquidacion con distribucion sobre cantidad del acumulador por % de hs cargadas"
             RubroTipo = "ACU"
             Busca = "Monto"
             RubroVal = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
             
             Busca = "Cantidad"
             RubroVal2 = bus_ACL(Desde, Hasta, Pliqnro, Proceso, Ternro, Valor, Busca)
        Case "HSC"
            Flog.writeline "Concepto configurado para el 100% de horas: " & rs_confrep!confval2
        Case "HSM"
            Flog.writeline "Concepto configurado para el 100% de horas: " & rs_confrep!confval2
        Case Else
             Flog.writeline "Tipo incorrecto configurado (" & tipo & "). Revisar columna " & Col
             Inserta = False
        End Select
         If RubroVal = 0 Then
             Inserta = False
         Else
             Inserta = True
         End If
        
        If Inserta Then
             'Inserto cabecera
             
             'Flog.writeline "Inserto los datos en la BD - Legajo = " & Legajo & " TERNRO " & Ternro & " proceso " & Proceso & " periodo " & Pliqnro
             
             StrSql = " INSERT INTO rep_dist_cont_det (bpronro,ternro,empleg,nombre,pliqnro,pliqmes,pliqanio,periodo,modelo,proceso,rubro_tipo,rubro_cod,rubro_desc,rubro_val"
             StrSql = StrSql & ",tedabr1,tedabr2,tedabr3) VALUES ("
             StrSql = StrSql & NroProceso
             StrSql = StrSql & "," & Ternro
             StrSql = StrSql & "," & Legajo
             StrSql = StrSql & ",'" & Left(nomape, 200) & "'"
             StrSql = StrSql & "," & Pliqnro
             StrSql = StrSql & "," & PliqMes
             StrSql = StrSql & "," & PliqAnio
             StrSql = StrSql & ",'" & PliqDesc & "'"
             StrSql = StrSql & ",'" & Left(ModeloDesc, 60) & "'"
             StrSql = StrSql & ",'" & Left(ProDesc, 60) & "'"
             StrSql = StrSql & ",'" & Left(RubroTipo, 3) & "'"
             StrSql = StrSql & ",'" & Left(RubroCod, 5) & "'"
             StrSql = StrSql & ",'" & Left(Etiqueta, 200) & "'"
             StrSql = StrSql & "," & RubroVal
             StrSql = StrSql & ",'" & estrnomb1 & "'"
             StrSql = StrSql & ",'" & estrnomb2 & "'"
             StrSql = StrSql & ",'" & estrnomb3 & "'"
             StrSql = StrSql & ")"
             objConn.Execute StrSql, , adExecuteNoRecords
             
             rep_detnro = getLastIdentity(objConn, "rep_dist_cont_det")
             
             'Busco detalle
             Select Case AuxTipo
             Case "CDC", "ADC":
                Select Case cliente
                    Case 0: 'CDA
                        Call DistribuirHorasPorCantidad(Desde, Hasta, Ternro, RubroVal, rep_detnro, RubroVal2)
                    Case 294: 'Punto Farma
                        Call DistribuirRestoHorasPorCantidad(Desde, Hasta, Ternro, RubroVal, rep_detnro, RubroVal2)
                    End Select
             Case "CSD", "ASD":
                 Call NoDistribuirHoras(Desde, Hasta, Ternro, RubroVal, rep_detnro)
             Case "CDP", "ADP":
                 'Call DistribuirHorasAProyec(Desde, Hasta, Ternro, RubroVal, rep_detnro, RubroCod2)
                 Call DistribuirHorasAProyec(Desde, Hasta, Ternro, RubroVal, rep_detnro, RubroCod2, RubroVal2)
             Case "AGA", "CGA":
                 Call DistribuirGastos(Desde, Hasta, Ternro, RubroVal, rep_detnro, RubroCod2, TotalGasto)
             Case "MGA":
                 Call DistribuirGastos2(Desde, Hasta, Ternro, RubroVal, rep_detnro, RubroCod2)
             Case Else
                Select Case cliente
                    Case 0:  'CDA
                        Call DistribuirHoras(Desde, Hasta, Ternro, RubroVal, rep_detnro)
                    Case 294: 'Punto Farma
                        Call DistribuirRestoHoras(Desde, Hasta, Ternro, RubroVal, rep_detnro)
                    End Select
             End Select
        End If
    rs_confrep.MoveNext
   Loop
   rs_procesos.MoveNext
   
Loop


Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "SQL: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

Public Function bus_ACL(ByVal Desde, ByVal Hasta, ByVal Pliqnro, ByVal Proceso, ByVal Ternro, ByVal Rubro, ByVal tipo) As Double
' --------------------------------------------------------------------------------------------
' Descripcion: Funcion que devuelve el monto o cantidad de un acumulador de liquidacion en un proceso
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim I
Dim rsConsult As New ADODB.Recordset
Dim Val As Double

Val = 0
'------------------------------------------------------------------
'Busco los valores de los  acumuladores
'------------------------------------------------------------------
StrSql = " SELECT acu_liq.acunro rubro, alcant cantidad, acu_liq.almonto monto "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro =" & Pliqnro & " AND cabliq.pronro = " & Proceso
StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = " & Proceso
'StrSql = StrSql & " INNER JOIN acumulador on acumulador.acunro= acu_liq.acunro "
StrSql = StrSql & " WHERE acu_liq.acunro = " & Rubro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If UCase(tipo) = "MONTO" Then
        Val = rsConsult!Monto
    Else
        Val = rsConsult!Cantidad
    End If
End If
bus_ACL = Val

End Function


Public Function bus_ACM(ByVal Desde, ByVal Hasta, ByVal Pliqnro, ByVal Proceso, ByVal Ternro, ByVal Rubro, ByVal tipo, ByVal Mes, ByVal Anio) As Double
' --------------------------------------------------------------------------------------------
' Descripcion: Funcion que devuelve el monto o cantidad de un acumulador mensual en un periodo
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim I
Dim rsConsult As New ADODB.Recordset
Dim Val As Double

Val = 0
'------------------------------------------------------------------
'Busco los valores de los  acumuladores
'------------------------------------------------------------------
StrSql = " SELECT acunro rubro, amcant cantidad, ammonto monto "
StrSql = StrSql & " FROM acu_mes "
StrSql = StrSql & " WHERE ternro = " & Ternro
StrSql = StrSql & " AND acunro = " & Rubro
StrSql = StrSql & " AND ammes = " & Mes
StrSql = StrSql & " AND amanio = " & Anio
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If UCase(tipo) = "MONTO" Then
        Val = rsConsult!Monto
    Else
        Val = rsConsult!Cantidad
    End If
End If
bus_ACM = Val

End Function


Public Function bus_CON(ByVal Desde, ByVal Hasta, ByVal Pliqnro, ByVal Proceso, ByVal Ternro, ByVal Rubro, ByVal tipo) As Double
' --------------------------------------------------------------------------------------------
' Descripcion: Funcion que devuelve el monto o cantidad de un detalle de liquidacion en un proceso
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim I
Dim rsConsult As New ADODB.Recordset
Dim Val As Double

Val = 0
'------------------------------------------------------------------
'Busco los valores de los  detliq
'------------------------------------------------------------------
StrSql = " SELECT detliq.concnro rubro, dlicant cantidad, dlimonto monto "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro =" & Pliqnro & " AND cabliq.pronro = " & Proceso
StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = " & Proceso
StrSql = StrSql & " INNER JOIN concepto on concepto.concnro = detliq.concnro "
StrSql = StrSql & " WHERE concepto.conccod = '" & Rubro & "'"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If UCase(tipo) = "MONTO" Then
        Val = rsConsult!Monto
    Else
        Val = rsConsult!Cantidad
    End If
End If
bus_CON = Val

End Function



Public Function bus_Gastos(ByVal Desde, ByVal Hasta, ByVal Pliqnro, ByVal Proceso, ByVal Ternro, ByVal NroProg, ByRef tipo) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna el monto de los Gastos cargados
' Autor      : FGZ
' Fecha      : 11/01/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Parametros
'   Tipo de Gastos      --> 0 = TODOS
'   Paga                --> 0 = AMBOS, 1 =
'   Moneda              -->
'   Estado              --> 0 = TODOS
Dim FechaDeInicio As Date
Dim FechaDeFin As Date
Dim TipoGastos As Long
Dim Paga As Integer
Dim Moneda As Long
Dim Estado As Long
Dim MontoGastos As Double
Dim Firmado As Boolean
Dim Bien As Boolean

Dim Param_cur As New ADODB.Recordset
Dim rs_Gastos As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset

    Bien = False
    MontoGastos = 0
    
    Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    OpenRecordset StrSql, Param_cur
    
    ' Obtener los parametros de la Busqueda
    If Not Param_cur.EOF Then
        If Param_cur!Tprognro = 117 Then
            TipoGastos = IIf(Not EsNulo(Param_cur!Auxint1), Param_cur!Auxint1, 0)
            tipo = TipoGastos
            Paga = IIf(Not EsNulo(Param_cur!Auxint2), Param_cur!Auxint2, 0)
            Moneda = IIf(Not EsNulo(Param_cur!Auxint3), Param_cur!Auxint3, 0)
            Estado = IIf(Not EsNulo(Param_cur!Auxint4), Param_cur!Auxint4, 0)
            Bien = True
        Else
            Flog.writeline Espacios(Tabulador * 4) & "La busqueda configurada NO es de Gastos " & NroProg
        End If
    Else
        Flog.writeline Espacios(Tabulador * 4) & "No se encuentra la Busqueda de Gastos " & NroProg
    End If

    If Bien Then
    
        FechaDeInicio = Desde
        FechaDeFin = Hasta
    
        StrSql = "SELECT gasnro, gasvalor FROM gastos "
        StrSql = StrSql & " WHERE ternro =" & Ternro
        StrSql = StrSql & " AND ( gasfechavuelta >=" & ConvFecha(FechaDeInicio) & " AND gasfechavuelta <= " & ConvFecha(FechaDeFin) & ")"
        If TipoGastos <> 0 Then    ' TODOS
            StrSql = StrSql & " AND tipgasnro = " & TipoGastos
        End If
        If Paga <> -1 Then    ' AMBOS
            StrSql = StrSql & " AND gaspagacliente = " & Paga
        End If
        StrSql = StrSql & " AND monnro = " & Moneda
        If Estado <> 0 Then    ' TODOS
            StrSql = StrSql & " AND gaspagado = " & Estado
        End If
        'StrSql = StrSql & " AND pronro = 0 OR pronro IS NULL"
        OpenRecordset StrSql, rs_Gastos
        Do While Not rs_Gastos.EOF
            Firmado = False
            If Not IsNull(rs_Gastos!gasvalor) Then
                If FirmaActiva165 Then
                    '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                        StrSql = "select cysfirfin from cysfirmas where cysfirfin = -1 " & _
                                 " AND cysfircodext = '" & rs_Gastos!gasnro & "' and cystipnro = 165"
                        OpenRecordset StrSql, rs_firmas
                        If rs_firmas.EOF Then
                            Firmado = False
                        Else
                            Firmado = True
                        End If
                Else
                    Firmado = True
                End If
                
                If Firmado Then
                    MontoGastos = MontoGastos + rs_Gastos!gasvalor
                End If
           End If
           
           rs_Gastos.MoveNext
        Loop
    End If

    bus_Gastos = MontoGastos
    
' Cierro todo y libero
If Param_cur.State = adStateOpen Then Param_cur.Close
Set Param_cur = Nothing

If rs_Gastos.State = adStateOpen Then rs_Gastos.Close
Set rs_Gastos = Nothing

If rs_firmas.State = adStateOpen Then rs_firmas.Close
Set rs_firmas = Nothing
End Function




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



Public Sub DistribuirHoras(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
Dim Distribuye As Boolean

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset


    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = 0
    'StrSql = " SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h "
    'StrSql = StrSql & " WHERE ternro = " & Ternro
    'StrSql = StrSql & " AND fecha <=" & ConvFecha(Hasta) & " AND fecha >=" & ConvFecha(Desde)
    
    
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    If Not rs_Horas.EOF Then
        tothoras = IIf(Not EsNulo(rs_Horas!horas), rs_Horas!horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If tothoras = 0 Then
            Distribuye = False
        Else
            Distribuye = True
        End If
    End If
    
    
    
If Not Distribuye Then
    AuxHoras = 0
    AuxMin = 0
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Proyecto = "** SIN C. COSTO **"
    End If
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    
    Fact = "NO"

    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Else
    StrSql = "SELECT h.ternro, es.estrdabr, z.przonacodext, h.horasfact, sum(h.CantHoras) horas, sum(h.cantmin) min FROM horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    'StrSql = StrSql & " INNER JOIN empleado emp ON emp.ternro = h.ternro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    StrSql = StrSql & " GROUP BY h.ternro, es.estrdabr, z.przonacodext, h.horasfact"
    OpenRecordset StrSql, rs_Horas
    Do While Not rs_Horas.EOF
     
     'debo calcular
     AuxMin = rs_Horas!Min Mod 60
     AuxHoras = rs_Horas!horas + (rs_Horas!Min / 60)
     
     '-----------
     Dim h As Double
     Dim decimales As Double
     Dim horas As Double
     Dim minutos As Double
     horas = CDbl(rs_Horas!horas)
     minutos = CDbl(rs_Horas!Min)
     h = 0
     h = Fix(minutos / 60) 'Paso todas las minutos a horas
     horas = CDbl(horas) + CDbl(h)
     minutos = minutos - (Fix(minutos / 60) * 60)
     'If Minutos = 0 Then Minutos = "00"
     'HorasTotales = Horas & ":" & Minutos
     HS_MIN = Format(horas, "#####00") & ":" & Format(minutos, "00")
     '-----------
        
        
        'HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        AuxHoraDec = CLng(rs_Horas!horas) + (CLng(rs_Horas!Min) / 60)
        
        Porcentaje = AuxHoraDec * 100 / tothoras
        Nominal = RubroVal * Porcentaje / 100
        
        'Debo insertar un detalle
        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & rep_detnro
        StrSql = StrSql & ",'" & Left(rs_Horas!estrdabr, 60) & "'"
        StrSql = StrSql & ",'" & Left(rs_Horas!przonacodext, 20) & "'"
        StrSql = StrSql & ",'" & IIf(rs_Horas!horasfact = 1, "SI", "NO") & "'"
        StrSql = StrSql & ",'" & HS_MIN & "'"
        StrSql = StrSql & "," & Porcentaje
        StrSql = StrSql & "," & Nominal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        rs_Horas.MoveNext
    Loop
End If

End Sub

Public Sub DistribuirHorasPorCantidad(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal RubroVal2)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
Dim Distribuye As Boolean

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset


    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = 0
    'StrSql = " SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h "
    'StrSql = StrSql & " WHERE ternro = " & Ternro
    'StrSql = StrSql & " AND fecha <=" & ConvFecha(Hasta) & " AND fecha >=" & ConvFecha(Desde)
    
    
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    If Not rs_Horas.EOF Then
        tothoras = IIf(Not EsNulo(rs_Horas!horas), rs_Horas!horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If tothoras = 0 Then
            Distribuye = False
        Else
            Distribuye = True
        End If
    End If
    
    
    
If Not Distribuye Then
    AuxHoras = CLng(RubroVal2)
    AuxMin = 0
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Proyecto = "** SIN C. COSTO **"
    End If
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    
    Fact = "NO"

    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Else
    StrSql = "SELECT h.ternro, es.estrdabr, z.przonacodext, h.horasfact, sum(h.CantHoras) horas, sum(h.cantmin) min FROM horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    'StrSql = StrSql & " INNER JOIN empleado emp ON emp.ternro = h.ternro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    StrSql = StrSql & " GROUP BY h.ternro, es.estrdabr, z.przonacodext, h.horasfact"
    OpenRecordset StrSql, rs_Horas
    Do While Not rs_Horas.EOF
        'debo calcular
        AuxMin = rs_Horas!Min Mod 60
        AuxHoras = rs_Horas!horas + (rs_Horas!Min / 60)
        
        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        AuxHoraDec = CLng(rs_Horas!horas) + (CLng(rs_Horas!Min) / 60)
        
        Porcentaje = AuxHoraDec * 100 / tothoras
        
        
        'redefino la cantidad en base al valor de la cantidad del rubro (concepto o acumulador) y al porcentaje de hs
        AuxHoras = CLng(RubroVal2)
        AuxMin = CLng((RubroVal2 - AuxHoras) * 60 / 100)
        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        AuxHoraDec = RubroVal2
        
        HS_MIN = CHorasSF(RubroVal2, 60)
        
        Nominal = RubroVal * Porcentaje / 100
        
        'Debo insertar un detalle
        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & rep_detnro
        StrSql = StrSql & ",'" & Left(rs_Horas!estrdabr, 60) & "'"
        StrSql = StrSql & ",'" & Left(rs_Horas!przonacodext, 20) & "'"
        StrSql = StrSql & ",'" & IIf(rs_Horas!horasfact = 1, "SI", "NO") & "'"
        StrSql = StrSql & ",'" & HS_MIN & "'"
        StrSql = StrSql & "," & Porcentaje
        StrSql = StrSql & "," & Nominal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        rs_Horas.MoveNext
    Loop
End If

End Sub


Public Sub NoDistribuirHoras(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 14/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
    
    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = 0
    'StrSql = " SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h "
    'StrSql = StrSql & " WHERE ternro = " & Ternro
    'StrSql = StrSql & " AND fecha <=" & ConvFecha(Hasta) & " AND fecha >=" & ConvFecha(Desde)
    
    
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    AuxHoras = 0
    AuxMin = 0
    If Not rs_Horas.EOF Then
        'tothoras = IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If Not EsNulo(rs_Horas!horas) Then
            AuxHoras = rs_Horas!horas
        End If
        If Not EsNulo(rs_Horas!Min) Then
            AuxHoras = AuxHoras + (rs_Horas!Min / 60)
            AuxMin = rs_Horas!Min Mod 60
        End If
    End If
       
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Proyecto = "** SIN C. COSTO **"
    End If
    
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    Fact = "NO"
    
    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
        
'Cierro y libero
    If rs_Horas.State = adStateOpen Then rs_Horas.Close
    If rs_Estr.State = adStateOpen Then rs_Estr.Close
    Set rs_Horas = Nothing
    Set rs_Estr = Nothing
End Sub


Public Sub NoDistribuirHorasConCantidad(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal RubroVal2)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 14/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
    
    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = RubroVal2
    AuxHoras = CLng(RubroVal2)
    AuxMin = 0
       
       
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Proyecto = "** SIN C. COSTO **"
    End If
    
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    Fact = "NO"
    
    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
        
'Cierro y libero
    If rs_Horas.State = adStateOpen Then rs_Horas.Close
    If rs_Estr.State = adStateOpen Then rs_Estr.Close
    Set rs_Horas = Nothing
    Set rs_Estr = Nothing
End Sub


Public Sub DistribuirHorasAProyec(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal AuxProyecto, ByVal RubroVal2)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 14/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
    
    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = 0
    
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    StrSql = StrSql & " AND p.proyecnro = " & AuxProyecto
    OpenRecordset StrSql, rs_Horas
    AuxHoras = 0
    AuxMin = 0
    If Not rs_Horas.EOF Then
        'tothoras = IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If Not EsNulo(rs_Horas!horas) Then
            AuxHoras = rs_Horas!horas
        End If
        If Not EsNulo(rs_Horas!Min) Then
            AuxHoras = AuxHoras + (rs_Horas!Min / 60)
            AuxMin = rs_Horas!Min Mod 60
        End If
    Else
        AuxHoras = CLng(RubroVal2)
    End If
    If AuxHoras = 0 And AuxMin = 0 Then
        AuxHoras = CLng(RubroVal2)
    End If
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM estructura "
    StrSql = StrSql & " INNER JOIN proyecto p ON p.ccosto = estructura.estrnro "
    StrSql = StrSql & " WHERE p.proyecnro = " & AuxProyecto
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El proyecto asignado para distribuir no existe " & AuxProyecto
        Proyecto = "** SIN C. COSTO **"
    End If
    
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    Fact = "NO"
    
    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
        
'Cierro y libero
    If rs_Horas.State = adStateOpen Then rs_Horas.Close
    If rs_Estr.State = adStateOpen Then rs_Estr.Close
    Set rs_Horas = Nothing
    Set rs_Estr = Nothing
End Sub



Public Sub DistribuirGastos_old(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal Busqueda)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
Dim Distribuye As Boolean

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim Param_cur As New ADODB.Recordset
Dim rs_Gastos As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset


Dim TipoGastos As Long
Dim Paga As Integer
Dim Moneda As Long
Dim Estado As Long
Dim MontoGastos As Double
Dim Firmado As Boolean
Dim Bien As Boolean
    
    tothoras = 0
    
    
    
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(Busqueda)
    OpenRecordset StrSql, Param_cur
    If Not Param_cur.EOF Then
        If Param_cur!Tprognro = 117 Then
            TipoGastos = IIf(Not EsNulo(Param_cur!Auxint1), Param_cur!Auxint1, 0)
            'tipo = TipoGastos
            Paga = IIf(Not EsNulo(Param_cur!Auxint2), Param_cur!Auxint2, 0)
            Moneda = IIf(Not EsNulo(Param_cur!Auxint3), Param_cur!Auxint3, 0)
            Estado = IIf(Not EsNulo(Param_cur!Auxint4), Param_cur!Auxint4, 0)
            Bien = True
        Else
            Flog.writeline Espacios(Tabulador * 4) & "La busqueda configurada NO es de Gastos " & Busqueda
        End If
    Else
        Flog.writeline Espacios(Tabulador * 4) & "No se encuentra la Busqueda de Gastos " & Busqueda
    End If

If Bien Then
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    If Not rs_Horas.EOF Then
        tothoras = IIf(Not EsNulo(rs_Horas!horas), rs_Horas!horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If tothoras = 0 Then
            Distribuye = False
        Else
            Distribuye = True
        End If
    End If
   
    
    If Distribuye Then
        'En lugar de buscar hs debo buscar gastos
        StrSql = "SELECT gasnro,proyecnro, gasvalor FROM gastos "
        StrSql = StrSql & " WHERE ternro =" & Ternro
        StrSql = StrSql & " AND ( gasfechavuelta >=" & ConvFecha(Desde) & " AND gasfechavuelta <= " & ConvFecha(Hasta) & ")"
        If TipoGastos <> 0 Then    ' TODOS
            StrSql = StrSql & " AND tipgasnro = " & TipoGastos
        End If
        If Paga <> -1 Then    ' AMBOS
            StrSql = StrSql & " AND gaspagacliente = " & Paga
        End If
        StrSql = StrSql & " AND monnro = " & Moneda
        If Estado <> 0 Then    ' TODOS
            StrSql = StrSql & " AND gaspagado = " & Estado
        End If
        StrSql = StrSql & " AND pronro = 0 OR pronro IS NULL"
        OpenRecordset StrSql, rs_Gastos
        Do While Not rs_Gastos.EOF
            Firmado = False
            If Not IsNull(rs_Gastos!gasvalor) Then
                If FirmaActiva165 Then
                    '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                        StrSql = "select cysfirfin from cysfirmas where cysfirfin = -1 " & _
                                 " AND cysfircodext = '" & rs_Gastos!gasnro & "' and cystipnro = 165"
                        OpenRecordset StrSql, rs_firmas
                        If rs_firmas.EOF Then
                            Firmado = False
                        Else
                            Firmado = True
                        End If
                Else
                    Firmado = True
                End If
                If Firmado Then
                
                    'Busco las hs imputadas al proyecto del gasto
                    'StrSql = "SELECT h.ternro, es.estrdabr, z.przonacodext, h.horasfact, sum(h.CantHoras) horas, sum(h.cantmin) min FROM horas h"
                    StrSql = "SELECT es.estrdabr, z.przonacodext, FROM proyecto p ON p.proyecnro = e.proyecnro"
                    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
                    StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
                    StrSql = StrSql & " WHERE p.proyecnro = " & rs_Gastos!proyecnro
                    Select Case TipoProyecto
                        Case 0:
                            StrSql = StrSql & " AND p.proyeccosteable = -1 "
                        Case 1:
                            StrSql = StrSql & " AND p.proyeccosteable = 0 "
                    Case Else
                    StrSql = StrSql & " AND p.proyecnro = " & rs_Gastos!proyecnro
                    End Select
                    OpenRecordset StrSql, rs_Horas
                    Do While Not rs_Horas.EOF
                        'debo calcular
                        'AuxMin = rs_Horas!Min Mod 60
                        'AuxHoras = rs_Horas!Horas + (rs_Horas!Min / 60)
                        AuxMin = 0
                        AuxHoras = 0
                        
                        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
                        AuxHoraDec = CLng(rs_Horas!horas) + (CLng(rs_Horas!Min) / 60)
                        'Porcentaje = AuxHoraDec * 100 / tothoras
                        Porcentaje = 100
                        
                        'Nominal = RubroVal / tothoras * AuxHoraDec
                        Nominal = RubroVal / tothoras * rs_Gastos!gasvalor
                        
                        'Debo insertar un detalle
                        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
                        StrSql = StrSql & ") VALUES ("
                        StrSql = StrSql & rep_detnro
                        StrSql = StrSql & ",'" & Left(rs_Horas!estrdabr, 60) & "'"
                        StrSql = StrSql & ",'" & Left(rs_Horas!przonacodext, 20) & "'"
                        StrSql = StrSql & ",'SI'"
                        StrSql = StrSql & ",'" & HS_MIN & "'"
                        StrSql = StrSql & "," & Porcentaje
                        StrSql = StrSql & "," & Nominal
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                
                        rs_Horas.MoveNext
                    Loop
                End If
           End If
            
            rs_Gastos.MoveNext
        Loop
    End If
End If


' Cierro todo y libero
If Param_cur.State = adStateOpen Then Param_cur.Close
Set Param_cur = Nothing

If rs_Gastos.State = adStateOpen Then rs_Gastos.Close
Set rs_Gastos = Nothing

If rs_firmas.State = adStateOpen Then rs_firmas.Close
Set rs_firmas = Nothing

If rs_Horas.State = adStateOpen Then rs_Horas.Close
Set rs_Horas = Nothing

If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

End Sub

Public Sub DistribuirGastos(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal Busqueda, ByVal TotalGasto)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
'Dim Distribuye As Boolean

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim Param_cur As New ADODB.Recordset
Dim rs_Gastos As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset


Dim TipoGastos As Long
Dim Paga As Integer
Dim Moneda As Long
Dim Estado As Long
Dim MontoGastos As Double
Dim Firmado As Boolean
Dim Bien As Boolean
    
    tothoras = 0
    
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(Busqueda)
    OpenRecordset StrSql, Param_cur
    If Not Param_cur.EOF Then
        If Param_cur!Tprognro = 117 Then
            TipoGastos = IIf(Not EsNulo(Param_cur!Auxint1), Param_cur!Auxint1, 0)
            'tipo = TipoGastos
            Paga = IIf(Not EsNulo(Param_cur!Auxint2), Param_cur!Auxint2, 0)
            Moneda = IIf(Not EsNulo(Param_cur!Auxint3), Param_cur!Auxint3, 0)
            Estado = IIf(Not EsNulo(Param_cur!Auxint4), Param_cur!Auxint4, 0)
            Bien = True
        Else
            Flog.writeline Espacios(Tabulador * 4) & "La busqueda configurada NO es de Gastos " & Busqueda
        End If
    Else
        Flog.writeline Espacios(Tabulador * 4) & "No se encuentra la Busqueda de Gastos " & Busqueda
    End If

    If Bien Then
    
        'StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
        'StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
        'StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
        'StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
        'StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
        'StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
        'StrSql = StrSql & " WHERE h.ternro = " & Ternro
        'StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
        'Select Case TipoProyecto
        '    Case 0:
        '        StrSql = StrSql & " AND p.proyeccosteable = -1 "
        '    Case 1:
        '        StrSql = StrSql & " AND p.proyeccosteable = 0 "
        'Case Else
        'End Select
        'OpenRecordset StrSql, rs_Horas
        'If Not rs_Horas.EOF Then
        '    tothoras = IIf(Not EsNulo(rs_Horas!Horas), rs_Horas!Horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        '    If tothoras = 0 Then
        '        Distribuye = False
        '    Else
        '        Distribuye = True
        '    End If
        'End If
        
        'Zona Default
        Zona = " "
        StrSql = " SELECT estrdabr, estrcodext "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
        OpenRecordset StrSql, rs_Estr
        If Not rs_Estr.EOF Then
           Zona = rs_Estr!estrcodext
        Else
            Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
        End If

        
        StrSql = "SELECT gasnro, p.proyecnro, gasvalor, es.estrdabr, 'SI' horasfact FROM gastos g "
        StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = g.proyecnro"
        StrSql = StrSql & " LEFT JOIN estructura es ON es.estrnro = p.ccosto"
        StrSql = StrSql & " WHERE g.ternro =" & Ternro
        StrSql = StrSql & " AND ( gasfechavuelta >=" & ConvFecha(Desde) & " AND gasfechavuelta <= " & ConvFecha(Hasta) & ")"
        If TipoGastos <> 0 Then    ' TODOS
            StrSql = StrSql & " AND tipgasnro = " & TipoGastos
        End If
        If Paga <> -1 Then    ' AMBOS
            StrSql = StrSql & " AND gaspagacliente = " & Paga
        End If
        StrSql = StrSql & " AND monnro = " & Moneda
        If Estado <> 0 Then    ' TODOS
            StrSql = StrSql & " AND gaspagado = " & Estado
        End If
        'StrSql = StrSql & " AND pronro = 0 OR pronro IS NULL"
        'Select Case TipoProyecto
        '    Case 0:
        '        StrSql = StrSql & " AND p.proyeccosteable = -1 "
        '    Case 1:
        '        StrSql = StrSql & " AND p.proyeccosteable = 0 "
        'Case Else
        'End Select
        OpenRecordset StrSql, rs_Gastos
        Do While Not rs_Gastos.EOF
            Firmado = False
            If Not IsNull(rs_Gastos!gasvalor) Then
                If FirmaActiva165 Then
                    '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                        StrSql = "select cysfirfin from cysfirmas where cysfirfin = -1 " & _
                                 " AND cysfircodext = '" & rs_Gastos!gasnro & "' and cystipnro = 165"
                        OpenRecordset StrSql, rs_firmas
                        If rs_firmas.EOF Then
                            Firmado = False
                        Else
                            Firmado = True
                        End If
                Else
                    Firmado = True
                End If
                If Firmado Then
                    AuxMin = 0
                    AuxHoras = rs_Gastos!gasvalor
                    
                    
                    '--------------------
                    AuxHoras = Int(rs_Gastos!gasvalor)
                    AuxMin = (rs_Gastos!gasvalor - Int(rs_Gastos!gasvalor)) * 60
                    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
                    AuxHoraDec = 0
                    Porcentaje = 100

                    'HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
                    'AuxHoraDec = 0
                    'Porcentaje = 100
                    '--------------------
                    
                    
                    'Nominal = RubroVal / tothoras * AuxHoraDec
                    'Nominal = rs_Gastos!gasvalor
                    If TotalGasto = 0 Then
                        Flog.writeline Espacios(Tabulador * 4) & "El total de Gastos  de la busqueda " & Busqueda & " es 0. No proporciona"
                        Nominal = RubroVal
                    Else
                        Nominal = RubroVal / TotalGasto * rs_Gastos!gasvalor
                        Porcentaje = 100 * rs_Gastos!gasvalor / TotalGasto
                    End If
                    
                    'Debo insertar un detalle
                    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
                    StrSql = StrSql & ") VALUES ("
                    StrSql = StrSql & rep_detnro
                    StrSql = StrSql & ",'" & Left(rs_Gastos!estrdabr, 60) & "'"
                    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
                    StrSql = StrSql & ",'" & rs_Gastos!horasfact & "'"
                    StrSql = StrSql & ",'" & HS_MIN & "'"
                    StrSql = StrSql & "," & Porcentaje
                    StrSql = StrSql & "," & Nominal
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            
            rs_Gastos.MoveNext
        Loop
    End If
    
' Cierro todo y libero
If Param_cur.State = adStateOpen Then Param_cur.Close
Set Param_cur = Nothing

If rs_Gastos.State = adStateOpen Then rs_Gastos.Close
Set rs_Gastos = Nothing

If rs_firmas.State = adStateOpen Then rs_firmas.Close
Set rs_firmas = Nothing

If rs_Horas.State = adStateOpen Then rs_Horas.Close
Set rs_Horas = Nothing

If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

End Sub



Public Sub DistribuirGastos2(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal Busqueda)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
Dim Distribuye As Boolean

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim Param_cur As New ADODB.Recordset
Dim rs_Gastos As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset


Dim TipoGastos As Long
Dim Paga As Integer
Dim Moneda As Long
Dim Estado As Long
Dim MontoGastos As Double
Dim Firmado As Boolean
Dim Bien As Boolean
    
    tothoras = 0
    
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(Busqueda)
    OpenRecordset StrSql, Param_cur
    If Not Param_cur.EOF Then
        If Param_cur!Tprognro = 117 Then
            TipoGastos = IIf(Not EsNulo(Param_cur!Auxint1), Param_cur!Auxint1, 0)
            'tipo = TipoGastos
            Paga = IIf(Not EsNulo(Param_cur!Auxint2), Param_cur!Auxint2, 0)
            Moneda = IIf(Not EsNulo(Param_cur!Auxint3), Param_cur!Auxint3, 0)
            Estado = IIf(Not EsNulo(Param_cur!Auxint4), Param_cur!Auxint4, 0)
            Bien = True
        Else
            Flog.writeline Espacios(Tabulador * 4) & "La busqueda configurada NO es de Gastos " & Busqueda
        End If
    Else
        Flog.writeline Espacios(Tabulador * 4) & "No se encuentra la Busqueda de Gastos " & Busqueda
    End If

    If Bien Then
        StrSql = "SELECT gasnro, p.proyecnro, gasvalor, es.estrdabr, 'SI' horasfact FROM gastos g "
        StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = g.proyecnro"
        StrSql = StrSql & " LEFT JOIN estructura es ON es.estrnro = p.ccosto"
        StrSql = StrSql & " WHERE g.ternro =" & Ternro
        StrSql = StrSql & " AND ( gasfechavuelta >=" & ConvFecha(Desde) & " AND gasfechavuelta <= " & ConvFecha(Hasta) & ")"
        If TipoGastos <> 0 Then    ' TODOS
            StrSql = StrSql & " AND tipgasnro = " & TipoGastos
        End If
        If Paga <> -1 Then    ' AMBOS
            StrSql = StrSql & " AND gaspagacliente = " & Paga
        End If
        StrSql = StrSql & " AND monnro = " & Moneda
        If Estado <> 0 Then    ' TODOS
            StrSql = StrSql & " AND gaspagado = " & Estado
        End If
        'StrSql = StrSql & " AND pronro = 0 OR pronro IS NULL"
        'Select Case TipoProyecto
        '    Case 0:
        '        StrSql = StrSql & " AND p.proyeccosteable = -1 "
        '    Case 1:
        '        StrSql = StrSql & " AND p.proyeccosteable = 0 "
        'Case Else
        'End Select
        OpenRecordset StrSql, rs_Gastos
        Do While Not rs_Gastos.EOF
            Firmado = False
            If Not IsNull(rs_Gastos!gasvalor) Then
                If FirmaActiva165 Then
                    '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                        StrSql = "select cysfirfin from cysfirmas where cysfirfin = -1 " & _
                                 " AND cysfircodext = '" & rs_Gastos!gasnro & "' and cystipnro = 165"
                        OpenRecordset StrSql, rs_firmas
                        If rs_firmas.EOF Then
                            Firmado = False
                        Else
                            Firmado = True
                        End If
                Else
                    Firmado = True
                End If
                If Firmado Then
                    AuxMin = 0
                    AuxHoras = 0
                    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
                    AuxHoraDec = 0
                    Porcentaje = 100
                                        
                    'Zona Default
                    Zona = " "
                    StrSql = " SELECT estrdabr, estrcodext "
                    StrSql = StrSql & " FROM his_estructura "
                    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
                    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
                    OpenRecordset StrSql, rs_Estr
                    If Not rs_Estr.EOF Then
                       Zona = rs_Estr!estrcodext
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
                    End If
                    
                    
                    'Nominal = RubroVal / tothoras * AuxHoraDec
                    'Nominal = RubroVal / tothoras * rs_Gastos!gasvalor
                    Nominal = rs_Gastos!gasvalor
                    
                    'Debo insertar un detalle
                    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
                    StrSql = StrSql & ") VALUES ("
                    StrSql = StrSql & rep_detnro
                    StrSql = StrSql & ",'" & Left(rs_Gastos!estrdabr, 60) & "'"
                    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
                    StrSql = StrSql & ",'" & rs_Gastos!horasfact & "'"
                    StrSql = StrSql & ",'" & HS_MIN & "'"
                    StrSql = StrSql & "," & Porcentaje
                    StrSql = StrSql & "," & Nominal
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            
            rs_Gastos.MoveNext
        Loop
    End If
    
' Cierro todo y libero
If Param_cur.State = adStateOpen Then Param_cur.Close
Set Param_cur = Nothing

If rs_Gastos.State = adStateOpen Then rs_Gastos.Close
Set rs_Gastos = Nothing

If rs_firmas.State = adStateOpen Then rs_firmas.Close
Set rs_firmas = Nothing

If rs_Horas.State = adStateOpen Then rs_Horas.Close
Set rs_Horas = Nothing

If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

End Sub

Public Sub DistribuirRestoHorasPorCantidad(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro, ByVal RubroVal2)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
Dim Distribuye As Boolean
Dim Valor As Double
Dim totalHora As Double
Dim totalNominal As Double
Dim totalPorcentaje As String
Dim ConcNro As String
Dim tipoConcepto As String

Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim rs_valor As New ADODB.Recordset


    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = 0
    
    'busco el concepto configurado para obtener la cantidad de horas
    StrSql = " SELECT concepto.concnro, conftipo FROM confrep " & _
             " INNER JOIN concepto ON concepto.conccod = confrep.confval2 " & _
             " WHERE (upper(conftipo) = 'HSC' OR upper(conftipo) = 'HSM') AND repnro = 394 "
    OpenRecordset StrSql, rs_valor
    If Not rs_valor.EOF Then
        ConcNro = rs_valor!ConcNro
        tipoConcepto = rs_valor!conftipo
    End If
    
    'busco el valor del concepto configurado para obtener la cantidad de horas
    StrSql = " SELECT detliq.dlicant, detliq.dlimonto FROM periodo " & _
             " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
             " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
             " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
             " WHERE cabliq.empleado = " & Ternro & " AND (pliqdesde <= " & ConvFecha(Hasta) & " AND pliqhasta >= " & ConvFecha(Desde) & ") and concnro = " & ConcNro
    OpenRecordset StrSql, rs_valor
    If Not rs_valor.EOF Then
        If UCase(tipoConcepto) = "HSC" Then
            Valor = rs_valor!dlicant
        ElseIf UCase(tipoConcepto) = "HSM" Then
            Valor = rs_valor!dlimonto
        Else
            Valor = 0
        End If
    End If
    
    
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    'StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    If Not rs_Horas.EOF Then
        tothoras = IIf(Not EsNulo(rs_Horas!horas), rs_Horas!horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If tothoras = 0 Then
            Distribuye = False
        Else
            Distribuye = True
        End If
    End If
    
    
    
If Not Distribuye Then
    AuxHoras = CLng(RubroVal2)
    AuxMin = 0
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Proyecto = "** SIN C. COSTO **"
    End If
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    
    Fact = "NO"

    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Else
    StrSql = "SELECT h.ternro, es.estrdabr, h.horasfact, sum(h.CantHoras) horas, sum(h.cantmin) min FROM horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    'StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    'StrSql = StrSql & " INNER JOIN empleado emp ON emp.ternro = h.ternro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    StrSql = StrSql & " GROUP BY h.ternro, es.estrdabr, h.horasfact"
    OpenRecordset StrSql, rs_Horas
    totalHora = 0
    totalNominal = 0
    totalPorcentaje = 0

    Do While Not rs_Horas.EOF
        'debo calcular
        'AuxMin = rs_Horas!Min Mod 60
        'AuxHoras = rs_Horas!Horas + (rs_Horas!Min / 60)
        
        'HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        'AuxHoraDec = CLng(rs_Horas!Horas) + (CLng(rs_Horas!Min) / 60)
        
                
        
        'Porcentaje = AuxHoraDec * 100 / tothoras
        
        'redefino la cantidad en base al valor de la cantidad del rubro (concepto o acumulador) y al porcentaje de hs
        AuxHoras = CLng(RubroVal2)
        AuxMin = CLng((RubroVal2 - AuxHoras) * 60 / 100)
        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        AuxHoraDec = RubroVal2
        totalHora = totalHora + AuxHoraDec
        
        HS_MIN = CHorasSF(RubroVal2, 60)
        
        'Nominal = RubroVal * Porcentaje / 100
        
        If Valor > 0 Then
            Porcentaje = AuxHoraDec * 100 / Valor
        Else
            Porcentaje = AuxHoraDec * 100 / tothoras
        End If
        If Porcentaje > 100 Then
            Porcentaje = 100
        End If
        
        totalPorcentaje = totalPorcentaje + Porcentaje
        
        If Valor > 0 Then
            Nominal = RubroVal * Porcentaje / 100
        Else
            Nominal = RubroVal * Porcentaje / tothoras
        End If
        totalNominal = totalNominal + Nominal

        'Debo insertar un detalle
        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & rep_detnro
        StrSql = StrSql & ",'" & Left(rs_Horas!estrdabr, 60) & "'"
        StrSql = StrSql & ",''"
        StrSql = StrSql & ",'" & IIf(rs_Horas!horasfact = 1, "SI", "NO") & "'"
        StrSql = StrSql & ",'" & HS_MIN & "'"
        StrSql = StrSql & "," & Porcentaje
        StrSql = StrSql & "," & Nominal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        rs_Horas.MoveNext
    Loop
    
    'inserto las horas que no se cargaron
    If totalHora < Valor Then
        Flog.writeline Espacios(Tabulador * 1) & " Entro a distribuir las horas que no cargadas"
        AuxMin = CStr(Round(((Valor - totalHora) - Int((Valor - totalHora))) / 100 * 60, 2)) & "0"
        AuxHoras = Int(Valor - totalHora)
        Flog.writeline Espacios(Tabulador * 1) & " Horas sin cargar: " & AuxHoras
        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        Porcentaje = 100 - totalPorcentaje
        
        If Valor > 0 Then
            Nominal = (RubroVal * Porcentaje / 100)
        Else
            Nominal = (RubroVal * Porcentaje) / tothoras
        End If
        
        'Busco el centro de costo
        StrSql = " SELECT estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
        OpenRecordset StrSql, rs_Estr
        If Not rs_Estr.EOF Then
           Proyecto = rs_Estr!estrdabr
        Else
            Proyecto = "** SIN C. COSTO **"
        End If
        
        'Zona Default
        Zona = " "
        StrSql = " SELECT estrdabr, estrcodext "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
        OpenRecordset StrSql, rs_Estr
        If Not rs_Estr.EOF Then
           Zona = rs_Estr!estrcodext
        Else
            Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
        End If
        
        Fact = "NO"
    
        'Debo insertar un detalle
        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & rep_detnro
        StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
        StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
        StrSql = StrSql & ",'" & Fact & "'"
        StrSql = StrSql & ",'" & HS_MIN & "'"
        StrSql = StrSql & "," & Porcentaje
        StrSql = StrSql & "," & Nominal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If

End If

End Sub

Public Sub DistribuirRestoHoras(ByVal Desde, ByVal Hasta, ByVal Ternro, ByVal RubroVal, ByVal rep_detnro)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 07/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim AuxHoras As String
Dim AuxMin As String
Dim HS_MIN As String
Dim tothoras As Double
Dim AuxHoraDec As Double
Dim Porcentaje As Double
Dim Nominal As Double
Dim Proyecto As String
Dim Zona As String
Dim Fact As String
Dim Distribuye As Boolean
Dim ConcNro As String
Dim tipoConcepto As String
Dim Valor As Double
Dim totalHora As Double
Dim totalNominal As Double
Dim totalPorcentaje As String


Dim rs_Horas As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim rs_valor As New ADODB.Recordset


    'TipoProyecto
    '0 - Costeable
    '1 - NO Costeable
    '2 - AMBOS
    
    tothoras = 0
    ConcNro = 0
    tipoConcepto = ""
    'busco el concepto configurado para obtener la cantidad de horas
    StrSql = " SELECT concepto.concnro, conftipo FROM confrep " & _
             " INNER JOIN concepto ON concepto.conccod = confrep.confval2 " & _
             " WHERE (upper(conftipo) = 'HSC' OR upper(conftipo) = 'HSM') AND repnro = 394 "
    OpenRecordset StrSql, rs_valor
    If Not rs_valor.EOF Then
        ConcNro = rs_valor!ConcNro
        tipoConcepto = rs_valor!conftipo
    End If
    
    'busco el valor del concepto configurado para obtener la cantidad de horas
    StrSql = " SELECT detliq.dlicant, detliq.dlimonto FROM periodo " & _
             " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
             " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
             " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
             " WHERE cabliq.empleado = " & Ternro & " AND (pliqdesde <= " & ConvFecha(Hasta) & " AND pliqhasta >= " & ConvFecha(Desde) & ") and concnro = " & ConcNro
    OpenRecordset StrSql, rs_valor
    If Not rs_valor.EOF Then
        If UCase(tipoConcepto) = "HSC" Then
            Valor = rs_valor!dlicant
        ElseIf UCase(tipoConcepto) = "HSM" Then
            Valor = rs_valor!dlimonto
        Else
            Valor = 0
        End If
    End If
    
    
    StrSql = "SELECT sum(h.CantHoras) horas, sum(h.cantmin) min FROM Horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    'StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    OpenRecordset StrSql, rs_Horas
    If Not rs_Horas.EOF Then
        tothoras = IIf(Not EsNulo(rs_Horas!horas), rs_Horas!horas, 0) + (IIf(Not EsNulo(rs_Horas!Min), rs_Horas!Min / 60, 0))
        If tothoras = 0 Then
            Distribuye = False
        Else
            Distribuye = True
        End If
    End If
    
If Not Distribuye Then
    AuxHoras = 0
    AuxMin = 0
    HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
    Porcentaje = 100
    Nominal = RubroVal
    
    'Busco el centro de costo
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Proyecto = rs_Estr!estrdabr
    Else
        Proyecto = "** SIN C. COSTO **"
    End If
    
    'Zona Default
    Zona = " "
    StrSql = " SELECT estrdabr, estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
    OpenRecordset StrSql, rs_Estr
    If Not rs_Estr.EOF Then
       Zona = rs_Estr!estrcodext
    Else
        Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
    End If
    
    Fact = "NO"

    'Debo insertar un detalle
    StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & rep_detnro
    StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
    StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
    StrSql = StrSql & ",'" & Fact & "'"
    StrSql = StrSql & ",'" & HS_MIN & "'"
    StrSql = StrSql & "," & Porcentaje
    StrSql = StrSql & "," & Nominal
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Else
    StrSql = "SELECT h.ternro, es.estrdabr, h.horasfact, sum(h.CantHoras) horas, sum(h.cantmin) min FROM horas h"
    StrSql = StrSql & " INNER JOIN tarea t ON t.tareanro = H.tareanro"
    StrSql = StrSql & " INNER JOIN etapas e ON e.etapanro = t.etapanro"
    StrSql = StrSql & " INNER JOIN proyecto p ON p.proyecnro = e.proyecnro"
    StrSql = StrSql & " INNER JOIN estructura es ON es.estrnro = p.ccosto"
    'StrSql = StrSql & " INNER JOIN proyecto_zona z ON z.przonanro = H.przonanro"
    'StrSql = StrSql & " INNER JOIN empleado emp ON emp.ternro = h.ternro"
    StrSql = StrSql & " WHERE h.ternro = " & Ternro
    StrSql = StrSql & " AND h.fecha <=" & ConvFecha(Hasta) & " AND h.fecha >=" & ConvFecha(Desde)
    Select Case TipoProyecto
        Case 0:
            StrSql = StrSql & " AND p.proyeccosteable = -1 "
        Case 1:
            StrSql = StrSql & " AND p.proyeccosteable = 0 "
    Case Else
    End Select
    StrSql = StrSql & " GROUP BY h.ternro, es.estrdabr, h.horasfact"
    OpenRecordset StrSql, rs_Horas
    totalHora = 0
    totalNominal = 0
    totalPorcentaje = 0
    Do While Not rs_Horas.EOF
        'debo calcular
        
        AuxMin = rs_Horas!Min Mod 60
        AuxHoras = rs_Horas!horas + (rs_Horas!Min / 60)
        
        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        AuxHoraDec = CLng(rs_Horas!horas) + (CLng(rs_Horas!Min) / 60)
        totalHora = totalHora + AuxHoraDec
        
        If Valor > 0 Then
            Porcentaje = AuxHoraDec * 100 / Valor
        Else
            Porcentaje = AuxHoraDec * 100 / tothoras
        End If
        
        If Porcentaje > 100 Then
            Porcentaje = 100
        End If

        totalPorcentaje = totalPorcentaje + Porcentaje
        
        If Valor > 0 Then
            Nominal = RubroVal * Porcentaje / 100
        Else
            Nominal = RubroVal * Porcentaje / tothoras
        End If
        totalNominal = totalNominal + Nominal
        
        'Debo insertar un detalle
        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & rep_detnro
        StrSql = StrSql & ",'" & Left(rs_Horas!estrdabr, 60) & "'"
        StrSql = StrSql & ",''"
        StrSql = StrSql & ",'" & IIf(rs_Horas!horasfact = 1, "SI", "NO") & "'"
        StrSql = StrSql & ",'" & HS_MIN & "'"
        StrSql = StrSql & "," & Porcentaje
        StrSql = StrSql & "," & Nominal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        rs_Horas.MoveNext
    Loop
    
    'inserto las horas que no se cargaron
    If totalHora < Valor Then
        Flog.writeline Espacios(Tabulador * 1) & " Entro a distribuir las horas que no cargadas"
        AuxMin = CStr(Round(((Valor - totalHora) - Int((Valor - totalHora))) / 100 * 60, 2)) & "0"
        AuxHoras = Int(Valor - totalHora)
        Flog.writeline Espacios(Tabulador * 1) & " Horas sin cargar: " & AuxHoras
        HS_MIN = Format(AuxHoras, "#####00") & ":" & Format(AuxMin, "00")
        Porcentaje = 100 - totalPorcentaje
        
        If Valor > 0 Then
            Nominal = (RubroVal * Porcentaje / 100)
        Else
            Nominal = (RubroVal * Porcentaje) / tothoras
        End If
        
        'Busco el centro de costo
        StrSql = " SELECT estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 5"
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
        OpenRecordset StrSql, rs_Estr
        If Not rs_Estr.EOF Then
           Proyecto = rs_Estr!estrdabr
        Else
            Proyecto = "** SIN C. COSTO **"
        End If
        
        'Zona Default
        Zona = " "
        StrSql = " SELECT estrdabr, estrcodext "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & "    WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 173"
        StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(Hasta) & " AND (htethasta is null or htethasta>=" & ConvFecha(Hasta) & "))"
        OpenRecordset StrSql, rs_Estr
        If Not rs_Estr.EOF Then
           Zona = rs_Estr!estrcodext
        Else
            Flog.writeline Espacios(Tabulador * 1) & " El empleado no tiene estructura de Zona default asignada en la fecha " & Hasta
        End If
        
        Fact = "NO"
    
        'Debo insertar un detalle
        StrSql = " INSERT INTO rep_dist_cont_det_det (rep_detnro,Proyecto,zona,Adicional,Horas,PorcAsig,ValorNominal"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & rep_detnro
        StrSql = StrSql & ",'" & Left(Proyecto, 60) & "'"
        StrSql = StrSql & ",'" & Left(Zona, 20) & "'"
        StrSql = StrSql & ",'" & Fact & "'"
        StrSql = StrSql & ",'" & HS_MIN & "'"
        StrSql = StrSql & "," & Porcentaje
        StrSql = StrSql & "," & Nominal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
End If

End Sub

Public Function CHorasSF(ByVal Cantidad As Single, ByVal Dur As Single) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna un string con la cantidad de hs y minutos a partir de un valor decimal
' Autor      : FGZ
' Fecha      : 09/11/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim minutos As Single
Dim horas As Single
    If Dur = 0 Then
        Dur = 60
    End If
    
    Cantidad = Cantidad * Dur
    horas = Int(Cantidad / Dur)
    minutos = Cantidad Mod Dur
    CHorasSF = Format(horas, "#####00") & ":" & Format(minutos, "00")
End Function
