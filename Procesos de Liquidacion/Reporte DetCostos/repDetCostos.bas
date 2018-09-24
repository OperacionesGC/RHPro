Attribute VB_Name = "repDetCostos"
Option Explicit

Global Const Version = "1.00"
Global Const FechaModificacion = "16/10/2009"
Global Const UltimaModificacion = "Encriptacion de string de conexion"
Global Const UltimaModificacion1 = "Manuel Lopez"

'--------------------------------------------------------------------------

Dim fs, f
' Global Flog

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

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String
Global Formato As Integer
Global Modelo As Long
Global CantColumnas As Integer
Global CodCols(200)
Global TitCols(200)
Global TipoCols(200)
Global NroCols(200)
Global ValCols(200)
Global TituloRep As String
Global descDesde
Global descHasta
Global fechaHasta
Global fechaDesde
Global Nro_Col

Const CasoSinNiveles = 1
Const CasoCon1Nivel = 2
Const CasoCon2Nivel = 3
Const CasoCon3Nivel = 4


Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim pliqdesde
Dim pliqhasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim proNro
Dim Ternro
Dim arrpronro
Dim Periodos
Dim rsEmpl As New ADODB.Recordset
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden

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
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteDetCostos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "_________________________________________________________________"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "               " & UltimaModificacion1
    Flog.writeline "_________________________________________________________________"
    Flog.writeline
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    Flog.writeline "Inicio Proceso de Conceptos y Acumuladores : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)
       
       'Obtengo el modelo a usar
       Modelo = CLng(ArrParametros(1))
       
       'Obtengo el periodo desde
       pliqdesde = CLng(ArrParametros(2))
       
       'Obtengo el periodo hasta
       pliqhasta = CLng(ArrParametros(3))
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(4))
       estrnro1 = CInt(ArrParametros(5))
       tenro2 = CInt(ArrParametros(6))
       estrnro2 = CInt(ArrParametros(7))
       tenro3 = CInt(ArrParametros(8))
       estrnro3 = CInt(ArrParametros(9))
       fecEstr = ArrParametros(10)
       
       'EMPIEZA EL PROCESO
       'Busco el periodo desde
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqdesde
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          fechaDesde = objRs!pliqdesde
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
          fechaHasta = objRs!pliqhasta
          descHasta = objRs!pliqDesc
       Else
          Flog.writeline "No se encontro el periodo hasta."
          Exit Sub
       End If
        
       objRs.Close
       
       'Cargo la configuracion del reporte
       Call CargarConfiguracionReporte(Modelo)
      
       'Obtengo los empleados sobre los que tengo que generar los datos
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       'Guardo en la BD el encabezado
       Call GenerarEncabezadoReporte
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
          arrpronro = Split(listapronro, ",")
          EmpErrores = False
          Ternro = rsEmpl!Ternro
          orden = rsEmpl!estado
          
          'Genero una entrada para el empleado por cada proceso
          For I = 0 To UBound(arrpronro)
             proNro = arrpronro(I)
             Flog.writeline "Generando datos empleado " & Ternro & " para el proceso " & proNro
             
             Call GenerarDatosEmpleado(proNro, Ternro, orden)
             
          Next
          
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
              StrSql = StrSql & " AND ternro = " & Ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
          End If
          
          rsEmpl.MoveNext
       Loop
        
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

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
Sub GenerarDatosEmpleado(proNro, Ternro, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim CasoEstructura As Integer

Dim EstrNomb1
Dim EstrNomb2
Dim EstrNomb3
Dim proDesc
Dim pliqDesc
Dim pliqNro
Dim pliqFecha
Dim I
Dim cliqnro As Long

Dim estr1Ant
Dim estr2Ant
Dim estr3Ant
Dim te1Ant
Dim te2Ant
Dim te3Ant

Dim guardar As Boolean

On Error GoTo MError

EstrNomb1 = ""
EstrNomb2 = ""
EstrNomb3 = ""

'------------------------------------------------------------------
'Controlo si el empleado esta en el proceso
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM cabliq"
StrSql = StrSql & " WHERE empleado = " & Ternro
StrSql = StrSql & "   AND pronro   = " & proNro

OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
   'Si el empleado no esta en el proceso entonces paso al proximo
   rsConsult.Close
   
   Exit Sub
Else
   cliqnro = rsConsult!cliqnro
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro

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
'Busco los datos del proceso
'------------------------------------------------------------------
StrSql = " SELECT * FROM proceso WHERE pronro = " & proNro

OpenRecordset StrSql, rsConsult

proDesc = ""
If Not rsConsult.EOF Then
   proDesc = rsConsult!proDesc
   pliqNro = rsConsult!pliqNro
End If

rsConsult.Close

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
'Busco en que caso de corte de estructuras estoy
'------------------------------------------------------------------

If tenro1 = 0 Then
   CasoEstructura = CasoSinNiveles
Else
   If tenro2 = 0 Then
      CasoEstructura = CasoCon1Nivel
   Else
      If tenro3 = 0 Then
         CasoEstructura = CasoCon2Nivel
      Else
         CasoEstructura = CasoCon3Nivel
      End If
   End If
End If

'------------------------------------------------------------------
'Busco los datos en detcostos
'------------------------------------------------------------------

Select Case CasoEstructura
    Case CasoSinNiveles
        StrSql = "SELECT * FROM detcostos "
        StrSql = StrSql & " WHERE cliqnro= " & cliqnro
        StrSql = StrSql & "   AND pronro = " & proNro
        StrSql = StrSql & "   AND tenro1 IS NULL "
        StrSql = StrSql & "   AND tenro2 IS NULL "
        StrSql = StrSql & "   AND tenro3 IS NULL "
    
    
    Case CasoCon1Nivel
        StrSql = "SELECT * FROM detcostos "
        StrSql = StrSql & " WHERE cliqnro= " & cliqnro
        StrSql = StrSql & "   AND pronro = " & proNro
        StrSql = StrSql & "   AND "
        
        StrSql = StrSql & " ( (tenro1 = " & tenro1
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro1
        End If
        StrSql = StrSql & " ) "
   
        StrSql = StrSql & " OR (tenro2 = " & tenro1
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro1
        End If
        StrSql = StrSql & " ) "
   
        StrSql = StrSql & " OR (tenro3 = " & tenro1
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro1
        End If
        StrSql = StrSql & " ) ) "
   
    
    Case CasoCon2Nivel

         StrSql = "SELECT * FROM detcostos "
         StrSql = StrSql & " WHERE cliqnro= " & cliqnro
         StrSql = StrSql & "   AND pronro = " & proNro
         
         'Armo una exp. booleana con todas las combinaciones de dos tipos de estructuras
         StrSql = StrSql & "   AND ( "
         
         '-- 1 --
         StrSql = StrSql & " ( "
         StrSql = StrSql & " tenro1 = " & tenro1
         StrSql = StrSql & " AND tenro2 = " & tenro2
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro2
         End If
         StrSql = StrSql & " ) "
    
         '-- 2 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro1 = " & tenro1
         StrSql = StrSql & " AND tenro3 = " & tenro2
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro2
         End If
         StrSql = StrSql & " ) "
         
         '-- 3 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro2 = " & tenro1
         StrSql = StrSql & " AND tenro3 = " & tenro2
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro2
         End If
         StrSql = StrSql & " ) "
         
         '-- 4 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro2 = " & tenro1
         StrSql = StrSql & " AND tenro1 = " & tenro2
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro2
         End If
         StrSql = StrSql & " ) "
    
         '-- 5 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro3 = " & tenro1
         StrSql = StrSql & " AND tenro1 = " & tenro2
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro2
         End If
         StrSql = StrSql & " ) "
         
         '-- 6 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro3 = " & tenro1
         StrSql = StrSql & " AND tenro2 = " & tenro2
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro2
         End If
         StrSql = StrSql & " ) "
         
         'Encierra a todos los OR
         StrSql = StrSql & " ) "
        
    
    Case CasoCon3Nivel
        
         StrSql = "SELECT * FROM detcostos "
         StrSql = StrSql & " WHERE cliqnro= " & cliqnro
         StrSql = StrSql & "   AND pronro = " & proNro
         
         'Armo una exp. booleana con todas las combinaciones de dos tipos de estructuras
         StrSql = StrSql & "   AND ( "
         
         '-- 1 --
         StrSql = StrSql & " ( "
         StrSql = StrSql & " tenro1 = " & tenro1
         StrSql = StrSql & " AND tenro2 = " & tenro2
         StrSql = StrSql & " AND tenro3 = " & tenro3
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro2
         End If
         If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro3
         End If
         
         StrSql = StrSql & " ) "
    
         '-- 2 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro1 = " & tenro1
         StrSql = StrSql & " AND tenro3 = " & tenro2
         StrSql = StrSql & " AND tenro2 = " & tenro3
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro2
         End If
         If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro3
         End If
         
         StrSql = StrSql & " ) "
         
         '-- 3 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro2 = " & tenro1
         StrSql = StrSql & " AND tenro1 = " & tenro2
         StrSql = StrSql & " AND tenro3 = " & tenro3
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro2
         End If
         If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro3
         End If
         
         StrSql = StrSql & " ) "
         
         '-- 4 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro2 = " & tenro1
         StrSql = StrSql & " AND tenro3 = " & tenro2
         StrSql = StrSql & " AND tenro1 = " & tenro3
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro2
         End If
         If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro3
         End If
         
         StrSql = StrSql & " ) "
    
         '-- 5 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro3 = " & tenro1
         StrSql = StrSql & " AND tenro1 = " & tenro2
         StrSql = StrSql & " AND tenro2 = " & tenro3
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro2
         End If
         If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro3
         End If
         
         StrSql = StrSql & " ) "
         
         '-- 6 --
         StrSql = StrSql & " OR ( "
         StrSql = StrSql & " tenro3 = " & tenro1
         StrSql = StrSql & " AND tenro2 = " & tenro2
         StrSql = StrSql & " AND tenro1 = " & tenro3
         If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estrnro3 = " & estrnro1
         End If
         If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estrnro2 = " & estrnro2
         End If
         If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estrnro1 = " & estrnro3
         End If
         
         StrSql = StrSql & " ) "
         
         'Encierra a todos los OR
         StrSql = StrSql & " ) "
        
End Select

StrSql = StrSql & " ORDER BY tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3  "


'Obtengo los datos de los conceptos y acumuladores
Call borrarValores

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   estr1Ant = rsConsult!estrnro1
   estr2Ant = rsConsult!estrnro2
   estr3Ant = rsConsult!estrnro3
   te1Ant = rsConsult!tenro1
   te2Ant = rsConsult!tenro2
   te3Ant = rsConsult!tenro3
End If

guardar = False

Do Until rsConsult.EOF
   If igualEstr(estr1Ant, rsConsult!estrnro1) And igualEstr(estr2Ant, rsConsult!estrnro2) And igualEstr(estr3Ant, rsConsult!estrnro3) Then
        If CInt(rsConsult!tipoorigen) = 1 Then
           Call agregarValor("CO", rsConsult!Origen, rsConsult!Monto)
        Else
           Call agregarValor("AC", rsConsult!Origen, rsConsult!Monto)
        End If
        
        rsConsult.MoveNext
   Else
        guardar = True
   End If
   
   If rsConsult.EOF Then
      guardar = True
   End If
   
   If guardar Then
        guardar = False
        
        If hayDatos() Then
        
            If Not IsNull(estr1Ant) Then
                StrSql = " SELECT estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " WHERE estructura.estrnro = " & estr1Ant
                       
                OpenRecordset StrSql, rsConsult2
                
                If Not rsConsult2.EOF Then
                   EstrNomb1 = rsConsult2!estrdabr
                Else
                   EstrNomb1 = ""
                End If
                
                rsConsult2.Close
            End If
    
            If Not IsNull(estr2Ant) Then
                StrSql = " SELECT estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " WHERE estructura.estrnro = " & estr2Ant
                       
                OpenRecordset StrSql, rsConsult2
                
                If Not rsConsult2.EOF Then
                   EstrNomb2 = rsConsult2!estrdabr
                Else
                   EstrNomb2 = ""
                End If
                
                rsConsult2.Close
            End If
            
            If Not IsNull(estr3Ant) Then
                StrSql = " SELECT estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " WHERE estructura.estrnro = " & estr3Ant
                       
                OpenRecordset StrSql, rsConsult2
                
                If Not rsConsult2.EOF Then
                   EstrNomb3 = rsConsult2!estrdabr
                Else
                   EstrNomb3 = ""
                End If
                
                rsConsult2.Close
            End If
            
            Select Case CasoEstructura
                Case CasoSinNiveles
                    EstrNomb1 = ""
                    EstrNomb2 = ""
                    EstrNomb3 = ""
                
                Case CasoCon1Nivel
                    EstrNomb1 = buscarTE(tenro1, te1Ant, te2Ant, te3Ant, EstrNomb1, EstrNomb2, EstrNomb3)
                    EstrNomb2 = ""
                    EstrNomb3 = ""
                
                Case CasoCon2Nivel
                    EstrNomb1 = buscarTE(tenro1, te1Ant, te2Ant, te3Ant, EstrNomb1, EstrNomb2, EstrNomb3)
                    EstrNomb2 = buscarTE(tenro2, te1Ant, te2Ant, te3Ant, EstrNomb1, EstrNomb2, EstrNomb3)
                    EstrNomb3 = ""
                
                
                Case CasoCon3Nivel
                    EstrNomb1 = buscarTE(tenro1, te1Ant, te2Ant, te3Ant, EstrNomb1, EstrNomb2, EstrNomb3)
                    EstrNomb2 = buscarTE(tenro2, te1Ant, te2Ant, te3Ant, EstrNomb1, EstrNomb2, EstrNomb3)
                    EstrNomb3 = buscarTE(tenro3, te1Ant, te2Ant, te3Ant, EstrNomb1, EstrNomb2, EstrNomb3)
                
                
            End Select
            
            
            '-------------------------------------------------------------------------------
            'Inserto los datos en la BD
            '-------------------------------------------------------------------------------
            
            StrSql = " INSERT INTO rep_detcostos_det ( " & _
                     " bpronro , legajo, ternro, apellido, apellido2, " & _
                     " nombre , nombre2, pronro, pliqnro,pliqfecha, " & _
                     " prodesc , pliqdesc, estrdabr1, estrdabr2, estrdabr3, " & _
                     " colval1 , colval2, colval3, colval4, colval5, " & _
                     " colval6 , colval7, colval8, colval9, colval10," & _
                     " colval11 , colval12, colval13, colval14, colval15, " & _
                     " colval16 , colval17, colval18, colval19, colval20, " & _
                     " colval21 , colval22, colval23, colval24, colval25, " & _
                     " colval26 , colval27, colval28, colval29, colval30, " & _
                     " colval31 , colval32, colval33, colval34, colval35, " & _
                     " colval36 , colval37, colval38, colval39, colval40, " & _
                     " colval41 , colval42, colval43, colval44, colval45, " & _
                     " colval46 , colval47, colval48, colval49, colval50, " & _
                     " orden ) VALUES ( "
            
            StrSql = StrSql & NroProceso & ","
            StrSql = StrSql & Legajo & ","
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & "'" & apellido & "',"
            StrSql = StrSql & "'" & apellido2 & "',"
            StrSql = StrSql & "'" & nombre & "',"
            StrSql = StrSql & "'" & nombre2 & "',"
            StrSql = StrSql & proNro & ","
            StrSql = StrSql & pliqNro & ","
            StrSql = StrSql & ConvFecha(pliqFecha) & ","
            StrSql = StrSql & "'" & proDesc & "',"
            StrSql = StrSql & "'" & pliqDesc & "',"
            StrSql = StrSql & controlNull(EstrNomb1) & ","
            StrSql = StrSql & controlNull(EstrNomb2) & ","
            StrSql = StrSql & controlNull(EstrNomb3) & ","
            
            For I = 1 To 50
                StrSql = StrSql & numberForSQL(ValCols(I)) & ","
            Next
            
            StrSql = StrSql & orden & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
        End If
        
        Call borrarValores
        
        If Not rsConsult.EOF Then
            estr1Ant = rsConsult!estrnro1
            estr2Ant = rsConsult!estrnro2
            estr3Ant = rsConsult!estrnro3
            te1Ant = rsConsult!tenro1
            te2Ant = rsConsult!tenro2
            te3Ant = rsConsult!tenro3
        End If
   
   End If
Loop
        
rsConsult.Close


Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
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


Sub CargarConfiguracionReporte(ByVal Modelo As Long)

    Dim objRs As New ADODB.Recordset
    Dim StrSql As String
    
    StrSql = " SELECT * FROM repliqcols INNER JOIN repliq ON repliq.rliqnro = repliqcols.rliqnro WHERE repliq.rliqnro=" & Modelo & " ORDER BY rliqnrocol "
    
    OpenRecordset StrSql, objRs
    
    Nro_Col = 0
    CantColumnas = 0
    
    Do Until objRs.EOF
       
       Nro_Col = Nro_Col + 1
       
       CodCols(Nro_Col) = objRs!rliqval
       
       If CLng(objRs!rliqtipo) = 0 Then
          TipoCols(Nro_Col) = "CO"
       Else
          TipoCols(Nro_Col) = "AC"
       End If
       
       TitCols(Nro_Col) = objRs!rliqetiq
       NroCols(Nro_Col) = objRs!rliqnrocol
       CantColumnas = CLng(objRs!rliqcantcol)
       TituloRep = objRs!rliqdesc
    
       objRs.MoveNext
    Loop
    
    objRs.Close

End Sub

Sub GenerarEncabezadoReporte()

Dim teNomb1
Dim teNomb2
Dim teNomb3
Dim I
Dim rsConsult As New ADODB.Recordset

On Error GoTo MError

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

StrSql = " INSERT INTO rep_detcostos ( " & _
         " bpronro , Formato, rliqdesc, pliqDesde, pliqHasta, " & _
         " tedabr1 , tedabr2, tedabr3, " & _
         " coletiq1 , coletiq2, coletiq3, coletiq4, coletiq5, " & _
         " coletiq6 , coletiq7, coletiq8, coletiq9, coletiq10, " & _
         " coletiq11 , coletiq12, coletiq13, coletiq14, coletiq15, " & _
         " coletiq16 , coletiq17, coletiq18, coletiq19, coletiq20, " & _
         " coletiq21 , coletiq22, coletiq23, coletiq24, coletiq25, " & _
         " coletiq26 , coletiq27, coletiq28, coletiq29, coletiq30, " & _
         " coletiq31 , coletiq32, coletiq33, coletiq34, coletiq35, " & _
         " coletiq36 , coletiq37, coletiq38, coletiq39, coletiq40, " & _
         " coletiq41 , coletiq42, coletiq43, coletiq44, coletiq45, " & _
         " coletiq46 , coletiq47, coletiq48, coletiq49, coletiq50, " & _
         " cantcols ) VALUES ( "

StrSql = StrSql & NroProceso & ","
StrSql = StrSql & Formato & ","
StrSql = StrSql & "'" & TituloRep & "',"
StrSql = StrSql & "'" & descDesde & "',"
StrSql = StrSql & "'" & descHasta & "',"
StrSql = StrSql & controlNull(teNomb1) & ","
StrSql = StrSql & controlNull(teNomb2) & ","
StrSql = StrSql & controlNull(teNomb3) & ","

For I = 1 To 50
    StrSql = StrSql & "'" & TitCols(I) & "',"
Next

StrSql = StrSql & CantColumnas & ")"

objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "Error al cargar los datos del reporte. Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub


Sub borrarValores()
  
  Dim J
  
  For J = J To CantColumnas
    ValCols(J) = 0
  Next
End Sub 'borrarValores

Sub agregarValor(tipo, codigo, Valor)
  Dim J
  
  For J = 1 To Nro_Col
    If TipoCols(J) = tipo And CodCols(J) = codigo Then
       ValCols(CInt(NroCols(J))) = CDbl(ValCols(CInt(NroCols(J)))) + CDbl(Valor)
    End If
  Next
  
End Sub 'agregarValor(tipo, codigo, valor)

Function hayDatos()
  Dim J
  Dim Salida As Boolean
  
  Salida = False
  For J = 1 To Nro_Col
    If ValCols(J) <> 0 Then
       Salida = True
       Exit For
    End If
  Next

  hayDatos = Salida
End Function

Function igualEstr(ByVal val1, ByVal val2)
  Dim Salida As Boolean
  
  If IsNull(val1) And IsNull(val2) Then
     Salida = True
  Else
     If Not IsNull(val1) And Not IsNull(val2) Then
        Salida = (CLng(val1) = CLng(val2))
     Else
        Salida = False
     End If
  End If
  
  igualEstr = Salida

End Function


Function buscarTE(ByVal TEActual, ByVal TEDetCostos1, ByVal TEDetCostos2, ByVal TEDetCostos3, ByVal EstrNomb1 As String, ByVal EstrNomb2 As String, ByVal EstrNomb3 As String)
   Dim Salida
   
   Salida = ""
   
   If Not IsNull(TEDetCostos1) Then
      If igualEstr(TEActual, TEDetCostos1) Then
         Salida = EstrNomb1
      End If
   End If
   
   If Not IsNull(TEDetCostos2) Then
      If igualEstr(TEActual, TEDetCostos2) Then
         Salida = EstrNomb2
      End If
   End If
   
   If Not IsNull(TEDetCostos3) Then
      If igualEstr(TEActual, TEDetCostos3) Then
         Salida = EstrNomb3
      End If
   End If
   
   buscarTE = Salida
End Function
