Attribute VB_Name = "ReporteAnuGanancias"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "27/01/2011"
'Global Const UltimaModificacion = "Diego Rosso "
'Global Const UltimaModificacion1 = ""

'Global Const Version = "1.01"
'Global Const FechaModificacion = "13/02/2011"
'Global Const UltimaModificacion = "Diego Rosso "
'Global Const UltimaModificacion1 = "" ' Se guarda el log dentro de la carpeta del usuario que ejecuta el proceso al igual que en la simulacion.
                                      ' Se cambia de lugar una inicializacion de array que hacia que los rubros 7 en adelante salgan todos en cero.

Global Const Version = "1.02"
Global Const FechaModificacion = "20/10/2011"
Global Const UltimaModificacion = "Lisandro Moro"
Global Const UltimaModificacion1 = "Validacion contra sim_empleado"

'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
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


Global docum_tipo As Integer
Global docum_desc As String

Global tipoModelo As Integer

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String
Global empresa
Global emprNro
Global emprActiv
Global emprTer
Global emprDire
Global emprCuit

Global ii As Integer
Global Itemtip(1000) As String
Global Itemval(1000) As Integer
Global acuconval2(8) As String
Global ItemDesc(100) As String
Global arrConfrep(100) As String

Global Indice As Integer




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
Dim objRs2 As New ADODB.Recordset
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim proNro
Dim Ternro
'Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim acunroSueldo
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim pliqdesde As Integer
Dim pliqhasta As Integer
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden As Integer
Dim Iduser As String

'Dim ArrParametros

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
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
     'Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
     Exit Sub
    End If
    
    'obtengo solamente el iduser para generar el log.
     StrSql = " SELECT iduser FROM batch_proceso WHERE bpronro = " & NroProceso
     OpenRecordset StrSql, objRs2
     Iduser = ""
     If Not objRs2.EOF Then
         Iduser = objRs2!Iduser
     End If
     objRs2.Close
    
    Nombre_Arch = PathFLog & Iduser & "\" & "ReporteAnualGanancias" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    'Si no existe creo la carpeta donde voy a guardar el log
    If Not fs.FolderExists(PathFLog & Iduser) Then fs.CreateFolder (PathFLog & Iduser)

    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    TiempoInicialProceso = GetTickCount
    
    HuboErrores = False
    
    On Error GoTo CE
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    'FGZ - 19/01/2007 - le cambié esto
    'cantRegistros = CInt(objRs!total)
    'por
    cantRegistros = CLng(objRs!total)
    
    totalEmpleados = cantRegistros
    
    objRs.Close
      
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    Flog.writeline
    Flog.writeline "Inicio Reporte Anual Ganancias: " & Now
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

    Flog.writeline "Cambio el estado del proceso a Procesando"
    Flog.writeline
    
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
       listapronro = ArrParametros(0) 'No la uso
       
       'Obtengo el modelo a usar para obtener los datos
       tipoModelo = ArrParametros(1)
              
       'Obtengo los cortes de estructura
       tenro1 = CLng(ArrParametros(2))
       estrnro1 = CLng(ArrParametros(3))
       tenro2 = CLng(ArrParametros(4))
       estrnro2 = CLng(ArrParametros(5))
       tenro3 = CLng(ArrParametros(6))
       estrnro3 = CLng(ArrParametros(7))
       fecEstr = ArrParametros(8)
       
       'Obtengo el titulo del reporte
       tituloReporte = ArrParametros(9)
       
       pliqdesde = ArrParametros(10)
       pliqhasta = ArrParametros(11)
       
       'EMPIEZA EL PROCESO
       
      
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 114 "
       StrSql = StrSql & " ORDER BY confnrocol "
       OpenRecordset StrSql, objRs2
        
       If objRs2.EOF Then
          Flog.writeline "Debe configurar el reporte"
          Exit Sub
       End If
       objRs2.Close
     
        
       'Borro los datos del mismo proceso por si se corre mas de una vez
       StrSql = "DELETE FROM rep_simAnuGanancias_det WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       StrSql = "DELETE FROM rep_simAnuGanancias WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       
       
       'Obtengo los empleados
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'arrpronro = Split(listapronro, ",")
       
       'Obtengo los datos de la empresa
       'Call buscarDatosEmpresa(NroProceso, arrpronro(0))
    
        orden = 0
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
'          arrpronro = Split(listapronro, ",")
          EmpErrores = False
          Ternro = rsEmpl!Ternro
          orden = orden + 1
          Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
             
       
          'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
          
          cantRegistros = cantRegistros - 1
          
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
             
          objConn.Execute StrSql, , adExecuteNoRecords
          
          'Si se generaron todos los datos del empleado correctamente lo borro
          If Not EmpErrores Then
              'StrSql = " DELETE FROM batch_empleado "
              'StrSql = StrSql & " WHERE bpronro = " & NroProceso
              'StrSql = StrSql & " AND ternro = " & Ternro
    
              'objConn.Execute StrSql, , adExecuteNoRecords
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
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    Flog.writeline "================================================================="
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline "================================================================="
End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function




Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    'StrEmpl = StrEmpl & " ORDER BY estado "
    OpenRecordset StrEmpl, rsEmpl
    
    If rsEmpl.EOF Then
         Flog.writeline "No hay empleados para procesar"
    End If
End Sub

Function numberForSQL(Str)
   
    If IsEmpty(Str) Or IsNull(Str) Then
        numberForSQL = "null"
    Else
        numberForSQL = Replace(Str, ",", ".")
    End If

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function

Public Function Calcular_Edad(ByVal Fecha As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Nombre progr. :
' tipo programa : FGZ
' Descripcion   :
'...........................................................................
Dim años  As Integer

    años = Year(Date) - Year(Fecha)
    If Month(Date) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(Date) = Month(Fecha) Then
            If Day(Date) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
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



Sub generarDatosEmpleado(Ternro, descripcion, orden, pliqdesde, pliqhasta)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsultPpal As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim tenomb1
Dim tenomb2
Dim tenomb3

Dim pliqdesdemes As Byte
Dim pliqdesdeanio As Integer
Dim pliqdesdefecha As Date
Dim pliqhastames  As Byte
Dim pliqhastaanio As Integer
Dim pliqhastafecha As Date
Dim DifPeriodos As Integer
Dim i As Integer
Dim liqReal As Boolean
Dim j As Integer
Dim lstpronro As String
Dim sql As String

Dim arrItemMonto(12) As Double
Dim arrItemReal(12) As Integer
Dim arrItempliqnro(12) As Long
Dim arrEsccuota(12) As Double
Dim arrExcedente(12) As Double

Dim itemtotal As Double
Dim X As Integer
Dim objRs2 As New ADODB.Recordset
Dim objRs1 As New ADODB.Recordset
Dim l_escala_ded_porc
Dim l_rubro1 As Double
Dim l_rubro2 As Double
Dim l_rubro3 As Double
Dim l_rubro4 As Double
Dim l_rubro5 As Double
Dim l_rubro6(12) As Double
Dim l_rubro8(12) As Double
Dim l_rubro9(12) As Double
Dim l_rubro10(12) As Double

Dim l_concepto As Long
Dim l_gan_impo As Double
Dim l_imp_por_escala As Double
Dim l_ret_dev As Double
Dim l_escala_inf As Double
Dim l_escporexe As Double
Dim l_esccuota As Double


On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
tenomb1 = ""
tenomb2 = ""
tenomb3 = ""

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
Flog.writeline "Buscando datos del empleado"
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
    StrSql = StrSql & " FROM sim_empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = sim_empleado.ternro "
    StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
    Flog.writeline "Buscando datos del Sim_empleado"
    OpenRecordset StrSql, rsConsult
End If
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
   
   legajo = rsConsult!empleg
   Flog.writeline "Datos obtenidos Legajo: " & legajo & " - " & apellido & " " & apellido2 & ", " & nombre
   Flog.writeline "-------------------------------------------------------------------------------------------"
Else
   Flog.writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro = " & pliqdesde

Flog.writeline "Buscando datos del periodo desde"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   
   pliqdesdemes = rsConsult!pliqmes
   pliqdesdeanio = rsConsult!pliqanio
   pliqdesdefecha = rsConsult!pliqdesde
   
Else
   Flog.writeline "Error al obtener los datos del  periodo desde"
'   GoTo MError
End If

rsConsult.Close

StrSql = " SELECT * "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro = " & pliqhasta

Flog.writeline "Buscando datos del periodo hasta"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   
   pliqhastames = rsConsult!pliqmes
   pliqhastaanio = rsConsult!pliqanio
   pliqhastafecha = rsConsult!pliqdesde
   
Else
   Flog.writeline "Error al obtener los datos del  periodo hasta"
'   GoTo MError
End If

rsConsult.Close

Flog.writeline "Comienza transaccion"
        
MyBeginTrans

Flog.writeline "Guardo Cabecera para el empleado"

'------------------------------------------------------------------
'Guardo la cabecera
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_simAnuGanancias "
StrSql = StrSql & " (bpronro , ternro, apellido,apellido2, nombre, nombre2,"
StrSql = StrSql & " legajo , pliqdesde, pliqhasta, pronro,"
StrSql = StrSql & " orden , descabr) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & "," & legajo
StrSql = StrSql & "," & pliqdesde
StrSql = StrSql & "," & pliqhasta
StrSql = StrSql & ",'" & 0 & "'"
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Left(descripcion, 200) & "'"
StrSql = StrSql & ")"

objConn.Execute StrSql, , adExecuteNoRecords


'Cantidad de periodos
DifPeriodos = DateDiff("m", pliqdesdefecha, pliqhastafecha)

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 114 "
StrSql = StrSql & " ORDER BY confnrocol "
OpenRecordset StrSql, objRs2
        
If objRs2.EOF Then
   Flog.writeline "Debe configurar el reporte"
   Exit Sub
End If
Flog.writeline "Obtengo los datos del confrep"
       
Indice = 0
l_rubro1 = 0
l_rubro2 = 0
l_rubro3 = 0
l_rubro4 = 0
l_rubro5 = 0
    
For i = 0 To 12
    l_rubro6(i) = 0
    arrEsccuota(i) = 0
    arrExcedente(i) = 0
    l_rubro8(i) = 0
    l_rubro9(i) = 0
    l_rubro10(i) = 0
Next i
      
    
Do Until objRs2.EOF
'         ItemDesc(objRs2!confnrocol) = Left(objRs2!confetiq, 30)
'         Itemtip(objRs2!confnrocol) = objRs2!conftipo
'         If Itemtip(objRs2!confnrocol) = "ITM" Then
'             Itemval(objRs2!confnrocol) = objRs2!confval
'         Else
'             Itemval(objRs2!confnrocol) = objRs2!confval2
'         End If
'          objRs2.MoveNext


    'Inicializo arrays
    For i = 0 To 12
        arrItemMonto(i) = 0
        arrItemReal(i) = -1
        arrItempliqnro(i) = 0
       ' l_rubro6(i) = 0
       ' arrEsccuota(i) = 0
       ' arrExcedente(i) = 0
       ' l_rubro8(i) = 0
       ' l_rubro9(i) = 0
       ' l_rubro10(i) = 0
    Next i
    
    itemtotal = 0

    
    
    For i = 0 To DifPeriodos
        Flog.writeline "Buscando datos de liquidaciones"
        
        '------------------------------------------------------------------
        'Busco datos de cabeceras de liquidacion
        '------------------------------------------------------------------
        StrSql = " SELECT c.pronro, p2.pliqnro,  p.profecpago FROM cabliq c"
        StrSql = StrSql & " INNER JOIN proceso p ON p.pronro = c.pronro"
        StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", i, pliqdesdefecha))
        StrSql = StrSql & " AND c.empleado= " & Ternro
        StrSql = StrSql & " AND p.tprocnro = " & tipoModelo
        
        OpenRecordset StrSql, rsConsultPpal
        
        If Not rsConsultPpal.EOF Then
           liqReal = True
           arrItemReal(i) = -1
        Else
           arrItemReal(i) = 0
           liqReal = False
           rsConsultPpal.Close
           
            StrSql = " SELECT c.pronro, p2.pliqnro,  p.profecpago FROM sim_cabliq c"
            StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
            StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
            StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", i, pliqdesdefecha))
            StrSql = StrSql & " AND c.empleado= " & Ternro
            StrSql = StrSql & " AND p.tprocnro = " & tipoModelo
            OpenRecordset StrSql, rsConsultPpal
        End If
        
        'Me guardo el pliqnro
        If Not rsConsultPpal.EOF Then
           arrItempliqnro(i) = rsConsultPpal!pliqnro
        End If
        'armar lst de cliqnro
        
        'lstpronro = "0"
        
        'Do While Not rsConsult.EOF
        '    lstpronro = lstpronro & "," & rsConsult!proNro
        '    rsConsult.MoveNext
        'Loop
        
        'If lstpronro <> "0" Then
        If Not rsConsultPpal.EOF Then
             If objRs2!conftipo = "ITM" Then
                
                StrSql = " SELECT sum(tg.monto) monto "
                If liqReal Then
                    StrSql = StrSql & " FROM traza_gan_item_top tg "
                Else
                    StrSql = StrSql & " FROM sim_traza_gan_item_top tg "
                End If
                StrSql = StrSql & " WHERE ternro = " & Ternro
                StrSql = StrSql & " and itenro= " & objRs2!confval
                'StrSql = StrSql & " AND pronro in ( " & lstpronro & " )"
                StrSql = StrSql & " AND pronro = " & rsConsultPpal!proNro
                '---LOG---
                'Flog.writeline "Procesando acumulador: " & Itemval(j) & "para el periodo: " & arrItempliqnro(i) & "(" & DateAdd("m", i, pliqdesdefecha) & ")"
                OpenRecordset StrSql, rsConsult
                If Not rsConsult.EOF Then
                    arrItemMonto(i) = IIf(IsNull(rsConsult!Monto), 0, Abs(rsConsult!Monto))
                    'itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                End If
                rsConsult.Close
                
                If liqReal And arrItemMonto(i) = 0 Then
                    'Si tiene cabeceras en real pero igual no encontro ninguna liquidada
                    'o valores en la misma fuerzo la busqueda en sim
                    arrItemReal(i) = 0

                    StrSql = " SELECT c.pronro, p2.pliqnro FROM cabliq c"
                    StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
                    StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
                    StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", i, pliqdesdefecha))
                    StrSql = StrSql & " AND c.empleado= " & Ternro
                    OpenRecordset StrSql, rsConsult

'                    lstpronro = "0"
'
'                    Do While Not rsConsult.EOF
'                        lstpronro = lstpronro & "," & rsConsult!cliqnro
'                        rsConsult.MoveNext
'                    Loop

'                    If lstpronro <> "0" Then
                    If Not rsConsult.EOF Then
                        StrSql = " SELECT sum(tg.monto) monto  "
                        StrSql = StrSql & " FROM sim_traza_gan_item_top tg "
                        StrSql = StrSql & " WHERE ternro = " & Ternro
                        StrSql = StrSql & " and itenro= " & objRs2!confval
                        'StrSql = StrSql & " AND pronro in ( " & lstpronro & " )"
                        StrSql = StrSql & " AND pronro = " & rsConsultPpal!proNro
                        OpenRecordset StrSql, rsConsult
                        If Not rsConsult.EOF Then
                            arrItemMonto(i) = IIf(IsNull(rsConsult!Monto), 0, Abs(rsConsult!Monto))
                            'itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                        End If
                        rsConsult.Close
                    End If
                End If 'If liqReal And arrItemMonto(i) = 0 Then
                
                
                
               Select Case objRs2!confnrocol

                Case 2, 3, 4
                         l_rubro1 = l_rubro1 + arrItemMonto(i)
                Case 5 To 14
                         l_rubro2 = l_rubro2 + arrItemMonto(i)
                Case 15
                         l_rubro3 = (l_rubro1 - l_rubro2) - arrItemMonto(i)
                Case 16
                         l_rubro3 = l_rubro3 - arrItemMonto(i)
                         
                Case 17
                          l_rubro5 = l_rubro3 - arrItemMonto(i)
                         
                Case 18

                       
                   '-----------------------------------------------------------------------
                   ' Busco los datos de la escala de deduccion
                   '-----------------------------------------------------------------------

                    If CLng(Year(DateAdd("m", i, pliqdesdefecha))) > 2000 And CDbl(l_rubro5) > 0 Then

                        StrSql = " SELECT * "
                        StrSql = StrSql & " FROM escala_ded "
                       ' Si es una liquidacion final la considero a diciembre
                        If CLng(tipoModelo) = 5 Or CLng(tipoModelo) = 10 Then
                           StrSql = StrSql & " WHERE esd_topeinf <= " & Int(l_rubro5)
                           StrSql = StrSql & "   AND esd_topesup >= " & Int(l_rubro5)
                        Else
                           StrSql = StrSql & " WHERE esd_topeinf <= " & Int((l_rubro5 / Month(rsConsultPpal!profecpago)) * 12)
                           StrSql = StrSql & "   AND esd_topesup >= " & Int((l_rubro5 / Month(rsConsultPpal!profecpago)) * 12)
                        End If
                        OpenRecordset StrSql, objRs1

                        l_escala_ded_porc = 0

                        If Not objRs1.EOF Then
                            l_escala_ded_porc = objRs1!esd_porcentaje
                        End If
                        
                        l_rubro6(i) = l_rubro6(i) + arrItemMonto(i)
                        
                        arrItemMonto(i) = arrItemMonto(i) * (CDbl(l_escala_ded_porc) / 100)
                        
                    End If
              Case 19, 20, 21
                        l_rubro6(i) = l_rubro6(i) + arrItemMonto(i)
                        
                        arrItemMonto(i) = arrItemMonto(i) * (CDbl(l_escala_ded_porc) / 100)
              Case 22
                        l_rubro6(i) = l_rubro6(i) + arrItemMonto(i)
                        
                        arrItemMonto(i) = arrItemMonto(i) * (CDbl(l_escala_ded_porc) / 100)
                        
                        l_rubro6(i) = l_rubro6(i) * CDbl(l_escala_ded_porc) / 100
                        
              End Select
                
             Else
                 'No es tipo ITM
                 'Procesar CO
                 'l_concepto = objRs2!confval
                 StrSql = " SELECT * "
                 StrSql = StrSql & " FROM concepto "
                 StrSql = StrSql & " WHERE conccod = " & objRs2!confval
                 OpenRecordset StrSql, objRs1
    
                 If Not objRs1.EOF Then
                    l_concepto = objRs1!concnro
                 End If
        
        
                '-----------------------------------------------------------------------
                'Busco los datos de la tabla traza_gan
                '-----------------------------------------------------------------------
                
                StrSql = " SELECT * "
                If liqReal Then
                    StrSql = StrSql & " FROM traza_gan "
                Else
                    StrSql = StrSql & " FROM sim_traza_gan "
                End If
                StrSql = StrSql & " WHERE ternro  = " & Ternro
                StrSql = StrSql & "   AND pliqnro = " & arrItempliqnro(i)
                StrSql = StrSql & "   AND concnro = " & l_concepto
                StrSql = StrSql & "   AND pronro  = " & rsConsultPpal!proNro
    
                OpenRecordset StrSql, objRs1
    
                l_gan_impo = 0
                l_imp_por_escala = 0
                l_ret_dev = 0
    
                If Not objRs1.EOF Then
                    l_gan_impo = objRs1!ganimpo
                    l_imp_por_escala = objRs1!imp_deter
                    l_ret_dev = objRs1!saldo
                    l_rubro8(i) = l_imp_por_escala
                    l_rubro9(i) = l_imp_por_escala - l_ret_dev
                    l_rubro10(i) = l_ret_dev
                End If
                
                '-----------------------------------------------------------------------
                ' Busco los datos de la escala
                '-----------------------------------------------------------------------
                
                If l_gan_impo >= 0 Then
                
                    StrSql = " SELECT * "
                    StrSql = StrSql & " FROM escala "
                    StrSql = StrSql & " WHERE escinf <= " & l_gan_impo
                    StrSql = StrSql & "   AND escsup >= " & l_gan_impo
                    StrSql = StrSql & "   AND escano = " & Year(rsConsultPpal!profecpago)
                    If CLng(tipoModelo) = 5 Or CLng(tipoModelo) = 10 Then
                    'If l_final = -1 Then
                       StrSql = StrSql & "   AND escmes = 12 "
                    Else
                       'l_sql = l_sql & "   AND escmes = " & l_pliqmes
                       StrSql = StrSql & "   AND escmes = " & Month(rsConsultPpal!profecpago)
                    End If
                
                     OpenRecordset StrSql, objRs1
                
                    l_escala_inf = 0
                    l_escporexe = 0
                    l_esccuota = 0
                    
                    If Not objRs1.EOF Then
                        l_escala_inf = objRs1!escinf
                        l_escporexe = objRs1!escporexe
                        l_esccuota = objRs1!esccuota
                        arrEsccuota(i) = objRs1!esccuota
                    End If
                    
                End If
                
                arrExcedente(i) = ((CDbl(l_gan_impo) - CDbl(l_escala_inf)) * CDbl(l_escporexe)) / 100
                 
             End If
           
        End If
        
       
    Next i 'Proximo periodo
    'grabo det
    Flog.writeline "Grabando detalle para la columna: " & objRs2!confnrocol
    StrSql = " INSERT INTO rep_simAnuGanancias_det "
    StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
    StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
    StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
    StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
    StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
    StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
    StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
    StrSql = StrSql & " ) VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & ",'" & objRs2!conftipo & "'"
    StrSql = StrSql & "," & objRs2!confval
    StrSql = StrSql & ",'" & objRs2!confetiq & "'"
    StrSql = StrSql & "," & Indice
    For X = 0 To 11
        StrSql = StrSql & "," & arrItemReal(X)
        StrSql = StrSql & "," & arrItempliqnro(X)
        StrSql = StrSql & "," & numberForSQL(arrItemMonto(X))
    Next X
    StrSql = StrSql & "," & itemtotal
    
    StrSql = StrSql & ")"
                
    objConn.Execute StrSql, , adExecuteNoRecords
    Indice = Indice + 1
    objRs2.MoveNext
Loop
objRs2.Close

'Guardo datos especificos ya que no puedo generarlos en el asp por falta de informacion.
 Flog.writeline "Grabando detalle para la columna: total 6"
        StrSql = " INSERT INTO rep_simAnuGanancias_det "
        StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
        StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
        StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
        StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
        StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
        StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
        StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
        StrSql = StrSql & " ) VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Ternro
        StrSql = StrSql & ",'TOT'"
        StrSql = StrSql & "," & 0
        StrSql = StrSql & ",'Total 6'"
        StrSql = StrSql & "," & Indice
        For X = 0 To 11
            StrSql = StrSql & "," & CInt(liqReal)
            StrSql = StrSql & "," & arrItempliqnro(X)
            StrSql = StrSql & "," & numberForSQL(l_rubro6(X))
        Next X
        StrSql = StrSql & "," & itemtotal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

Indice = Indice + 1

Flog.writeline "Grabando detalle para la columna: EscCuota"
        StrSql = " INSERT INTO rep_simAnuGanancias_det "
        StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
        StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
        StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
        StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
        StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
        StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
        StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
        StrSql = StrSql & " ) VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Ternro
        StrSql = StrSql & ",'TOT'"
        StrSql = StrSql & "," & 0
        StrSql = StrSql & ",'Valor fijo de tabla'"
        StrSql = StrSql & "," & Indice
        For X = 0 To 11
            StrSql = StrSql & "," & CInt(liqReal)
            StrSql = StrSql & "," & arrItempliqnro(X)
            StrSql = StrSql & "," & numberForSQL(arrEsccuota(X))
        Next X
        StrSql = StrSql & "," & itemtotal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

Indice = Indice + 1

Flog.writeline "Grabando detalle para la columna: % sobre el excedente"
        StrSql = " INSERT INTO rep_simAnuGanancias_det "
        StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
        StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
        StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
        StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
        StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
        StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
        StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
        StrSql = StrSql & " ) VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Ternro
        StrSql = StrSql & ",'TOT'"
        StrSql = StrSql & "," & 0
        StrSql = StrSql & ",'% sobre el excedente'"
        StrSql = StrSql & "," & Indice
        For X = 0 To 11
            StrSql = StrSql & "," & CInt(liqReal)
            StrSql = StrSql & "," & arrItempliqnro(X)
            StrSql = StrSql & "," & numberForSQL(arrExcedente(X))
        Next X
        StrSql = StrSql & "," & itemtotal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

Indice = Indice + 1

Flog.writeline "Grabando detalle para la columna: RUBRO 8 - Total impuesto"
        StrSql = " INSERT INTO rep_simAnuGanancias_det "
        StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
        StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
        StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
        StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
        StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
        StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
        StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
        StrSql = StrSql & " ) VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Ternro
        StrSql = StrSql & ",'TOT'"
        StrSql = StrSql & "," & 0
        StrSql = StrSql & ",'RUBRO 8 - Total impuesto'"
        StrSql = StrSql & "," & Indice
        For X = 0 To 11
            StrSql = StrSql & "," & CInt(liqReal)
            StrSql = StrSql & "," & arrItempliqnro(X)
            StrSql = StrSql & "," & numberForSQL(l_rubro8(X))
        Next X
        StrSql = StrSql & "," & itemtotal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

Indice = Indice + 1

Flog.writeline "Grabando detalle para la columna: RUBRO 9 - Impuesto retenido"
        StrSql = " INSERT INTO rep_simAnuGanancias_det "
        StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
        StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
        StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
        StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
        StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
        StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
        StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
        StrSql = StrSql & " ) VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Ternro
        StrSql = StrSql & ",'TOT'"
        StrSql = StrSql & "," & 0
        StrSql = StrSql & ",'RUBRO 9 - Impuesto retenido'"
        StrSql = StrSql & "," & Indice
        For X = 0 To 11
            StrSql = StrSql & "," & CInt(liqReal)
            StrSql = StrSql & "," & arrItempliqnro(X)
            StrSql = StrSql & "," & numberForSQL(l_rubro9(X))
        Next X
        StrSql = StrSql & "," & itemtotal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
Indice = Indice + 1

Flog.writeline "Grabando detalle para la columna: RUBRO 10 - Saldo del impuesto"
        StrSql = " INSERT INTO rep_simAnuGanancias_det "
        StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
        StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
        StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
        StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
        StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
        StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
        StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal "
        StrSql = StrSql & " ) VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Ternro
        StrSql = StrSql & ",'TOT'"
        StrSql = StrSql & "," & 0
        StrSql = StrSql & ",'RUBRO 10 - Saldo del impuesto'"
        StrSql = StrSql & "," & Indice
        For X = 0 To 11
            StrSql = StrSql & "," & CInt(liqReal)
            StrSql = StrSql & "," & arrItempliqnro(X)
            StrSql = StrSql & "," & numberForSQL(l_rubro10(X))
        Next X
        StrSql = StrSql & "," & itemtotal
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
Flog.writeline "Commit trasaccion"
MyCommitTrans



Siguiente:

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & legajo & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    MyRollbackTrans
    Exit Sub
End Sub
