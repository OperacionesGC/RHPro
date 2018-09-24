Attribute VB_Name = "ReporteAnualConcAcu"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "26/12/2010"
'Global Const UltimaModificacion = "Diego Rosso "
'Global Const UltimaModificacion1 = ""
'Basado en la version 1.03 03/08/2009 -  'Encriptacion de string connection

'Global Const Version = "1.01"
'Global Const FechaModificacion = "20/10/2011"
'Global Const UltimaModificacion = "Lisandro Moro"
'Global Const UltimaModificacion1 = "Se unifico con Simulacion - Gestion Presupuestaria."

'Global Const Version = "1.02"
'Global Const FechaModificacion = "03/11/2011"
'Global Const UltimaModificacion = "Sebastian Stremel"
'Global Const UltimaModificacion1 = "se modifico para realizar comparaciones entre simulacion e historicos"

'Global Const Version = "1.03"
'Global Const FechaModificacion = "29/11/2012"
'Global Const UltimaModificacion = "Deluchi Ezequiel"
'Global Const UltimaModificacion1 = "se modifico las consultas de los historicos, para que filtre por nro de historico"

'Global Const Version = "1.04"
'Global Const FechaModificacion = "23/10/2012"
'Global Const UltimaModificacion = "Lisandro Moro"
'Global Const UltimaModificacion1 = "Si se busca en un proceso y no se encuentran resultados, no se fuerza a la busqueda en las simulaciones."

'Global Const Version = "1.05"
'Global Const FechaModificacion = "08/11/2012"
'Global Const UltimaModificacion = "Lisandro Moro"
'Global Const UltimaModificacion1 = "Se agrega la comparacion entre liquidacion y simulacion."

Global Const Version = "1.06"
Global Const FechaModificacion = "28/11/2014"
Global Const UltimaModificacion = "Ruiz Miriam"
Global Const UltimaModificacion1 = "Se modificaron los modelos de comparación, ahora permite comparar liquidación tanto con simulación como con históricos"

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
Global Itemval(1000) As String
Global acuconval2(8) As String
Global ItemDesc(1000) As String
Global Indice As Integer
'variables nuevas 31/10/2011
Global pliqdesde2 As Integer
Global pliqhasta2 As Integer
Global historico1
Global historico2
Global listaelegida2
Global desde2
Global hasta2
Global aprobado2





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
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim pliqdesde As Integer
Dim pliqhasta As Integer
Dim parametros As String
Dim ArrParametros
Dim TipoReporte As String
Dim strTempo As String
Dim orden As Integer



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
    
    Nombre_Arch = PathFLog & "ReporteAnualAC-CO" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    TiempoInicialProceso = GetTickCount
        
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
     Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
     Exit Sub
    End If
    
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
    Flog.writeline "Inicio Reporte Anual Simulación: " & Now
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
       'Lisandro Moro - Se cambio el tipo de modelo por el tipo de reporte, liq o sim, etc.
       '1-liq/rep_anual_acu_con_sim_00.asp - liq
       '2-sim/rep_anual_acu_con_sim_00.asp - liq
       '3-sim/rep_anual_acu_con_sim_00.asp - liq vs sim
       '4-sim/rep_comp_anual_acu_con_sim_00.asp - sim vs his
       
        listapronro = ArrParametros(0) 'No la uso
        
       'Obtengo el modelo a usar para obtener los datos
       tipoModelo = ArrParametros(1)
       If CStr(tipoModelo) = CStr(0) Then
        tipoModelo = 1
       End If
       'If Len(tipoModelo) > 2 Then
       '     tipoModelo = 4 'es el comparativo de sim vs his que selecciona procesos que no los usa.
       'End If
                     
                     
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
       
       '31/10/2011 sebastian stremel
       If UBound(ArrParametros) > 12 Then
         historico1 = ArrParametros(12)
         historico2 = ArrParametros(13)
         listaelegida2 = ArrParametros(14)
         aprobado2 = ArrParametros(15)
         pliqdesde2 = ArrParametros(16)
         pliqhasta2 = ArrParametros(17)
         desde2 = ArrParametros(18)
         hasta2 = ArrParametros(19)
       End If
       
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep el numero de cuenta que se va a usar para
       ' buscar el valor del sueldo
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 302 "
       Flog.writeline StrSql
       OpenRecordset StrSql, objRs2
        
       If objRs2.EOF Then
          Flog.writeline "Debe configurar el reporte con los acumuladores o conceptos a mostrar"
          Exit Sub
       End If
       Flog.writeline "Obtengo los datos del confrep"
       
       Indice = 1
       
       Do Until objRs2.EOF
         ItemDesc(Indice) = objRs2!confetiq
         Itemtip(Indice) = objRs2!conftipo
         If Itemtip(Indice) = "AC" Then
             Itemval(Indice) = objRs2!confval
         Else
             Itemval(Indice) = objRs2!confval2
         End If
    
          objRs2.MoveNext
          Indice = Indice + 1
       Loop
       objRs2.Close
     
      ' ReDim Preserve itemtip(indice)
      ' ReDim Preserve itemval(indice)
      '' ReDim Preserve itemmonto(indice)
       
           
        
       'Borro los datos del mismo proceso por si se corre mas de una vez
       StrSql = "DELETE FROM rep_simAnuAcuCon_det WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       StrSql = "DELETE FROM rep_simAnuAcuCon WHERE bpronro = " & NroProceso
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
          If tipoModelo = 1 Then 'liquidacion sin comparacion
            Call generarDatosEmpleadoLiq(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
          End If
          If tipoModelo = 2 Then 'simulacion sin comparacion
            Call generarDatosEmpleadoSim(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
          End If
          If tipoModelo = 3 Then 'histórico sin comparacion
            Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
          End If
          If tipoModelo = 4 Then 'liquidacion vs simulacion
            Call generarDatosEmpleadoLiq(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
            Call generarDatosEmpleadoSim(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
 
          End If
          If tipoModelo = 5 Then 'liquidacion vs histórico
            Call generarDatosEmpleadoLiq(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
            Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
          End If
          If tipoModelo = 6 Then 'simulacion vs liquidacion
            Call generarDatosEmpleadoSim(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
            Call generarDatosEmpleadoLiq(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
          End If
          If tipoModelo = 7 Then 'simulacion vs historico
            Call generarDatosEmpleadoSim(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
            Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
          End If
          If tipoModelo = 8 Then 'historico vs liquidacion
           
            Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
            Call generarDatosEmpleadoLiq(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
          End If
          If tipoModelo = 9 Then 'historico vs simulacion
            Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
            Call generarDatosEmpleadoSim(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
 
          End If
          If tipoModelo = 10 Then 'historico vs historico
           Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta, -1)
           Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2, 0)
          End If
          
          
          ' comento desde  acá MR
          '  If tipoModelo = 1 Then 'liquidacion
          '      Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
          '  End If
          '  If tipoModelo = 2 Then 'no compara
          '      Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
            '    'Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
           ' End If
           ' If tipoModelo = 3 Then 'si compara
           '     Call generarDatosEmpleadoLiq(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
           '     Call generarDatosEmpleadoSim(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
        '    End If
            
       '     If tipoModelo = 4 Then 'si compara
       '         'agrego sebastian 31/10/2011
       '           If historico2 = "" Then ' si no hay comparacion funciona como antes
       '               If historico1 = "" Then '-> ""*0_lm
       '                   Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta) 'sin comparacion eligio sim
       '               Else
       '                   Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta) 'sin comparacion eligio his
       '               End If
       '           Else 'hay comparacion
       '               If historico1 = 0 And historico2 = 0 Then ' compara sim con sim
       '                   Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
       '                   Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2)
       '               Else
       '                   If historico1 = 0 And historico2 <> 0 Then 'compara sim con historico
       '                       Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
       '                       Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2)
       '                   Else
       '                       If historico1 <> 0 And historico2 = 0 Then 'compara his con sim
       '                           Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
       '                           Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2)
       '                       Else
       '                           If historico1 <> 0 And historico2 <> 0 Then 'compara his con his
       '                               Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
       '                               Call generarDatosEmpleadoHis(Ternro, tituloReporte, orden, pliqdesde2, pliqhasta2)
       '                           End If
       '                       End If
       '                   End If
       '               End If
       '           End If
       '     End If
           ' comento hasta acá MR
            
            
          'hasta aca
          'Call generarDatosEmpleado(Ternro, tituloReporte, orden, pliqdesde, pliqhasta)
             
       
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

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
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
Dim I As Integer
Dim liqReal As Boolean
Dim j As Integer
Dim lstcliqnro As String
Dim sql As String
Dim arrItemMonto(12) As Double
Dim arrItemReal(12) As Integer
Dim arrItempliqnro(12) As Long
Dim itemtotal As Double
Dim X As Integer

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
Flog.writeline StrSql
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontro el empleado."
    StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
    StrSql = StrSql & " FROM sim_empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = sim_empleado.ternro "
    StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
    Flog.writeline "Buscando datos del SIM empleado."
    Flog.writeline StrSql
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
   
   Legajo = rsConsult!empleg
   Flog.writeline "Datos obtenidos Legajo: " & Legajo
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
      Flog.writeline StrSql
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
       Flog.writeline StrSql
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

 Flog.writeline "Guardo Cabecera para el empleado"

'------------------------------------------------------------------
'Guardo la cabecera
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_simAnuAcuCon "
StrSql = StrSql & " (bpronro , ternro, apellido,apellido2, nombre, nombre2,"
StrSql = StrSql & " legajo , pliqdesde, pliqhasta, pronro,"
StrSql = StrSql & " orden , descabr, historico) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqdesde
StrSql = StrSql & "," & pliqhasta
StrSql = StrSql & ",'" & 0 & "'"
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Left(descripcion, 200) & "'"
StrSql = StrSql & ",1"
'For ii = 1 To 8
'    StrSql = StrSql & "," & numberForSQL(acuconmonto(ii))
'    StrSql = StrSql & ",'" & acuconetiq(ii) & "'"
'Next
StrSql = StrSql & ")"
If Legajo > 0 Then
    Flog.writeline StrSql
   objConn.Execute StrSql, , adExecuteNoRecords
End If

'Cantidad de periodos
DifPeriodos = DateDiff("m", pliqdesdefecha, pliqhastafecha)


'Ciclo por columna(ac-co)
For j = 1 To Indice - 1
    
    'Inicializo arrays
    For I = 0 To 12
        arrItemMonto(I) = 0
        arrItemReal(I) = -1
        arrItempliqnro(I) = 0
    Next I
    
    itemtotal = 0
    
    'Ciclo por periodo
    For I = 0 To DifPeriodos
        Flog.writeline "Buscando datos de cabeceras de liquidacion para el empleado"
        
        '------------------------------------------------------------------
        'Busco datos de cabeceras de liquidacion
        '------------------------------------------------------------------
        StrSql = " SELECT c.cliqnro, p2.pliqnro FROM cabliq c"
        StrSql = StrSql & " INNER JOIN proceso p ON p.pronro = c.pronro"
        StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        StrSql = StrSql & " AND c.empleado= " & Ternro
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           liqReal = True
           arrItemReal(I) = -1
        Else
           arrItemReal(I) = 0
           liqReal = False
           rsConsult.Close
           
           StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_cabliq c"
           StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
           StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
           StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
           StrSql = StrSql & " AND c.empleado= " & Ternro
           Flog.writeline StrSql
           OpenRecordset StrSql, rsConsult
        End If
        
        'Me guardo el pliqnro
        If Not rsConsult.EOF Then
           arrItempliqnro(I) = rsConsult!pliqnro
        End If
        'armar lst de cliqnro
        
        lstcliqnro = "0"
        
        Do While Not rsConsult.EOF
            lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
            rsConsult.MoveNext
        Loop
        
        If lstcliqnro <> "0" Then
             If Itemtip(j) = "AC" Then
                
                StrSql = " SELECT sum(almonto) monto "
                If liqReal Then
                    StrSql = StrSql & " FROM acu_liq"
                Else
                    StrSql = StrSql & " FROM sim_acu_liq"
                End If
                StrSql = StrSql & " WHERE acunro = " & Itemval(j)
                StrSql = StrSql & " AND cliqnro in ( " & lstcliqnro & " )"
                '---LOG---
                Flog.writeline "Procesando acumulador: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                Flog.writeline StrSql
                OpenRecordset StrSql, rsConsult
                If Not rsConsult.EOF Then
                    arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                    itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                End If
                rsConsult.Close
               'Lisandro Moro - no fuerzo nada
                If liqReal And arrItemMonto(I) = 0 Then
                    'Si tiene cabeceras en real pero igual no encontro ninguna liquidada
                    'o valores en la misma fuerzo la busqueda en sim
                    arrItemReal(I) = 0

                    StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_cabliq c"
                    StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
                    StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
                    StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
                    StrSql = StrSql & " AND c.empleado= " & Ternro
                    Flog.writeline StrSql
                    OpenRecordset StrSql, rsConsult

                    lstcliqnro = "0"

                    Do While Not rsConsult.EOF
                        lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
                        rsConsult.MoveNext
                    Loop

                    If lstcliqnro <> "0" Then

                        StrSql = " SELECT sum(almonto) monto "
                        StrSql = StrSql & " FROM sim_acu_liq"
                        StrSql = StrSql & " WHERE acunro = " & Itemval(j)
                        StrSql = StrSql & " AND cliqnro in ( " & lstcliqnro & " )"
                        Flog.writeline StrSql
                        OpenRecordset StrSql, rsConsult
                        If Not rsConsult.EOF Then
                            arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                            itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                        End If
                        rsConsult.Close
                    End If
                End If
            ElseIf Itemtip(j) = "CO" Then
                     StrSql = " SELECT sum(d.dlimonto) monto "
                     If liqReal Then
                        StrSql = StrSql & " FROM detliq d "
                     Else
                        StrSql = StrSql & " FROM sim_detliq d"
                     End If
                     StrSql = StrSql & " INNER JOIN concepto c ON c.concnro = d.concnro AND  c.conccod = '" & Itemval(j) & "'"
                     StrSql = StrSql & " WHERE cliqnro in ( " & lstcliqnro & " )"
                     '---LOG---
                     Flog.writeline "Procesando Concepto: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                     Flog.writeline StrSql
                     OpenRecordset StrSql, rsConsult
                     If Not rsConsult.EOF Then
                        arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                        itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                     End If
                     rsConsult.Close
                    'Lisandro Moro - no fuerzo nada
                     If liqReal And arrItemMonto(I) = 0 Then
                        'Si tiene cabeceras en real pero igual no encontro ninguna liquidada
                        'o valores en la misma fuerzo la busqueda en sim
                        arrItemReal(I) = 0

                        StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_cabliq c"
                        StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
                        StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
                        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
                        StrSql = StrSql & " AND c.empleado= " & Ternro
                        Flog.writeline StrSql
                        OpenRecordset StrSql, rsConsult

                        lstcliqnro = "0"

                        Do While Not rsConsult.EOF
                            lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
                            rsConsult.MoveNext
                        Loop

                        If lstcliqnro <> "0" Then

                            StrSql = " SELECT sum(d.dlimonto) monto "
                            StrSql = StrSql & " FROM sim_detliq d"
                            StrSql = StrSql & " INNER JOIN concepto c ON c.concnro = d.concnro AND  c.conccod = " & Itemval(j)
                            StrSql = StrSql & " WHERE cliqnro in ( " & lstcliqnro & " )"
                            Flog.writeline StrSql
                            OpenRecordset StrSql, rsConsult
                            If Not rsConsult.EOF Then
                                arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                                itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                            End If
                            rsConsult.Close
                        End If
                    End If ' If liqReal And arrItemMonto(i) = 0 Then
                Else
                    Flog.writeline "El tipo de columna debe ser AC o CO no puede ser: " & Itemtip(j)
            End If
           
        End If
    Next I 'Proximo periodo
    
    
    'grabo det
    
    Flog.writeline "Grabando detalle para la columna: " & j
    StrSql = " INSERT INTO rep_simAnuAcuCon_det "
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
    StrSql = StrSql & ",'" & Itemtip(j) & "'"
    StrSql = StrSql & "," & Itemval(j)
    StrSql = StrSql & ",'" & ItemDesc(j) & "'"
    StrSql = StrSql & "," & j
    For X = 0 To 11
        StrSql = StrSql & "," & arrItemReal(X)
        StrSql = StrSql & "," & arrItempliqnro(X)
        StrSql = StrSql & "," & numberForSQL(arrItemMonto(X))
    Next X
    StrSql = StrSql & "," & itemtotal
    
    StrSql = StrSql & ")"
    If Itemtip(j) = "CO" Or Itemtip(j) = "AC" Then
            Flog.writeline StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
    End If
          
Next j 'Proximo ac o co



Siguiente:

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

Sub generarDatosEmpleadoLiq(Ternro, descripcion, orden, pliqdesde, pliqhasta, historial)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
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
Dim I As Integer
Dim liqReal As Boolean
Dim j As Integer
Dim lstcliqnro As String
Dim sql As String
Dim arrItemMonto(12) As Double
Dim arrItemReal(12) As Integer
Dim arrItempliqnro(12) As Long
Dim itemtotal As Double
Dim X As Integer

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
Flog.writeline StrSql

OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontro el empleado."
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
   
   Legajo = rsConsult!empleg
   Flog.writeline "Datos obtenidos Legajo: " & Legajo
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
Flog.writeline StrSql
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
 Flog.writeline StrSql
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

 Flog.writeline "Guardo Cabecera para el empleado"

'------------------------------------------------------------------
'Guardo la cabecera
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_simAnuAcuCon "
StrSql = StrSql & " (bpronro , ternro, apellido,apellido2, nombre, nombre2,"
StrSql = StrSql & " legajo , pliqdesde, pliqhasta, pronro,"
StrSql = StrSql & " orden , descabr, historico) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqdesde
StrSql = StrSql & "," & pliqhasta
StrSql = StrSql & ",'" & 0 & "'"
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Left(descripcion, 200) & "'"
StrSql = StrSql & ",2"
'For ii = 1 To 8
'    StrSql = StrSql & "," & numberForSQL(acuconmonto(ii))
'    StrSql = StrSql & ",'" & acuconetiq(ii) & "'"
'Next
StrSql = StrSql & ")"
If Legajo > 0 Then
    Flog.writeline StrSql
    objConn.Execute StrSql, , adExecuteNoRecords
End If


'Cantidad de periodos
DifPeriodos = DateDiff("m", pliqdesdefecha, pliqhastafecha)


'Ciclo por columna(ac-co)
For j = 1 To Indice - 1
    
    'Inicializo arrays
    For I = 0 To 12
        arrItemMonto(I) = 0
        arrItemReal(I) = -1
        arrItempliqnro(I) = 0
    Next I
    
    itemtotal = 0
    
    'Ciclo por periodo
    For I = 0 To DifPeriodos
        Flog.writeline "Buscando datos de cabeceras de liquidacion para el empleado"
        
        StrSql = " SELECT p2.pliqnro FROM periodo p2 "
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                 arrItempliqnro(I) = rsConsult!pliqnro
            End If
        rsConsult.Close
        
        '------------------------------------------------------------------
        'Busco datos de cabeceras de liquidacion
        '------------------------------------------------------------------
        StrSql = " SELECT c.cliqnro, p2.pliqnro FROM cabliq c"
        StrSql = StrSql & " INNER JOIN proceso p ON p.pronro = c.pronro"
        StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        StrSql = StrSql & " AND c.empleado= " & Ternro
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           liqReal = True
           arrItemReal(I) = -1
        End If
        
        'Me guardo el pliqnro
        If Not rsConsult.EOF Then
           arrItempliqnro(I) = rsConsult!pliqnro
        End If
        'armar lst de cliqnro
        
        lstcliqnro = "0"
        
        Do While Not rsConsult.EOF
            lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
            rsConsult.MoveNext
        Loop
        
        If lstcliqnro <> "0" Then
             If Itemtip(j) = "AC" Then
                
                StrSql = " SELECT sum(almonto) monto "
                StrSql = StrSql & " FROM acu_liq"
                StrSql = StrSql & " WHERE acunro = " & Itemval(j)
                StrSql = StrSql & " AND cliqnro in ( " & lstcliqnro & " )"
                '---LOG---
                Flog.writeline "Procesando acumulador: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                Flog.writeline StrSql
                OpenRecordset StrSql, rsConsult
                If Not rsConsult.EOF Then
                    arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                    itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                End If
                rsConsult.Close
            ElseIf Itemtip(j) = "CO" Then
                     StrSql = " SELECT sum(d.dlimonto) monto "
                     StrSql = StrSql & " FROM detliq d "
                     StrSql = StrSql & " INNER JOIN concepto c ON c.concnro = d.concnro AND  c.conccod = '" & Itemval(j) & "'"
                     StrSql = StrSql & " WHERE cliqnro in ( " & lstcliqnro & " )"
                     '---LOG---
                     Flog.writeline "Procesando Concepto: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                     Flog.writeline StrSql
                     OpenRecordset StrSql, rsConsult
                     If Not rsConsult.EOF Then
                        arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                        itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                     End If
                     rsConsult.Close
                Else
                    Flog.writeline "El tipo de columna debe ser AC o CO no puede ser: " & Itemtip(j)
            End If
           
        End If
    Next I 'Proximo periodo
    
    
    'grabo det
    
    Flog.writeline "Grabando detalle para la columna: " & j
    StrSql = " INSERT INTO rep_simAnuAcuCon_det "
    StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
    StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
    StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
    StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
    StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
    StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
    StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal, historial "
    StrSql = StrSql & " ) VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & ",'" & Itemtip(j) & "'"
    StrSql = StrSql & "," & Itemval(j)
    StrSql = StrSql & ",'" & ItemDesc(j) & "'"
    StrSql = StrSql & "," & j
    For X = 0 To 11
        StrSql = StrSql & "," & arrItemReal(X)
        StrSql = StrSql & "," & arrItempliqnro(X)
        StrSql = StrSql & "," & numberForSQL(arrItemMonto(X))
    Next X
    StrSql = StrSql & "," & itemtotal
    StrSql = StrSql & "," & historial
    StrSql = StrSql & ")"
    If Itemtip(j) = "CO" Or Itemtip(j) = "AC" Then
        Flog.writeline StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
          
Next j 'Proximo ac o co



Siguiente:

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub


Sub generarDatosEmpleadoSim(Ternro, descripcion, orden, pliqdesde, pliqhasta, historial)
'------------------------------------------
'Busco los datos de SIMULACIONES
'------------------------------------------

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
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
Dim I As Integer
Dim liqReal As Boolean
Dim j As Integer
Dim lstcliqnro As String
Dim sql As String
Dim arrItemMonto(12) As Double
Dim arrItemReal(12) As Integer
Dim arrItempliqnro(12) As Long
Dim itemtotal As Double
Dim X As Integer

On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
tenomb1 = ""
tenomb2 = ""
tenomb3 = ""

'------------------------------------------------------------------
'Busco los datos del empleado SIMULADO
'------------------------------------------------------------------
StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
StrSql = StrSql & " FROM sim_empleado "
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = sim_empleado.ternro "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
Flog.writeline "Buscando datos del SIM empleado."
Flog.writeline StrSql
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
   Flog.writeline "Datos obtenidos Legajo: " & Legajo
Else
   Flog.writeline "Error al obtener los datos del empleado SIMULADO"
'   GoTo MError
End If

rsConsult.Close

'MR comento la búsqueda del período y la mando al main

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro = " & pliqdesde

Flog.writeline "Buscando datos del periodo desde"
       Flog.writeline StrSql
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
       Flog.writeline StrSql
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
  
   pliqhastames = rsConsult!pliqmes
   pliqhastaanio = rsConsult!pliqanio
   pliqhastafecha = rsConsult!pliqdesde
   
Else
   Flog.writeline "Error al obtener los datos del  periodo hasta"
 '  GoTo MError
End If

rsConsult.Close

Flog.writeline "Guardo Cabecera para el empleado SIMULADO"

'------------------------------------------------------------------
'Guardo la cabecera
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_simAnuAcuCon "
StrSql = StrSql & " (bpronro , ternro, apellido,apellido2, nombre, nombre2,"
StrSql = StrSql & " legajo , pliqdesde, pliqhasta, pronro,"
StrSql = StrSql & " orden , descabr, historico) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqdesde
StrSql = StrSql & "," & pliqhasta
StrSql = StrSql & ",'" & 0 & "'"
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Left(descripcion, 200) & "'"
StrSql = StrSql & ",3"
'For ii = 1 To 8
'    StrSql = StrSql & "," & numberForSQL(acuconmonto(ii))
'    StrSql = StrSql & ",'" & acuconetiq(ii) & "'"
'Next
StrSql = StrSql & ")"
If Legajo > 0 Then
   Flog.writeline StrSql
   objConn.Execute StrSql, , adExecuteNoRecords
End If

'Cantidad de periodos
DifPeriodos = DateDiff("m", pliqdesdefecha, pliqhastafecha)


'Ciclo por columna(ac-co)
For j = 1 To Indice - 1
    
    'Inicializo arrays
    For I = 0 To 12
        arrItemMonto(I) = 0
        arrItemReal(I) = 0
        arrItempliqnro(I) = 0
    Next I
    
    itemtotal = 0
    
    'Ciclo por periodo
    For I = 0 To DifPeriodos
        Flog.writeline "Buscando datos de cabeceras de liquidacion para el empleado"
        
        '------------------------------------------------------------------
        'Busco datos de cabeceras de liquidacion SIMULACION
        '------------------------------------------------------------------
        arrItemReal(I) = 0
        liqReal = False
        'rsConsult.Close
        
        StrSql = " SELECT p2.pliqnro FROM periodo p2 "
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                 arrItempliqnro(I) = rsConsult!pliqnro
            End If
        rsConsult.Close
        StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_cabliq c"
        StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
        StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        StrSql = StrSql & " AND c.empleado= " & Ternro
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
        
        'Me guardo el pliqnro
        If Not rsConsult.EOF Then
           arrItempliqnro(I) = rsConsult!pliqnro
        End If
        'armar lst de cliqnro
        
        lstcliqnro = "0"
        
        Do While Not rsConsult.EOF
            lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
            rsConsult.MoveNext
        Loop
        
        If lstcliqnro <> "0" Then
             If Itemtip(j) = "AC" Then
                
                StrSql = " SELECT sum(almonto) monto "
                StrSql = StrSql & " FROM sim_acu_liq"
                StrSql = StrSql & " WHERE acunro = " & Itemval(j)
                StrSql = StrSql & " AND cliqnro in ( " & lstcliqnro & " )"
                '---LOG---
                Flog.writeline "Procesando acumulador: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                Flog.writeline StrSql
                OpenRecordset StrSql, rsConsult
                If Not rsConsult.EOF Then
                    arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                    itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                End If
                rsConsult.Close
            ElseIf Itemtip(j) = "CO" Then
                     StrSql = " SELECT sum(d.dlimonto) monto "
                     StrSql = StrSql & " FROM sim_detliq d"
                     StrSql = StrSql & " INNER JOIN concepto c ON c.concnro = d.concnro AND  c.conccod = '" & Itemval(j) & "'"
                     StrSql = StrSql & " WHERE cliqnro in ( " & lstcliqnro & " )"
                     '---LOG---
                     Flog.writeline "Procesando Concepto: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                     Flog.writeline StrSql
                     OpenRecordset StrSql, rsConsult
                     If Not rsConsult.EOF Then
                        arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                        itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                     End If
                     rsConsult.Close
                Else
                    Flog.writeline "El tipo de columna debe ser AC o CO no puede ser: " & Itemtip(j)
            End If
        End If
    Next I 'Proximo periodo
    
    'grabo det
    
    Flog.writeline "Grabando detalle para la columna: " & j
    StrSql = " INSERT INTO rep_simAnuAcuCon_det "
    StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
    StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
    StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
    StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
    StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
    StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
    StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal, historial "
    StrSql = StrSql & " ) VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & ",'" & Itemtip(j) & "'"
    StrSql = StrSql & "," & Itemval(j)
    StrSql = StrSql & ",'" & ItemDesc(j) & "'"
    StrSql = StrSql & "," & j
    For X = 0 To 11
        StrSql = StrSql & "," & arrItemReal(X)
        StrSql = StrSql & "," & arrItempliqnro(X)
        StrSql = StrSql & "," & numberForSQL(arrItemMonto(X))
    Next X
    StrSql = StrSql & "," & itemtotal
    StrSql = StrSql & "," & historial
    StrSql = StrSql & ")"
    If Itemtip(j) = "CO" Or Itemtip(j) = "AC" Then
        Flog.writeline StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
          
Next j 'Proximo ac o co



Siguiente:

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

Sub generarDatosEmpleadoHis(Ternro, descripcion, orden, pliqdesde2, pliqhasta2, historial)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
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
Dim I As Integer
Dim liqReal As Boolean
Dim j As Integer
Dim lstcliqnro As String
Dim sql As String
Dim arrItemMonto(12) As Double
Dim arrItemReal(12) As Integer
Dim arrItempliqnro(12) As Long
Dim itemtotal As Double
Dim X As Integer


On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
tenomb1 = ""
tenomb2 = ""
tenomb3 = ""

'-------------------------------------------------------------------------------
'busco los datos del empleado del historico de simulacion sebastian stremel 01/11/2011
'-------------------------------------------------------------------------------
StrSql = " SELECT * FROM sim_his_empleado "
StrSql = StrSql & " INNER JOIN tercero on tercero.ternro = sim_his_empleado.ternro "
StrSql = StrSql & " WHERE sim_his_empleado.simhisnro = " & historico2 & " AND  sim_his_empleado.Ternro =" & Ternro
Flog.writeline "Buscando datos del SIM_HIS empleado."
Flog.writeline StrSql
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
   Flog.writeline "Datos obtenidos Legajo: " & Legajo & " nro de tercero:" & Ternro
Else
   Flog.writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'-------------------------------------------------------------------------------

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro = " & pliqdesde2

Flog.writeline "Buscando datos del periodo desde"
      Flog.writeline StrSql
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
StrSql = StrSql & " WHERE pliqnro = " & pliqhasta2

Flog.writeline "Buscando datos del periodo hasta"
Flog.writeline StrSql
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then

   pliqhastames = rsConsult!pliqmes
   pliqhastaanio = rsConsult!pliqanio
   pliqhastafecha = rsConsult!pliqhasta

Else
   Flog.writeline "Error al obtener los datos del  periodo hasta"
  ' GoTo MError
End If

rsConsult.Close

Flog.writeline "Guardo Cabecera para el empleado"

'------------------------------------------------------------------
'Guardo la cabecera
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_simAnuAcuCon "
StrSql = StrSql & " (bpronro , ternro, apellido,apellido2, nombre, nombre2,"
StrSql = StrSql & " legajo , pliqdesde, pliqhasta, pronro,"
StrSql = StrSql & " orden , descabr,historico) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqdesde2
StrSql = StrSql & "," & pliqhasta2
StrSql = StrSql & ",'" & 0 & "'"
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Left(descripcion, 200) & "'," & -1
'For ii = 1 To 8
'    StrSql = StrSql & "," & numberForSQL(acuconmonto(ii))
'    StrSql = StrSql & ",'" & acuconetiq(ii) & "'"
'Next
StrSql = StrSql & ")"
If Legajo > 0 Then
   Flog.writeline StrSql
   objConn.Execute StrSql, , adExecuteNoRecords
End If


'Cantidad de periodos
DifPeriodos = DateDiff("m", pliqdesdefecha, pliqhastafecha)


'Ciclo por columna(ac-co)
For j = 1 To Indice - 1
    
    'Inicializo arrays
    For I = 0 To 12
        arrItemMonto(I) = 0
        arrItemReal(I) = -1
        arrItempliqnro(I) = 0
    Next I
    
    itemtotal = 0
    
    'Ciclo por periodo
    For I = 0 To DifPeriodos
        Flog.writeline "Buscando datos de cabeceras de liquidacion para el empleado"
        
        StrSql = " SELECT p2.pliqnro FROM periodo p2 "
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                 arrItempliqnro(I) = rsConsult!pliqnro
            End If
        rsConsult.Close
        
        '------------------------------------------------------------------
        'Busco datos de cabeceras de liquidacion
        '------------------------------------------------------------------
        StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_his_cabliq c"
        StrSql = StrSql & " INNER JOIN sim_his_proceso p ON p.pronro = c.pronro AND p.simhisnro = " & historico2
        StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
        StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
        StrSql = StrSql & " AND c.empleado= " & Ternro & " AND c.simhisnro = " & historico2
        Flog.writeline StrSql
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           arrItemReal(I) = 0
           liqReal = False
           rsConsult.Close
           
           StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_his_cabliq c"
           StrSql = StrSql & " INNER JOIN sim_his_proceso p ON p.pronro = c.pronro AND p.simhisnro = " & historico2
           StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
           StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
           StrSql = StrSql & " AND c.empleado= " & Ternro & " AND c.simhisnro = " & historico2
           Flog.writeline StrSql
           OpenRecordset StrSql, rsConsult
        End If
        
        'Me guardo el pliqnro
        If Not rsConsult.EOF Then
           arrItempliqnro(I) = rsConsult!pliqnro
        End If
        'armar lst de cliqnro
        
        lstcliqnro = "0"
        
        Do While Not rsConsult.EOF
            lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
            rsConsult.MoveNext
        Loop
        
        If lstcliqnro <> "0" Then
             If Itemtip(j) = "AC" Then
                
                StrSql = " SELECT sum(almonto) monto "
                StrSql = StrSql & " FROM sim_his_acu_liq "
                StrSql = StrSql & " WHERE acunro = " & Itemval(j)
                StrSql = StrSql & " AND cliqnro in ( " & lstcliqnro & " )  AND simhisnro = " & historico2
                '---LOG---
                Flog.writeline "Procesando acumulador: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                Flog.writeline StrSql
                OpenRecordset StrSql, rsConsult
                If Not rsConsult.EOF Then
                    arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                    itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                End If
                rsConsult.Close
                
                'If liqReal And arrItemMonto(I) = 0 Then
                    'Si tiene cabeceras en real pero igual no encontro ninguna liquidada
                    'o valores en la misma fuerzo la busqueda en sim
                '    arrItemReal(I) = 0
                    
                '    StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_cabliq c"
                '    StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
                 '   StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
                  '  StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
                  '  StrSql = StrSql & " AND c.empleado= " & Ternro
                  '  OpenRecordset StrSql, rsConsult
                    
                   ' lstcliqnro = "0"
        
                    'Do While Not rsConsult.EOF
                     '   lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
                      '  rsConsult.MoveNext
                   ' Loop
        
                    'If lstcliqnro <> "0" Then
                    
                     '   StrSql = " SELECT sum(almonto) monto "
                      '  StrSql = StrSql & " FROM sim_acu_liq"
                       ' StrSql = StrSql & " WHERE acunro = " & Itemval(j)
                        'StrSql = StrSql & " AND cliqnro in ( " & lstcliqnro & " )"
                        'OpenRecordset StrSql, rsConsult
                       ' If Not rsConsult.EOF Then
                      '      arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                      '      itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                      '  End If
                      '  rsConsult.Close
                   ' End If
               ' End If
            ElseIf Itemtip(j) = "CO" Then
                     StrSql = " SELECT sum(d.dlimonto) monto "
                     StrSql = StrSql & " FROM sim_his_detliq d"
                     StrSql = StrSql & " INNER JOIN concepto c ON c.concnro = d.concnro AND  c.conccod = " & Itemval(j)
                     StrSql = StrSql & " WHERE cliqnro in ( " & lstcliqnro & " ) AND d.simhisnro = " & historico2
                     '---LOG---
                     Flog.writeline "Procesando Concepto: " & Itemval(j) & "para el periodo: " & arrItempliqnro(I) & "(" & DateAdd("m", I, pliqdesdefecha) & ")"
                     Flog.writeline StrSql
                     OpenRecordset StrSql, rsConsult
                     If Not rsConsult.EOF Then
                        arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                        itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                     End If
                     rsConsult.Close
                     'If liqReal And arrItemMonto(I) = 0 Then
                        'Si tiene cabeceras en real pero igual no encontro ninguna liquidada
                        'o valores en la misma fuerzo la busqueda en sim
                     '   arrItemReal(I) = 0
                    
                     '   StrSql = " SELECT c.cliqnro, p2.pliqnro FROM sim_cabliq c"
                     '   StrSql = StrSql & " INNER JOIN sim_proceso p ON p.pronro = c.pronro"
                     '   StrSql = StrSql & " INNER JOIN periodo p2 ON p2.pliqnro = p.pliqnro"
                     '   StrSql = StrSql & " WHERE p2.pliqdesde = " & ConvFecha(DateAdd("m", I, pliqdesdefecha))
                     '   StrSql = StrSql & " AND c.empleado= " & Ternro
                     '   OpenRecordset StrSql, rsConsult
                        
                     '   lstcliqnro = "0"
            '
              '          Do While Not rsConsult.EOF
             '               lstcliqnro = lstcliqnro & "," & rsConsult!cliqnro
               '             rsConsult.MoveNext
                '        Loop
        
                 '       If lstcliqnro <> "0" Then
                        
                  '          StrSql = " SELECT sum(d.dlimonto) monto "
                    '        StrSql = StrSql & " FROM sim_detliq d"
                   '         StrSql = StrSql & " INNER JOIN concepto c ON c.concnro = d.concnro AND  c.conccod = " & Itemval(j)
                    '        StrSql = StrSql & " WHERE cliqnro in ( " & lstcliqnro & " )"
                     '       OpenRecordset StrSql, rsConsult
                     '       If Not rsConsult.EOF Then
                     '           arrItemMonto(I) = IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                     '           itemtotal = itemtotal + IIf(IsNull(rsConsult!Monto), 0, rsConsult!Monto)
                      '      End If
                      '      rsConsult.Close
                      '  End If
                   ' End If ' If liqReal And arrItemMonto(i) = 0 Then
                Else
                    Flog.writeline "El tipo de columna debe ser AC o CO no puede ser: " & Itemtip(j)
            End If
           
        End If
    Next I 'Proximo periodo
    
    
    'grabo det
    
    Flog.writeline "Grabando detalle para la columna: " & j
    StrSql = " INSERT INTO rep_simAnuAcuCon_det "
    StrSql = StrSql & " (bpronro,ternro,itemTipo,itemnro,itemdesc,itemorden,"
    StrSql = StrSql & " itemreal1,pliqnro1,itemmonto1,itemreal2,pliqnro2,itemmonto2,"
    StrSql = StrSql & " itemreal3,pliqnro3,itemmonto3,itemreal4,pliqnro4,itemmonto4, "
    StrSql = StrSql & " itemreal5,pliqnro5,itemmonto5,itemreal6,pliqnro6,itemmonto6, "
    StrSql = StrSql & " itemreal7,pliqnro7,itemmonto7,itemreal8,pliqnro8,itemmonto8, "
    StrSql = StrSql & " itemreal9,pliqnro9,itemmonto9,itemreal10,pliqnro10,itemmonto10, "
    StrSql = StrSql & " itemreal11,pliqnro11,itemmonto11,itemreal12,pliqnro12,itemmonto12,itemtotal,historial "
    StrSql = StrSql & " ) VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & ",'" & Itemtip(j) & "'"
    StrSql = StrSql & "," & Itemval(j)
    StrSql = StrSql & ",'" & ItemDesc(j) & "'"
    StrSql = StrSql & "," & j
    For X = 0 To 11
        StrSql = StrSql & "," & arrItemReal(X)
        StrSql = StrSql & "," & arrItempliqnro(X)
        StrSql = StrSql & "," & numberForSQL(arrItemMonto(X))
    Next X
    StrSql = StrSql & "," & itemtotal
    StrSql = StrSql & "," & historial
    StrSql = StrSql & ")"
    If Itemtip(j) = "CO" Or Itemtip(j) = "AC" Then
             Flog.writeline StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
    End If
          
Next j 'Proximo ac o co



Siguiente:

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

