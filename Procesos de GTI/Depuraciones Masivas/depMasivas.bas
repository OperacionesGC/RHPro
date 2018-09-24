Attribute VB_Name = "depMasivas"
Global Const Version = 1
Global Const FechaVersion = "14/08/2009"   'Encriptacion de string connection
Global Const UltimaModificacion = "Manuel Lopez"
Global Const UltimaModificacion1 = "Encriptacion de string connection"

'------------------------------------------------------------------------------------------------------------------
Option Explicit

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
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim PID As String
Dim ArrParametros

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
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
    Exit Sub
    End If
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "DepMasivas" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Depuraciones Masivas : " & Now
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
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       fechadesde = objRs!bprcfecdesde
       Flog.writeline "Fecha Desde: " & fechadesde
       fechahasta = objRs!bprcfechasta
       Flog.writeline "Fecha Hasta: " & fechahasta
       param = Split(objRs!bprcparam, ".")
       tipoDepuracion = param(0)
       Flog.writeline "Tipo de Depuracion: " & tipoDepuracion
       Flog.writeline "Historico:  " & param(1)
       
        If UCase(param(1)) = "VERDADERO" Or param(1) = "TRUE" Then
            historico = True
        Else
            historico = False
        End If
       
       'historico = CBool(param(1))
       Flog.writeline "Historico:  " & IIf(CBool(historico), "SI", "NO")
        
       'Me fijo que tipo de depuracion tengo que realizar
       If tipoDepuracion = 1 Then
          Flog.writeline "depurarAcumDiario(NroProceso, fechadesde, fechahasta, historico)"
         'Acumulado Diario
          Call depurarAcumDiario(NroProceso, fechadesde, fechahasta, historico)
       Else
         If tipoDepuracion = 2 Then
            'Horario Cumplido
            Flog.writeline "Call depurarHorCumplido(NroProceso, fechadesde, fechahasta, historico)"
            Call depurarHorCumplido(NroProceso, fechadesde, fechahasta, historico)
         Else
            If tipoDepuracion = 3 Then
               'Registraciones
               Flog.writeline "call depurarRegistraciones (NroProceso, fechadesde, fechahasta, historico)"
               Call depurarRegistraciones(NroProceso, fechadesde, fechahasta, historico)
            Else
               Exit Sub
            End If
         End If
       End If
    Else
        Exit Sub
    End If
   
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

'--------------------------------------------------------------------
' Se encarga de depurar el Acumulado Diario
'--------------------------------------------------------------------
Sub depurarAcumDiario(NroProc, desde, hasta, historico)

Dim StrSql As String
Dim rsEmpl As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset

'Obtengo los empleados sobre los que tengo que realizar la depuracion
CargarEmpleados NroProc, rsEmpl

'Si no hay datos cargados en la tabla de empleados significa que
'hay que aplicar la depuracion sobre todos los empleados
If rsEmpl.EOF Then
       'Se aplica para todos los empleados
       
       If (historico) Then
           StrSql = " INSERT INTO gti_his_achdiario "
           StrSql = StrSql & " SELECT  * "
           StrSql = StrSql & " FROM gti_achdiario "
           StrSql = StrSql & " WHERE achdfecha >=" & ConvFecha(desde) & " AND achdfecha <=" & ConvFecha(hasta)
             
           objConn.Execute StrSql, , adExecuteNoRecords
           
           StrSql = " INSERT INTO gti_hisad "
           StrSql = StrSql & " SELECT  * "
           StrSql = StrSql & " FROM gti_acumdiario "
           StrSql = StrSql & " WHERE adfecha >=" & ConvFecha(desde) & " AND adfecha <=" & ConvFecha(hasta)
             
           objConn.Execute StrSql, , adExecuteNoRecords
       End If
       
       StrSql = " SELECT * FROM gti_achdiario "
       StrSql = StrSql & " WHERE achdfecha>=" & ConvFecha(desde)
       StrSql = StrSql & " AND achdfecha<=" & ConvFecha(hasta)
       
       OpenRecordset StrSql, rsAD
       
       Do Until rsAD.EOF
    
         If (historico) Then
           ' Guarda los datos de "gti_achdiario_estr" en el historico
           StrSql = " INSERT INTO gti_his_achdiario_estr "
           StrSql = StrSql & " SELECT  * "
           StrSql = StrSql & " FROM gti_achdiario_estr "
           StrSql = StrSql & " WHERE achdnro = " & rsAD!achdnro
        
           objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        StrSql = " DELETE FROM gti_achdiario_estr "
        StrSql = StrSql & " WHERE achdnro = " & rsAD!achdnro
    
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rsAD.MoveNext
      
      Loop
    
      StrSql = " DELETE FROM gti_acumdiario "
      StrSql = StrSql & " WHERE adfecha >=" & ConvFecha(desde) & " AND adfecha <=" & ConvFecha(hasta)
      
      objConn.Execute StrSql, , adExecuteNoRecords
       
      StrSql = " DELETE FROM gti_achdiario "
      StrSql = StrSql & " WHERE achdfecha >=" & ConvFecha(desde) & " AND achdfecha <=" & ConvFecha(hasta)
      
      objConn.Execute StrSql, , adExecuteNoRecords
       
Else

   'Solo depuro los empleados de selectados
   Do While Not rsEmpl.EOF
       
       If (historico) Then
           StrSql = " INSERT INTO gti_his_achdiario "
           StrSql = StrSql & " SELECT  * "
           StrSql = StrSql & " FROM gti_achdiario "
           StrSql = StrSql & " WHERE achdfecha >=" & ConvFecha(desde) & " AND achdfecha <=" & ConvFecha(hasta)
           StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
             
           objConn.Execute StrSql, , adExecuteNoRecords
           
           StrSql = " INSERT INTO gti_hisad "
           StrSql = StrSql & " SELECT  * "
           StrSql = StrSql & " FROM gti_acumdiario "
           StrSql = StrSql & " WHERE adfecha >=" & ConvFecha(desde) & " AND adfecha <=" & ConvFecha(hasta)
           StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
             
           objConn.Execute StrSql, , adExecuteNoRecords
       End If
       
       StrSql = " SELECT * FROM gti_achdiario "
       StrSql = StrSql & " WHERE achdfecha>=" & ConvFecha(desde)
       StrSql = StrSql & " AND achdfecha<=" & ConvFecha(hasta)
       StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
       
       OpenRecordset StrSql, rsAD
       
       Do While Not rsAD.EOF
    
         If (historico) Then
           ' Guarda los datos de "gti_achdiario_estr" en el historico
           StrSql = " INSERT INTO gti_his_achdiario_estr "
           StrSql = StrSql & " SELECT  * "
           StrSql = StrSql & " FROM gti_achdiario_estr "
           StrSql = StrSql & " WHERE achdnro = " & rsAD!achdnro
        
           objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        StrSql = " DELETE FROM gti_achdiario_estr "
        StrSql = StrSql & " WHERE achdnro = " & rsAD!achdnro
    
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rsAD.MoveNext
      
      Loop
    
      StrSql = " DELETE FROM gti_acumdiario "
      StrSql = StrSql & " WHERE adfecha >=" & ConvFecha(desde) & " AND adfecha <=" & ConvFecha(hasta)
      StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
      
      objConn.Execute StrSql, , adExecuteNoRecords
       
      StrSql = " DELETE FROM gti_achdiario "
      StrSql = StrSql & " WHERE achdfecha >=" & ConvFecha(desde) & " AND achdfecha <=" & ConvFecha(hasta)
      StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
      
      objConn.Execute StrSql, , adExecuteNoRecords
   
      rsEmpl.MoveNext
   Loop
End If

'Borro los empleados de la tabla
BorrarEmpleados NroProc

End Sub

'--------------------------------------------------------------------
' Se encarga de depurar el Horario Cumplido
'--------------------------------------------------------------------
Sub depurarHorCumplido(NroProc, desde, hasta, historico)

Dim StrSql As String
Dim rsEmpl As New ADODB.Recordset

'Obtengo los empleados sobre los que tengo que realizar la depuracion
CargarEmpleados NroProc, rsEmpl

'Si no hay datos cargados en la tabla de empleados significa que
'hay que aplicar la depuracion sobre todos los empleados
If rsEmpl.EOF Then
   'Se aplica para todos los empleados
   
   If (historico) Then
       StrSql = " INSERT INTO gti_hishc "
       StrSql = StrSql & " SELECT  * "
       StrSql = StrSql & " FROM gti_horcumplido "
       StrSql = StrSql & " WHERE horfecrep >=" & ConvFecha(desde) & " AND horfecrep <=" & ConvFecha(hasta)
         
       objConn.Execute StrSql, , adExecuteNoRecords
   End If
   
   StrSql = "DELETE FROM gti_horcumplido "
   StrSql = StrSql & " WHERE horfecrep >=" & ConvFecha(desde) & " AND horfecrep <=" & ConvFecha(hasta)

   objConn.Execute StrSql, , adExecuteNoRecords
   
Else
   'Solo depuro los empleados de selectados
   Do While Not rsEmpl.EOF
      If (historico) Then
         StrSql = " INSERT INTO gti_hishc "
         StrSql = StrSql & " SELECT  * "
         StrSql = StrSql & " FROM gti_horcumplido "
         StrSql = StrSql & " WHERE horfecrep >=" & ConvFecha(desde) & " AND horfecrep <=" & ConvFecha(hasta)
         StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
              
         objConn.Execute StrSql, , adExecuteNoRecords
      End If
      
      StrSql = "DELETE FROM gti_horcumplido "
      StrSql = StrSql & " WHERE horfecrep >=" & ConvFecha(desde) & " AND horfecrep <=" & ConvFecha(hasta)
      StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
      
      objConn.Execute StrSql, , adExecuteNoRecords
      
      rsEmpl.MoveNext
   Loop
     
End If

'Borro los empleados de la tabla
BorrarEmpleados NroProc

End Sub

'----------------------------------------------------------
' Se encarga de depurar las registraciones
'----------------------------------------------------------
Sub depurarRegistraciones(NroProc, desde, hasta, historico)

Dim StrSql
Dim rsEmpl As New ADODB.Recordset

'Obtengo los empleados sobre los que tengo que realizar la depuracion
CargarEmpleados NroProc, rsEmpl

'Si no hay datos cargados en la tabla de empleados significa que
'hay que aplicar la depuracion sobre todos los empleados
If rsEmpl.EOF Then
   'Se aplica para todos los empleados
   
   If (historico) Then
       StrSql = " INSERT INTO gti_hisreg "
       StrSql = StrSql & " SELECT  * "
       StrSql = StrSql & " FROM gti_registracion "
       StrSql = StrSql & " WHERE regfecha >=" & ConvFecha(desde) & " AND regfecha <=" & ConvFecha(hasta)
         
       objConn.Execute StrSql, , adExecuteNoRecords
   End If
   
   StrSql = "DELETE FROM gti_registracion "
   StrSql = StrSql & " WHERE regfecha >= " & ConvFecha(desde) & " AND regfecha <= " & ConvFecha(hasta)

   objConn.Execute StrSql, , adExecuteNoRecords
   
Else
   'Solo depuro los empleados de selectados
   Do While Not rsEmpl.EOF
      If (historico) Then
         StrSql = " INSERT INTO gti_hisreg "
         StrSql = StrSql & " SELECT  * "
         StrSql = StrSql & " FROM gti_registracion "
         StrSql = StrSql & " WHERE regfecha >=" & ConvFecha(desde) & " AND regfecha <=" & ConvFecha(hasta)
         StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
              
         objConn.Execute StrSql, , adExecuteNoRecords
      End If
      
      StrSql = "DELETE FROM gti_registracion "
      StrSql = StrSql & " WHERE regfecha >= " & ConvFecha(desde) & " AND regfecha <= " & ConvFecha(hasta)
      StrSql = StrSql & "   AND ternro = " & rsEmpl!Ternro
      
      objConn.Execute StrSql, , adExecuteNoRecords
      
      rsEmpl.MoveNext
   Loop
     
End If

'Borro los empleados de la tabla
BorrarEmpleados NroProc

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

'--------------------------------------------------------------------
' Se encarga de borrar todos los empleados del proceso
'--------------------------------------------------------------------
Sub BorrarEmpleados(NroProc)

Dim StrEmpl As String

    StrEmpl = " DELETE FROM batch_empleado "
    StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
    
    objConn.Execute StrEmpl, , adExecuteNoRecords
End Sub
