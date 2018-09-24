Attribute VB_Name = "mdlLeerRelojDBF"
Option Explicit

Dim fs, f
Global Flog
Dim objFechasHoras As New FechasHoras
Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global ObjetoVentana As Object
Global HuboError As Boolean

Global objConndBase As New ADODB.Connection

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long




Private Sub InsertaRegistracion(ByVal NroLegajo As String, ByVal Fecha As Date, ByVal hora As String, ByVal EntradaSalida As String, ByVal NroReloj As Long)
Dim Ternro As Long
Dim codReloj As Integer
Dim TipoTarj As Integer

    RegLeidos = RegLeidos + 1
    
    'Fecha
    RegFecha = Fecha
    
    'Hora
    If Not objFechasHoras.ValidarHoraAMPM(hora) Then
        InsertaError 5, 38
        Exit Sub
    End If
    
    'Nro Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & NroReloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        InsertaError 0, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        TipoTarj = objRs!tptrnro
    End If
    ' ----------------------------------------------------
    
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = '" & CLng(NroLegajo) & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & CLng(NroLegajo) & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
         InsertaError 2, 33
         Exit Sub
      End If
    End If
    ' Primero se busca con el numero de tarjeta y el tipo asociado al reloj. Si no se encuentra,
    ' se busca solo con el número de la tarjeta. O.D.A. 06/08/2003
    ' ----------------------------------------------------
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & hora & "' AND ternro = " & Ternro & " AND regentsal = '" & EntradaSalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & hora & "','" & EntradaSalida & "'," & codReloj & ",'I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        ' Inserto el par (Ternro,Fecha)
        Call InsertarWF_Lecturas(Ternro, Fecha)
        
    Else
        InsertaError 0, 92
    End If
        
End Sub

Private Sub InsertaError(NroCampo As Byte, nroError As Long)

    Flog.writeline "antes de insertar error en car_err" & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    StrSql = "INSERT INTO Car_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "insertó en car_err" & Format(Now, "dd/mm/yyyy hh:mm:ss")
        
    RegError = RegError + 1
    
End Sub

Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim Rs_WF_Lec_Fechas As New ADODB.Recordset
Dim Rs_WF_Lec_Terceros As New ADODB.Recordset
Dim NroProcesoHC As Long
Dim NroProcesoAD As Long

Dim rs_ONLINE As New ADODB.Recordset
Dim Proc_ONLINE As Boolean 'Si el procesamiento On Line está o no activo
Dim HC_ONLINE As Boolean 'Si genera o no procesos de Horario Cumplido por procesamiento On Line
Dim AD_ONLINE As Boolean 'Si genera o no procesos de Acumulado Diario por procesamiento On Line
Dim PID As String

    If App.PrevInstance Then End
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas


    OpenConnection strconexion, objConn
    strCmdLine = Command()
    If IsNumeric(strCmdLine) Then
        NroProceso = strCmdLine
    Else
        Exit Sub
    End If
        
    'Crea el archivo de log
    ' ----------
    Nombre_Arch = PathFLog & "InterfaceRelojDBF " & "-" & NroProceso & "-" & Format(Date, "dd-mm-yyyy") & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        
    Call CargarNombresTablasTemporales
    
    ' Creo la tabla temporal
    Call CreateTempTable(TTempWFLecturas)
    
        
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Flog.writeline "Inicio Transferencia " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Call ComenzarTransferencia
    
    ' Revisa y setea si el procesamineto OnLine está Activo o No
    StrSql = "SELECT * FROM GTI_puntos_proc " & _
             " INNER JOIN GTI_proc_online ON GTI_puntos_proc.ptoprcnro = GTI_proc_online.ptoprcnro " & _
             " WHERE GTI_puntos_proc.ptoprcid = 19 AND GTI_puntos_proc.ptoprcact = -1 "
    OpenRecordset StrSql, rs_ONLINE
    
    Proc_ONLINE = False
    HC_ONLINE = False
    AD_ONLINE = False
    
    If rs_ONLINE.EOF Then
        Proc_ONLINE = False
    Else
        Proc_ONLINE = True
        Do While Not rs_ONLINE.EOF
        Select Case rs_ONLINE!btprcnro
        Case 1:
            HC_ONLINE = True
        Case 2:
            AD_ONLINE = True
        Case Else
        End Select
        
            rs_ONLINE.MoveNext
        Loop
    End If
    
    If Proc_ONLINE Then
        StrSql = "SELECT DISTINCT fecha FROM " & TTempWFLecturas
        OpenRecordset StrSql, Rs_WF_Lec_Fechas
    
        Do While Not Rs_WF_Lec_Fechas.EOF
            If HC_ONLINE Then
                ' Inserto en batch_proceso un HC
                StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, " & _
                         "bprcestado, empnro) " & _
                         "VALUES (" & 1 & "," & ConvFecha(Date) & ", 'super'" & ",'" & Format(Now, "hh:mm:ss ") & "' " & _
                         ", " & ConvFecha(Rs_WF_Lec_Fechas!Fecha) & ", " & ConvFecha(Rs_WF_Lec_Fechas!Fecha) & _
                         ", 'Pendiente', 0)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'recupero el numero de proceso generado
                NroProcesoHC = getLastIdentity(objConn, "Batch_Proceso")
            End If
            
            If AD_ONLINE Then
                ' Inserto en batch_proceso un AD
                StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, " & _
                         "bprcestado, empnro) " & _
                         "VALUES (" & 2 & "," & ConvFecha(Date) & ", 'super'" & ",'" & Format(Now, "hh:mm:ss ") & "' " & _
                         ", " & ConvFecha(Rs_WF_Lec_Fechas!Fecha) & ", " & ConvFecha(Rs_WF_Lec_Fechas!Fecha) & _
                         ", 'Pendiente', 0)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'recupero el numero de proceso generado
                NroProcesoAD = getLastIdentity(objConn, "Batch_Proceso")
            End If
            
            ' Inserto en batch_empleados los empleados de los procesos generados
            StrSql = "SELECT DISTINCT ternro FROM " & TTempWFLecturas & _
                     " WHERE fecha = " & ConvFecha(Rs_WF_Lec_Fechas!Fecha)
            OpenRecordset StrSql, Rs_WF_Lec_Terceros
            Do While Not Rs_WF_Lec_Terceros.EOF
                ' para HC
                If HC_ONLINE Then
                    StrSql = "INSERT INTO batch_empleado (bpronro, ternro, estado) VALUES (" & _
                             NroProcesoHC & "," & Rs_WF_Lec_Terceros!Ternro & ", NULL )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                If AD_ONLINE Then
                    ' para AD
                    StrSql = "INSERT INTO batch_empleado (bpronro, ternro, estado) VALUES (" & _
                             NroProcesoAD & "," & Rs_WF_Lec_Terceros!Ternro & ", NULL )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                Rs_WF_Lec_Terceros.MoveNext
            Loop
        
            Rs_WF_Lec_Fechas.MoveNext
        Loop
    End If
    
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Flog.writeline "Lectura completa. Fin " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    Flog.Close
    
    Call BorrarTempTable(TTempWFLecturas)
    
Terminar:
    ' eliminar el proceso de la tabla batch_proceso si es que termino correctamente
    Call TerminarTransferencia
    
End Sub

Public Sub ComenzarTransferencia()
Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim IncPorc As Single
Dim Progreso As Single

    'OpenConnection "DSN=Informix", objConn
    'OpenConnection "DSN=Rhpro", objConn
    'OpenConnection strconexion, objConn
    
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_dirsalidas)
    Else
        Exit Sub
    End If
    
    Flog.writeline "Directorio de Registraciones:  " & Directorio
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 213 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
     Else
        Exit Sub
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Path = Directorio
    
    Set Folder = fs.GetFolder(Directorio)
    Set CArchivos = Folder.Files
    
    'Determino la proporcion de progreso
    Progreso = 0
    If Not CArchivos.Count = 0 Then
        Flog.writeline CArchivos.Count & " Archivos de registraciones encontrados " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        IncPorc = 100 / CArchivos.Count
    End If
    
    HuboError = False
    For Each archivo In CArchivos
        If UCase(Right(archivo.Name, 4)) = ".DBF" Then
            NArchivo = archivo.Name
            Call LeeReloj(Directorio, NArchivo)
        End If
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Next
    
End Sub

Public Sub TerminarTransferencia()
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset

'    If Not HuboError Then
'        StrSql = "DELETE FROM batch_proceso WHERE bpronro = " & NroProceso
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
    
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If Not HuboError Then
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    End If
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
    
    If objConn.State = adStateOpen Then objConn.Close
    If objConndBase.State = adStateOpen Then objConndBase.Close
End Sub




Private Sub LeeReloj(ByVal Directorio As String, ByVal NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs As New ADODB.Recordset
Dim NroReloj As Long
Dim Fecha As Date

    OpenConnection "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directorio & ";Extended Properties='DBASE IV;';Persist Security Info=False", objConndBase
    
    If App.PrevInstance Then Exit Sub
    On Error GoTo CE

    StrSql = "SELECT * FROM asisten"
    OpenRecordsetdBase StrSql, rs

    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not rs.EOF Then
        StrSql = "INSERT INTO car_pin(modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      "213,'" & NombreArchivo & "',0,0," & ConvFecha(Date) & ",'Carga : " & Now & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "car_pin")
    End If
    
    Do While Not rs.EOF
        NroReloj = 1
        'Reviso la fecha porque los drivers estan haciendo macanas como
        ' si el formato es dd/mm/yy lo transforman en dd/mm/yyyy pero ...
        ' 19/02/04 lo transforma a 19/02/1904 en lugar de 19/02/2004
        If Not IsNull(rs!Fecha) Then
            If rs!Fecha < CDate("01/01/2000") Then
                Fecha = Format(rs!Fecha, "dd/mm/yy")
            Else
                Fecha = rs!Fecha
            End If
            NroLinea = NroLinea + 1
            Call InsertaRegistracion(rs!legajo, Fecha, IIf(IsNull(rs!entrada), "Nulo", rs!entrada), "E", NroReloj)
            NroLinea = NroLinea + 1
            Call InsertaRegistracion(rs!legajo, Fecha, IIf(IsNull(rs!salida), "Nulo", rs!salida), "S", NroReloj)
        Else
            'Fecha nula
            InsertaError 4, 41
        End If
        rs.MoveNext
    Loop
    
    StrSql = "UPDATE car_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'cierro el recordset
    If rs.State = adStateOpen Then rs.Close
    
    Flog.writeline "archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
   
    Set f = fs.getfile(Directorio & "\" & NombreArchivo)
    
    Archivo_Aux = Replace(Format(Now, "yyyy-mm-dd hh:mm:ss"), ":", "-") & " " & NArchivo
    'f.Move Path & "\bk\" & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "bk"
    f.Move Path & "\bk\" & Archivo_Aux
    
    Flog.writeline "archivo movido " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'antes
    'f.Move Mid(NombreArchivo, 1, Len(NombreArchivo) - 3) & "bk"
    'MyCommitTrans
    
    
    ' FGZ 24/07/2003 -----------------
fin:
    
    Exit Sub
    
CE:
    'MyRollbackTrans
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    HuboError = True
    
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
   'FGZ - 09/09/2003
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    
    GoTo fin
        
End Sub


Public Sub OpenRecordsetdBase(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConndBase, adOpenDynamic, lockType, adCmdText
End Sub


