Attribute VB_Name = "mdlLeerRegistraciones"
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



Private Sub InsertaRegistracion(strReg As String)
Dim nroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim hora As String
Dim EntradaSalida As String
Dim nroReloj As Long
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim TipoTarj As Integer

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, " ")
    nroLegajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strReg, " ")
    Fecha = Mid(strReg, pos1, pos2 - pos1)
    RegFecha = Fecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strReg, " ")
    hora = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Not objFechasHoras.ValidarHora(hora) Then
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strReg, " ")
    nroReloj = Mid(strReg, pos1, pos2 - pos1)
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroReloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        TipoTarj = objRs!tptrnro
    End If
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strReg, " ")
    EntradaSalida = IIf(Trim(Mid(strReg, pos1)) = "20", "E", "S")
       
    ' 15/07/2003
    ' no poner comillas al nroLegajo porque no toma bien los legajos que comienzan con 000...
    'If codReloj = 13 Then
    '    StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = " & nroLegajo & " AND tptrnro = 2 AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    'Else
    '    StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = " & nroLegajo & " AND tptrnro = 1 AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    'End If

    ' ----------------------------------------------------
    'FZG 06/08/2003
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = " & nroLegajo & " AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = " & nroLegajo & " AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
         InsertaError 1, 33
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
    Else
        InsertaError 1, 92
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

    If App.PrevInstance Then End
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas


'    'Crea el archivo de log
'    ' ----------
'    Nombre_Arch = PathFLog & "LecturaReg " & Format(Date, "dd-mm-yyyy") & ".log"
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set Flog = fs.CreateTextFile(Nombre_Arch, True)

    OpenConnection strconexion, objConn
    strCmdLine = Command()
    'strCmdLine = "20396"
    If IsNumeric(strCmdLine) Then
        NroProceso = strCmdLine
    Else
        Exit Sub
    End If
    
    
    'Crea el archivo de log
    ' ----------
    Nombre_Arch = PathFLog & "LecturaReg " & "-" & NroProceso & "-" & Format(Date, "dd-mm-yyyy") & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Inicio Transferencia " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' 15/07/2003
    Call ComenzarTransferencia
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Flog.writeline "Lectura completa. Fin " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    Flog.Close
    
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
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 210 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
     Else
        Exit Sub
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Path = Directorio
    
    Dim fc, f1, s2
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
        If UCase(Right(archivo.Name, 4)) = ".REG" Then
            NArchivo = archivo.Name
            LeeRegistraciones Directorio & "\" & archivo.Name
        End If
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Next
    
End Sub

Public Sub TerminarTransferencia_new()
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
End Sub

Public Sub TerminarTransferencia()

    If Not HuboError Then
        StrSql = "DELETE FROM batch_proceso WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    If objConn.State = adStateOpen Then objConn.Close
End Sub


Private Sub LeeRegistraciones(NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String


    If App.PrevInstance Then Exit Sub

    'MyBeginTrans
    
    'Espero hasta que se crea el archivo de registraciones
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    On Error GoTo 0
    
    
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO car_pin(modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      "210,'" & NombreArchivo & "',0,0," & ConvFecha(Date) & ",'Carga : " & Now & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "car_pin")
    
    End If
    Do While Not f.AtEndOfStream
        
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If Trim(strLinea) <> "" Then InsertaRegistracion strLinea
        
    Loop
    
    StrSql = "UPDATE car_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    
    Flog.writeline "archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Set f = fs.getfile(NombreArchivo)
    
    Archivo_Aux = Replace(Format(Now, "yyyy-mm-dd hh:mm:ss"), ":", "-") & " " & NArchivo
    ' ahora
    f.Move Path & "\bk\" & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "bk"
    
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
