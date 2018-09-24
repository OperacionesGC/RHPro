Attribute VB_Name = "mdlInterface"
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
Global Rta
Global ObjetoVentana As Object
Global HuboError As Boolean
Global NombreArchivo As String
Global Separador As String
Global UsaEncabezado As Boolean

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Sub InsertaLinea(ByVal strLinea As String)

Dim CodLocal As Long
Dim Fecha As Date
Dim strfecha As String
Dim hora As String
Dim CodOperacion As Integer
Dim Estructura As Long

Dim rs_Estructura As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset

Dim pos1 As Byte
Dim pos2 As Byte


    RegLeidos = RegLeidos + 1
    
    'Codigo del Local (Sucursal)
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid(strLinea, pos1, pos2 - pos1)) Then
        CodLocal = CLng(Mid(strLinea, pos1, pos2 - pos1))
    Else
        InsertaError 1, 3
        Exit Sub
    End If
    
    'Fecha (dd/mm/aaaa)
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    strfecha = Mid(strLinea, pos1, pos2 - pos1)
    'Validar la fecha
    If Format(strfecha, "dd/mm/yyyy") = strfecha Then
        Fecha = CDate(strfecha)
    Else
        InsertaError 2, 4
        Exit Sub
    End If

    'Horario de Activacion / Desactivacion (hh:mm:ss)
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    hora = Mid(strLinea, pos1, pos2 - pos1)
    'Valido la hora
    If Not objFechasHoras.ValidarHoraLarga(hora) Then
        InsertaError 3, 7
        Exit Sub
    End If

    'Codigo de Operacion (Activacion = 1 y Desactivacion = 0)
    pos1 = pos2 + 1
    pos2 = Len(strLinea)
    CodOperacion = Mid(strLinea, pos1, pos2)
    If Not (CodOperacion = 0 Or CodOperacion = 1) Then
        InsertaError 4, 92
        Exit Sub
    End If
' ====================================================================
'   Validar los parametros Levantados
                
'Valido el codigo de la sucursal
StrSql = " SELECT * FROM estructura " & _
         " WHERE tenro = 1 AND estrcodext =" & CodLocal
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    StrSql = "SELECT * FROM sucursal " & _
         " WHERE estrnro =" & rs_Estructura!estrnro
    OpenRecordset StrSql, rs_Sucursal
        
    If Not rs_Sucursal.EOF Then
        Estructura = rs_Estructura!estrnro
    Else
       InsertaError 1, 56
    End If
End If
    
' ====================================================================
' Inserto en Mega_alarmas
StrSql = " INSERT INTO mega_alarmas (estrnro,alarfecha,alarhora,alaractivada) VALUES (" & _
        Estructura & "," & ConvFecha(Fecha) & ",'" & hora & "'," & CodOperacion & ")"
objConn.Execute StrSql, , adExecuteNoRecords
        
End Sub

Private Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
End Sub

Private Sub Main()
Dim strCmdLine As String

Dim rs_Batch_Proceso As New ADODB.Recordset
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
        
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 31 AND bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
    
    If Not rs_Batch_Proceso.EOF Then
        Call LevantarParamteros(rs_Batch_Proceso!bprcparam)
        Call ComenzarTransferencia
    End If
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
   
Terminar:
    ' eliminar el proceso de la tabla batch_proceso si es que termino correctamente
    'Call TerminarTransferencia
    
End Sub

Public Sub ComenzarTransferencia()
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim IncPorc As Single
Dim Progreso As Single

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_direntradas)
    Else
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 212 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        Separador = objRs!modseparador
        UsaEncabezado = CBool(objRs!modencab)
     Else
        Exit Sub
    End If
    
    Progreso = 0
    HuboError = False
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.fileexists(Directorio & "\" & NombreArchivo) Then
        Call LeeInterface(Directorio & "\" & NombreArchivo)
    Else
        ' El archivo no existe
        Exit Sub
    End If
    
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
End Sub



Private Sub LeeInterface(ByVal archivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String

    If App.PrevInstance Then Exit Sub

    'Espero hasta que se crea el archivo de registraciones
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(archivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    On Error GoTo 0
    
    
    On Error GoTo CE
    Set f = fs.OpenTextFile(archivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProceso & ",212,'" & NombreArchivo & "',0,0," & ConvFecha(Date) & ",'Mega_Alarma : " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
        
    End If
    Do While Not f.AtEndOfStream
        
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If Trim(strLinea) <> "" Then
            Call InsertaLinea(strLinea)
        End If
        
    Loop
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    
    'Borrar el archivo
    fs.Deletefile archivo, True

fin:
    
    Exit Sub
    
CE:
    HuboError = True
    GoTo fin
End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer


If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Pisa o no las novedades
        pos1 = 1
        pos2 = Len(parametros)
        NombreArchivo = Mid(parametros, pos1, pos2)
    End If
End If

End Sub

