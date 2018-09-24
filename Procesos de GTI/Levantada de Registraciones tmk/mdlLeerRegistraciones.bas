Attribute VB_Name = "mdlLeerRegistraciones"
Option Explicit

Dim f
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
'    strCmdLine = Command()
'    'strCmdLine = "20396"
'    If IsNumeric(strCmdLine) Then
'        NroProceso = strCmdLine
'    Else
'        Exit Sub
'    End If
        
    'Crea el archivo de log
    Nombre_Arch = PathFLog & "LecturaReg " & "-" & Format(Date, "dd-mm-yyyy") & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Transferencia " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Call LeeRegistraciones
    Flog.writeline "Lectura completa. Fin " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Close
    
End Sub

Private Sub LeeRegistraciones()
'Dim I As Integer
'Dim j As Integer
Dim Nroter As Long
Dim codReloj As Integer

Dim rs_TMK As New ADODB.Recordset
Dim rs_GTI_Registracion As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

codReloj = 1
StrSql = "SELECT * FROM TMK_BAJADA_REG "
OpenRecordset StrSql, rs_TMK

Do While Not rs_TMK.EOF
    'j = j + 1
    Nroter = 0

    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " WHERE empleg =" & rs_TMK!Legajo
    OpenRecordset StrSql, rs_Empleado

    If Not rs_Empleado.EOF Then
        Nroter = rs_Empleado!Ternro
        
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Nroter
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_TMK!RegFecha)
        StrSql = StrSql & " AND reghora ='" & rs_TMK!reghora & "'"
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado,regmanual) VALUES ("
            StrSql = StrSql & Nroter & ","
            StrSql = StrSql & ConvFecha(rs_TMK!RegFecha) & ",'"
            StrSql = StrSql & Replace(rs_TMK!reghora, ":", "") & "','"
            StrSql = StrSql & " ',"
            StrSql = StrSql & codReloj
            StrSql = StrSql & ",'I',"
            StrSql = StrSql & CInt(False) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_TMK!Legajo & " Fecha: " & rs_TMK!RegFecha & " Hora: " & rs_TMK!reghora
        End If
        
        'Borro
        StrSql = "DELETE TMK_BAJADA_REG "
        StrSql = StrSql & " WHERE regfecha =" & rs_TMK!RegFecha
        StrSql = StrSql & " AND reghora ='" & rs_TMK!reghora & "'"
        StrSql = StrSql & " AND legajo =" & rs_TMK!Legajo
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Flog.writeline " No se encontro el legajo " & rs_TMK!Legajo
    End If
    'I = I + 1
    
    rs_TMK.MoveNext
Loop

Fin:
    
    'Cierro todo
    If rs_GTI_Registracion.State = adStateOpen Then rs_GTI_Registracion.Close
    If rs_TMK.State = adStateOpen Then rs_TMK.Close
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    Set rs_GTI_Registracion = Nothing
    Set rs_TMK = Nothing
    Set rs_Empleado = Nothing
    
    Exit Sub
    
CE:
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin:
End Sub


