Attribute VB_Name = "MdlInterface"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "29/09/2005"
'Global Const UltimaModificacion = "Inicial"

Global Const Version = "1.01"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "Encriptacion de string connection"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
Global HuboError As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
Global usuario As String

Global Separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean
Global NroModelo As Integer
Global DescripcionModelo As String
Global NombreArchivo As String

Global Novedades As Boolean
Global Indice As Long
Global ListaNovedades(200000)
Global CantidadLineas As Long

'04/10/2004
'Dim objFeriado As New Feriado


Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial de Interface.
    ' Autor      : JMH
    ' Fecha      : 29/07/2004
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim objconnMain As New ADODB.Connection
    Dim strCmdLine
    Dim Nombre_Arch As String
    Dim rs_batch_proceso As New ADODB.Recordset
    Dim bprcparam As String
    Dim PID As String
    Dim ArrParametros

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    'Abro la conexion
    'OpenConnection strconexion, objConn
    'OpenConnection strconexion, objconnProgreso
    
        
    Nombre_Arch = PathFLog & "Interface_Redepa" & "-" & NroProcesoBatch & ".log"
    'Archivo de log
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
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 113 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    'Primera_Vez = True
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        Flog.writeline Espacios(Tabulador * 0) & "Parametros del proceso = " & bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
        Call ComenzarTransferencia
    End If
    
    If Not HuboError Then
        If ErroresNov Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        End If
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Resumen
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "===================================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Lineas Leidas    : " & RegLeidos
    Flog.writeline Espacios(Tabulador * 0) & "Lineas Procesadas: " & RegLeidos - RegError
    Flog.writeline Espacios(Tabulador * 0) & "===================================================================="
    objConn.Close
    objconnProgreso.Close
    Flog.Close

End Sub


Private Sub LeeArchivo(ByVal NombreArchivo As String)
    Const ForReading = 1
    Const TristateFalse = 0
    Dim strlinea As String
    Dim Archivo_Aux As String
    Dim rs_Lineas As New ADODB.Recordset
    Dim rs_Modelo As New ADODB.Recordset

    
    Const adStateOpen = &H1
    Const adChapter = 136
    
    If App.PrevInstance Then
        Flog.writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo "
        Exit Sub
    End If

    'Espero hasta que se crea el archivo
    'On Error Resume Next
    'Err.Number = 1
    'Do Until Err.Number = 0
    '    Err.Number = 0
    '    Set f = fs.getfile(NombreArchivo)
    '    If f.Size = 0 Then
    '        Flog.Writeline Espacios(Tabulador * 0) & "No anda el getfile "
    '        Err.Number = 1
    '    End If
    'Loop
    On Error GoTo 0
    'Flog.Writeline Espacios(Tabulador * 0) & "Archivo creado " & NombreArchivo

    Call Cantidad_Lineas(NombreArchivo)
    
    'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
        Flog.writeline Espacios(Tabulador * 0) & "Ultimo inter_pin " & crpNro
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se pudo abrir el archivo " & NombreArchivo
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If rs_Modelo.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No esta el modelo " & NroModelo
        Exit Sub
    End If
                
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = CantidadLineas
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (50 / CEmpleadosAProc)
    
    Do While Not f.AtEndOfStream
        strlinea = f.ReadLine
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strlinea = f.ReadLine
            'NroLinea = NroLinea + 1
            'rs_Lineas.MoveNext
        End If
        
        If Trim(strlinea) <> "" Then
        
            RegLeidos = RegLeidos + 1
            Call Insertar_Linea(strlinea)
                
            'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
            'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
            'como colgado
            Progreso = Progreso + IncPorc
            If Progreso > 50 Then Progreso = 50
            Flog.writeline Espacios(Tabulador * 0) & "Progreso = " & CLng(Progreso) & " (Incremento = " & IncPorc & ")"
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 0) & "Progreso actualizado"
        End If
        'If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
    Loop

    If Novedades = True Then
       Call Crear_Novedades
    End If
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    'fs.Deletefile NombreArchivo, True
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Insertar_Linea(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interfece Redepa
' Autor      : JMH
' Fecha      : 28/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Divisor As String
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date
Dim ThSigla As String
Dim TipoHora As Long
Dim TdSigla As String
Dim TipoDia As Long

Dim horas_lic As Double
Dim Legajo As Long
Dim Elemento As String
Dim Tercero As Long
Dim Cantidad As Double
Dim Valor1 As Double
Dim Valor2 As Double
Dim dia As Integer
Dim Mes As Integer
Dim Anio As Integer

Dim rs_TipoHora As New ADODB.Recordset
Dim rs_TipoDia As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

    'Divisor
    Divisor = Mid$(strlinea, 1, 2)
    
    'Legajo
    Legajo = CLng(Mid$(strlinea, 7, 6))
    
    StrSql = "SELECT * FROM empleado WHERE empleado.empleg = " & Legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
       Flog.writeline Espacios(Tabulador * 1) & "Empleado Desconocido"
       FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Empleado Desconocido"
       InsertaError 0, 8
       HuboError = True
       Exit Sub
    Else
        Tercero = rs_Empleado!ternro
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "Divisor: " & Divisor
    Select Case Divisor   'Evalúa divisor.
        Case "90"
           Novedades = True
           ThSigla = Mid$(strlinea, 22, 3)
           Flog.writeline Espacios(Tabulador * 1) & "Sigla: " & ThSigla
           
           StrSql = "SELECT * FROM tiphora WHERE tiphora.thsigla = '" & ThSigla & "'"
           OpenRecordset StrSql, rs_TipoHora
           If rs_TipoHora.EOF Then
              Flog.writeline Espacios(Tabulador * 1) & "Tipo de Hora Desconocida"
              FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Tipo de Hora Desconocida"
              InsertaError 0, 37
              HuboError = True
              Exit Sub
           Else
               TipoHora = rs_TipoHora!thnro
               Tercero = rs_Empleado!ternro
               Valor1 = CDbl(Mid$(strlinea, 46, 4))
               Valor2 = CDbl(Mid$(strlinea, 50, 4))
               Cantidad = (Valor1 + Valor2 / 1000)
               
               Elemento = CStr(TipoHora) & "@" & CStr(Tercero) & "@" & CStr(Cantidad)
               ListaNovedades(Indice) = Elemento
               Indice = Indice + 1
               
           End If
           rs_Empleado.Close
    
        Case "AL"
           TdSigla = Mid$(strlinea, 22, 3)
           Flog.writeline Espacios(Tabulador * 1) & "Sigla: " & TdSigla
           
           StrSql = "SELECT * FROM tipdia WHERE tipdia.tdsigla = '" & TdSigla & "'"
           OpenRecordset StrSql, rs_TipoDia
           If rs_TipoDia.EOF Then
              Flog.writeline Espacios(Tabulador * 1) & "Tipo de Dia Desconocido"
              FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Tipo de Dia Desconocido"
              InsertaError 0, 84
              HuboError = True
              Exit Sub
           Else
               If (rs_TipoDia!tdnro = 2) Then
                   Flog.writeline Espacios(Tabulador * 1) & "Licencia Por Vacaciones"
                   FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Licencia Por Vacaciones"
                   HuboError = True
                   Exit Sub
               End If
               
               Mes = CInt(Mid$(strlinea, 27, 2))
               dia = CInt(Mid$(strlinea, 25, 2))
               Anio = CInt(Mid$(strlinea, 29, 4))
               'Fecha_Desde = CDate(Mes & "/" & Dia & "/" & Anio)
               Fecha_Desde = CDate(dia & "/" & Mes & "/" & Anio)
               Mes = CInt(Mid$(strlinea, 35, 2))
               dia = CInt(Mid$(strlinea, 33, 2))
               Anio = CInt(Mid$(strlinea, 37, 4))
               'Fecha_Hasta = CDate(Mes & "/" & Dia & "/" & Anio)
               Fecha_Hasta = CDate(dia & "/" & Mes & "/" & Anio)
               horas_lic = CLng(Mid$(strlinea, 41, 3)) + (CLng(Mid$(strlinea, 44, 2)) / 100)
               
               Call Crear_Licencia(Fecha_Desde, Fecha_Hasta, horas_lic, Tercero, rs_TipoDia!tdnro)
                              
           End If
           'rs_Empleado.Close
    End Select
    'rs_Empleado.Close

'cierro y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_TipoHora.State = adStateOpen Then rs_TipoHora.Close
If rs_TipoDia.State = adStateOpen Then rs_TipoDia.Close

Set rs_Empleado = Nothing
Set rs_TipoHora = Nothing
Set rs_TipoDia = Nothing
End Sub

Public Sub Crear_Licencia(ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal Horas As Double, ByVal NroTer As Long, ByVal NroTDia As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Crea licencia
' Autor      : JMH
' Fecha      : 28/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim cantDias As Long
Dim rs_EmpLic As New ADODB.Recordset

Flog.writeline Espacios(Tabulador * 2) & "Inicio Crear Licencia ..."

StrSql = "SELECT * FROM emp_lic "
StrSql = StrSql & " WHERE emp_lic.empleado = " & NroTer
StrSql = StrSql & " AND emp_lic.tdnro = " & NroTDia
StrSql = StrSql & " AND ((emp_lic.elfechadesde >= " & ConvFecha(FechaDesde) & " AND emp_lic.elfechadesde <= " & ConvFecha(FechaHasta) & ") OR "
StrSql = StrSql & "      (emp_lic.elfechahasta >= " & ConvFecha(FechaDesde) & " AND emp_lic.elfechahasta <= " & ConvFecha(FechaHasta) & ") OR "
StrSql = StrSql & "      (emp_lic.elfechadesde <= " & ConvFecha(FechaDesde) & " AND emp_lic.elfechahasta >= " & ConvFecha(FechaHasta) & ") OR "
StrSql = StrSql & "      (emp_lic.elfechadesde >= " & ConvFecha(FechaDesde) & " AND emp_lic.elfechahasta <= " & ConvFecha(FechaHasta) & "))"
OpenRecordset StrSql, rs_EmpLic

If rs_EmpLic.EOF Then
        cantDias = DateDiff("d", FechaDesde, FechaHasta) + 1
        StrSql = "INSERT INTO emp_lic ("
        StrSql = StrSql & "tdnro,eldiacompleto,elfechadesde,elfechahasta,empleado,elcantdias,elmaxhoras,eltipo"
        StrSql = StrSql & ") VALUES (" & NroTDia
        StrSql = StrSql & ",-1"
        StrSql = StrSql & "," & ConvFecha(FechaDesde)
        StrSql = StrSql & "," & ConvFecha(FechaHasta)
        StrSql = StrSql & "," & NroTer
        StrSql = StrSql & "," & cantDias
        StrSql = StrSql & "," & Horas
        StrSql = StrSql & ",1"
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Flog.writeline "Licencia insertada "
Else
        Flog.writeline Espacios(Tabulador * 1) & "Superposición de Fecha en la Licencia: " & CStr(NroTDia)
        InsertaError 0, 99
        HuboError = True
        Exit Sub
End If
Flog.writeline Espacios(Tabulador * 2) & "Fin Crear Licencia ..."
End Sub

Public Sub Crear_Novedades()
' ---------------------------------------------------------------------------------------------
' Descripcion: Crea Novedades
' Autor      : JMH
' Fecha      : 29/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Desde As Date
Dim Hasta As Date
Dim Horas As Double
Dim i As Long
Dim ArrElemento
Dim EmpleadoAnt As Long
Dim ConceptoAnt As Long
Dim ParametroAnt As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset


Flog.writeline Espacios(Tabulador * 2) & "Inicio Crear Novedad ..."
'Calculo el resto del progreso
Progreso = 50
CEmpleadosAProc = Indice - 1

If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
IncPorc = (50 / CEmpleadosAProc)

For i = 0 To Indice - 1
    ArrElemento = Split(ListaNovedades(i), "@")
    Horas = 0
    StrSql = "SELECT tiph_con.thsuma, empleado.ternro, tiph_con.concnro, tiph_con.tpanro "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tiph_con ON tiph_con.thnro = " & ArrElemento(0)
    StrSql = StrSql & " INNER JOIN tiphora_estr ON tiphora_estr.thnro = tiph_con.thnro "
    StrSql = StrSql & " INNER JOIN his_estructura he ON he.ternro = empleado.ternro "
    StrSql = StrSql & " AND he.tenro = tiphora_estr.tenro AND he.estrnro = tiphora_estr.estrnro "
    StrSql = StrSql & " AND he.htethasta is null "
    StrSql = StrSql & " INNER JOIN tipopar ON tipopar.tpanro = tiph_con.tpanro "
    StrSql = StrSql & " WHERE empleado.ternro = " & ArrElemento(1)
    OpenRecordset StrSql, rs_Empleado
    
    If Not rs_Empleado.EOF Then
       EmpleadoAnt = rs_Empleado!ternro
       ConceptoAnt = rs_Empleado!concnro
       ParametroAnt = rs_Empleado!tpanro
    End If
    
    Do While Not rs_Empleado.EOF
    
       If (EmpleadoAnt <> rs_Empleado!ternro Or ConceptoAnt <> rs_Empleado!concnro Or _
           ParametroAnt <> rs_Empleado!tpanro) And (Horas <> 0) Then
          
           StrSql = "SELECT * "
           StrSql = StrSql & " FROM novemp "
           StrSql = StrSql & " WHERE novemp.empleado = " & EmpleadoAnt
           StrSql = StrSql & " AND novemp.concnro = " & ConceptoAnt
           StrSql = StrSql & " AND novemp.tpanro = " & ParametroAnt
           OpenRecordset StrSql, rs_NovEmp
           
           If rs_NovEmp.EOF Then
              StrSql = "INSERT INTO novemp ("
              StrSql = StrSql & "empleado,concnro,tpanro,nevalor"
              StrSql = StrSql & ") VALUES (" & EmpleadoAnt
              StrSql = StrSql & "," & ConceptoAnt
              StrSql = StrSql & "," & ParametroAnt
              StrSql = StrSql & "," & Horas
              StrSql = StrSql & " )"
              objConn.Execute StrSql, , adExecuteNoRecords
           Else
              StrSql = "UPDATE novemp "
              StrSql = StrSql & " SET nevalor= " & Horas
              StrSql = StrSql & " WHERE novemp.empleado = " & EmpleadoAnt
              StrSql = StrSql & " AND novemp.concnro = " & ConceptoAnt
              StrSql = StrSql & " AND novemp.tpanro = " & ParametroAnt
              objConn.Execute StrSql, , adExecuteNoRecords
           End If
           
           Horas = 0
           EmpleadoAnt = rs_Empleado!ternro
           ConceptoAnt = rs_Empleado!concnro
           ParametroAnt = rs_Empleado!tpanro
       Else
           If rs_Empleado!thsuma = -1 Then
              Horas = Horas + CDbl(ArrElemento(2))
           End If
           rs_Empleado.MoveNext
       End If
       
    Loop
    
    If rs_Empleado.EOF And Horas <> 0 Then
           StrSql = "SELECT * "
           StrSql = StrSql & " FROM novemp "
           StrSql = StrSql & " WHERE novemp.empleado = " & EmpleadoAnt
           StrSql = StrSql & " AND novemp.concnro = " & ConceptoAnt
           StrSql = StrSql & " AND novemp.tpanro = " & ParametroAnt
           OpenRecordset StrSql, rs_NovEmp
           
           If rs_NovEmp.EOF Then
              StrSql = "INSERT INTO novemp ("
              StrSql = StrSql & "empleado,concnro,tpanro,nevalor"
              StrSql = StrSql & ") VALUES (" & EmpleadoAnt
              StrSql = StrSql & "," & ConceptoAnt
              StrSql = StrSql & "," & ParametroAnt
              StrSql = StrSql & "," & Horas
              StrSql = StrSql & " )"
              objConn.Execute StrSql, , adExecuteNoRecords
           Else
              StrSql = "UPDATE novemp "
              StrSql = StrSql & " SET nevalor= " & Horas
              StrSql = StrSql & " WHERE novemp.empleado = " & EmpleadoAnt
              StrSql = StrSql & " AND novemp.concnro = " & ConceptoAnt
              StrSql = StrSql & " AND novemp.tpanro = " & ParametroAnt
              objConn.Execute StrSql, , adExecuteNoRecords
           End If
    End If
       
    Progreso = Progreso + IncPorc
    Flog.writeline Espacios(Tabulador * 0) & "Progreso = " & CLng(Progreso) & " (Incremento = " & IncPorc & ")"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 0) & "Progreso actualizado"
Next i

Flog.writeline Espacios(Tabulador * 2) & "Fin Crear Novedad ..."
End Sub

Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer


Separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        If pos2 > 0 Then
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
    End If
End If

End Sub

Public Sub Cantidad_Lineas(ByVal NombreArchivo As String)

Const ForReading = 1
Const TristateFalse = 0

Dim pos1 As Integer
Dim pos2 As Integer
Dim strlinea

'Abro el archivo
  Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
  CantidadLineas = 0
  Do While Not f.AtEndOfStream
     strlinea = f.ReadLine
     CantidadLineas = CantidadLineas + 1
  Loop
  
  f.Close
  
End Sub


Public Sub ComenzarTransferencia()
    Dim Directorio As String
    Dim CArchivos
    Dim Archivo
    Dim Folder
    Dim fc, F1, s2

    'Leo los datos del Sistema
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    'Leo los datos del modelo
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        Separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & NroModelo & " " & objRs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de importación :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    Progreso = 0
    HuboError = False
    NArchivo = Directorio & "\" & NombreArchivo
    Flog.writeline Espacios(Tabulador * 1) & "Archivo Procesado: " & NombreArchivo
    Call LeeArchivo(NArchivo)
    
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub


