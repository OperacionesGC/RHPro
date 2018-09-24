Attribute VB_Name = "Mdlexporemple"
Option Explicit

'Version: 1.01
' Exportación de Cafeteria

'Const Version = 1.01
'Const FechaVersion = "09/02/2011"
'Autor = Gonzalez Nicolás

'Const Version = "1.02"
'Const FechaVersion = "24/05/2011"
'Autor = Gonzalez Nicolás

'Const Version = "1.03"
'Const FechaVersion = "27/05/2011"
'Autor = Gonzalez Nicolás

'Const Version = "1.04"
'Const FechaVersion = "03/06/2011"
'Autor = FGZ

Const Version = "1.05"
Const FechaVersion = "29/06/2011"
'Autor = FGZ
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Global inx             As Integer
Global inxfin          As Integer

Global vec_testr1(50)  As Integer
Global vec_testr2(50)  As String
Global vec_testr3(50)  As String

Global vec_jor(50) As Single

Global Descripcion As String
Global cantidad As Single

'Global nListaProc As Long
Global nListaProc As String
Global nEmpresa As Long
Global ArchExp


Public Sub cafeteria(ByVal bpronro As Long, ByVal Parametros As String)

Dim rs_emple As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_direccion As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim objconnProgreso As New ADODB.Connection
OpenConnection strconexion, objconnProgreso

Dim ArregloParametros
Dim todos
Dim fecha_estruc
Dim Directorio
Dim Sep
Dim UsaEncabezado
Dim Nombre_Arch
Dim Carpeta
Dim linea As String

Dim delimitador

Dim arr_lista
Dim Nroliq
Dim Empresa
Dim pos1
Dim pos2
Dim Lista_Pro

Dim liq_Mes
Dim liq_anio
Dim conc_aux
Dim confval
Dim confval2
Dim conftipo
Dim nrodoc

Dim CEmpleadosAProc As Long
Dim IncPorc
Dim Progreso

'On Error GoTo CE

'----------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "@"
'----------------------------------------------------------------------------
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    Flog.writeline "If Len(Parametros) >= 1 Then"
       
        arr_lista = Split(Parametros, "@", -1, 1)
        Nroliq = arr_lista(0)
        Empresa = arr_lista(2)
        Lista_Pro = arr_lista(1)
        nListaProc = arr_lista(1)
Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline "terminó de levantar los parametros"

'-------------------
'TRAE MES Y ANIO DE LA LIQUIDACION
'-------------------
StrSql = "select pliqmes,pliqanio from periodo where pliqnro = " & Nroliq
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   liq_Mes = rs_aux!pliqmes
   liq_anio = rs_aux!pliqanio
End If
rs_aux.Close


'-------------------------------------------------------------------------------------------
' Configuracion Confrep nro 309
'-------------------------------------------------------------------------------------------

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 309 "
OpenRecordset StrSql, rs_aux

Do While Not rs_aux.EOF
    
    If Not rs_aux.EOF Then
        confval = rs_aux!confval
        confval2 = rs_aux!confval2
        conftipo = rs_aux!conftipo
    End If
    rs_aux.MoveNext
Loop
rs_aux.Close



'----------------------------------------------------------------------------
' Directorio de exportacion
'----------------------------------------------------------------------------
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   Directorio = Trim(rs_aux!sis_dirsalidas)
End If
rs_aux.Close
Flog.writeline Directorio


'----------------------------------------------------------------------------
' Modelo
'----------------------------------------------------------------------------
StrSql = "SELECT * FROM modelo WHERE modnro = " & 332
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   If Not IsNull(rs_aux!modarchdefault) Then
      Directorio = Directorio & Trim(rs_aux!modarchdefault)
      'Directorio = "C:\logs" & Trim(rs_aux!modarchdefault)
      
      'Si el delimitador esta vacio --> valor default ;
      If Not IsNull(rs_aux!modseparador) Then
        delimitador = rs_aux!modseparador
      Else
        delimitador = ";"
      End If
   Else
      Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino."
   End If
Else
   Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo 332. El archivo será generado en el directorio default"
End If

'Obtengo los datos del separador
Sep = rs_aux!modseparador
UsaEncabezado = rs_aux!modencab

Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep

On Error Resume Next

Nombre_Arch = Directorio & "\expcafeteria.csv"

Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch

Set fs = CreateObject("Scripting.FileSystemObject")
Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)

If Err.Number <> 0 Then
   Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
   Set Carpeta = fs.CreateFolder(Directorio)
   Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
End If


On Error GoTo 0
'--------------------------------------------------
If UsaEncabezado = -1 Then
    'IMPRIME ENCABEZADO
    linea = "BADGE" & delimitador
    linea = linea & "CEDULA" & delimitador
    linea = linea & "NOMBRE EMPLEADO" & delimitador
    linea = linea & "AÑO" & delimitador
    linea = linea & "NUM_PERIODO" & delimitador
    linea = linea & "MONTO APLICADO" & delimitador
    
    ArchExp.Write linea
    ArchExp.writeline ""
End If
   
    StrSql = "SELECT distinct Empleado.Ternro "
    StrSql = StrSql & " ,ter_doc.nrodoc,empleado.empleg,empleado.terape "
    StrSql = StrSql & " , empleado.ternom,empleado.terape2,ter_doc.tidnro "
    StrSql = StrSql & " FROM Empleado "
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE  "
    StrSql = StrSql & " (ter_doc.tidnro = 1 OR ter_doc.tidnro = 2 OR ter_doc.tidnro = 3) "
    StrSql = StrSql & " AND cabliq.pronro IN (" & Lista_Pro & ") "
    
    'Entra si se selecci   ono 1 empresa
    If Trim(Empresa) <> 1 Then
        StrSql = StrSql & " AND his_estructura.estrnro = " & Empresa
    End If
    
    StrSql = StrSql & " ORDER BY empleado.ternro "
    OpenRecordset StrSql, rs_aux
    
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs_aux.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = Int((99 / CEmpleadosAProc))
    
    
    Do While Not rs_aux.EOF
        'FGZ - 03/06/2011 -------------------------------------------------------
        'linea = Format_StrNro(rs_aux!empleg, 8, True, 0) & delimitador
        linea = Format_StrNro(Right(rs_aux!empleg, 5), 8, True, 0) & delimitador
        'FGZ - 03/06/2011 -------------------------------------------------------
        
        nrodoc = Replace(rs_aux!nrodoc, "-", "")
        nrodoc = Replace(nrodoc, ".", "")
        nrodoc = Format_StrNro(nrodoc, 10, True, 0)
        linea = linea & nrodoc & delimitador
        
        
        linea = linea & rs_aux!terape & " " & rs_aux!terape2 & " " & rs_aux!ternom & delimitador
        
        linea = linea & liq_anio & delimitador
        linea = linea & Format_StrNro(liq_Mes, 2, True, 0) & delimitador
        
        
        'Si es un concepto llama a la funcion concepto
        If conftipo = "CO" Then
            conc_aux = concepto(rs_aux!Ternro, confval2, Lista_Pro)
            
        'Si es AC llama a la funcion acumulador
        ElseIf conftipo = "AC" Then
            conc_aux = acumulador(rs_aux!Ternro, confval, liq_anio, liq_Mes)
        End If
        linea = linea & conc_aux
        
        
        
        'Incremento el progreso
        Progreso = Progreso + IncPorc
        'Inserto progreso
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
     
        'Flog.writelline linea
        If conc_aux = 0 Or conc_aux = 0 Then
        
        Else
            ArchExp.Write linea
            ArchExp.writeline ""
        End If
        rs_aux.MoveNext
     Loop



    'Redondeo a 100%
    If Int(Progreso) < 100 Then
        'Inserto progreso
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100"
        StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If

End Sub






Public Function concepto(ByVal idemp, ByVal conc, ByVal Proceso)
    
    Dim rs_aux As New ADODB.Recordset
    
    StrSql = "SELECT cabliq.empleado, SUM(dlimonto) dlimonto "
    StrSql = StrSql & "From detliq "
    StrSql = StrSql & "INNER JOIN cabliq ON  cabliq.cliqnro = detliq.cliqnro "
    StrSql = StrSql & "INNER JOIN concepto ON concepto.concnro = detliq.concnro "
    StrSql = StrSql & "WHERE empleado =  " & idemp
    StrSql = StrSql & " AND pronro IN (" & Proceso & ")"
    StrSql = StrSql & " AND (concepto.conccod = '" & conc & "')"
    StrSql = StrSql & " GROUP BY cabliq.empleado"
    OpenRecordset StrSql, rs_aux
    
    If Not rs_aux.EOF Then
        'concepto = Abs(rs_aux!dlimonto) * 100
        concepto = Abs(rs_aux!dlimonto)
    Else
        concepto = 0
    End If
    
    'rs.Close

End Function

Public Function acumulador(ByVal idemp, ByVal acuNro, ByVal Anio, ByVal Mes)
    
    Dim rs_aux As New ADODB.Recordset
    
    StrSql = "SELECT ternro,SUM(ammontoreal) ammontoreal "
    StrSql = StrSql & " FROM acu_mes "
    StrSql = StrSql & " WHERE acunro = " & acuNro
    StrSql = StrSql & " AND Ternro = " & idemp
    StrSql = StrSql & " AND amanio = " & Anio
    StrSql = StrSql & " AND ammes = " & Mes
    StrSql = StrSql & " GROUP BY  ternro "
    OpenRecordset StrSql, rs_aux
    
    If Not rs_aux.EOF Then
        'acumulador = Abs(rs_aux!ammontoreal) * 100
        acumulador = Abs(rs_aux!ammontoreal)
    Else
        acumulador = 0
    End If
    
    'rs.Close

End Function

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de la Exportación de Empleados
' Autor      : Gonzalez Nicolás
' Fecha      : 08/02/2011
' Descripcion:
' Ultima Mod.: 24/05/2011 - Gonzalez Nicolás
'             - Se cambio String de N° de Legajo de 6 a 8
'             - Se cambio formato del archivo de exportación de .txt  a .csv
'             - Se modificaron las funciones concepto() y acumulador () se quito *100
'             27/05/2011 - Gonzalez Nicolás - Se agregó Progreso
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Exportación_Cafeteria" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 293 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call cafeteria(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub



