Attribute VB_Name = "Mdlexporasosykes"
Option Explicit


Const Version = "1.05"
Const FechaVersion = "05/09/2013"
'Autor = Mauricio Zwenger - CAS-20813 - Se modificó el formato de salida

'Const Version = "1.04"
'Const FechaVersion = "13/06/2013"
'Autor = Sebastian Stremel - al nro de legajo se agrego un 'delante para excel lo reconozca como texto - CAS-18586 - SYKES COSTA RICA - Exportacion AsoSykes


'Const Version = "1.03"
'Const FechaVersion = "30/05/2013"
'Autor = Sebastian Stremel - se obtiene los ultimos 5 caracteres del legajo en lugar de los primeros. - CAS-18586 - SYKES COSTA RICA - Exportacion AsoSykes


'Const Version = "1.02"
'Const FechaVersion = "13/03/2013"
'Autor = Gonzalez Nicolás - CAS-18586 - SYKES COSTA RICA - Exportacion AsoSykes


' Version 1.01
'Const Version = "1.01"
'Const FechaVersion = "14/01/2013"
'Autor = Carmen Quintero
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Global inx             As Integer
Global inxfin          As Integer

Global vec_testr1(50)  As Integer
Global vec_testr2(50)  As String
Global vec_testr3(50)  As String

Global vec_jor(50) As Single

Global Descripcion As String
Global Cantidad As Single

'Global nListaProc As Long
Global nListaProc As String
Global nEmpresa As Long
Global ArchExp


Public Sub Asosykes_OLD(ByVal bpronro As Long, ByVal Parametros As String)

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
Dim directorio
Dim Sep
Dim UsaEncabezado
Dim Nombre_Arch
Dim carpeta
Dim Linea As String

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
StrSql = "SELECT pliqmes,pliqanio FROM periodo WHERE pliqnro = " & Nroliq
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   liq_Mes = rs_aux!pliqmes
   liq_anio = rs_aux!pliqanio
End If
rs_aux.Close


'-------------------------------------------------------------------------------------------
' Configuracion Confrep nro 393
'-------------------------------------------------------------------------------------------

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 393 "
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
   directorio = Trim(rs_aux!sis_dirsalidas)
End If
rs_aux.Close
Flog.writeline directorio


'----------------------------------------------------------------------------
' Modelo
'----------------------------------------------------------------------------
StrSql = "SELECT * FROM modelo WHERE modnro = " & 356
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   If Not IsNull(rs_aux!modarchdefault) Then
      directorio = directorio & Trim(rs_aux!modarchdefault)
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

Nombre_Arch = directorio & "\expasosykes.csv"

Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch

Set fs = CreateObject("Scripting.FileSystemObject")
Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)

If Err.Number <> 0 Then
   Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
   Set carpeta = fs.CreateFolder(directorio)
   Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
End If


On Error GoTo 0
'--------------------------------------------------
If UsaEncabezado = -1 Then
    'IMPRIME ENCABEZADO
    Linea = "BADGE" & delimitador
    Linea = Linea & "CEDULA" & delimitador
    Linea = Linea & "NOMBRE EMPLEADO" & delimitador
    Linea = Linea & "AÑO" & delimitador
    Linea = Linea & "NUM_PERIODO" & delimitador
    Linea = Linea & "MONTO APLICADO" & delimitador
    
    ArchExp.Write Linea
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
    
    'Entra si se seleccionó una empresa
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
        Linea = Format_StrNro(Right(rs_aux!empleg, 5), 8, True, 0) & delimitador
        
        nrodoc = Replace(rs_aux!nrodoc, "-", "")
        nrodoc = Replace(nrodoc, ".", "")
        nrodoc = Format_StrNro(nrodoc, 10, True, 0)
        Linea = Linea & nrodoc & delimitador
        
        
        Linea = Linea & rs_aux!terape & " " & rs_aux!terape2 & " " & rs_aux!ternom & delimitador
        
        Linea = Linea & liq_anio & delimitador
        Linea = Linea & Format_StrNro(liq_Mes, 2, True, 0) & delimitador
        
        
        'Si es un concepto llama a la funcion concepto
        If conftipo = "CO" Then
            conc_aux = concepto(rs_aux!Ternro, confval2, Lista_Pro)
            
        'Si es AC llama a la funcion acumulador
        ElseIf conftipo = "AC" Then
            conc_aux = acumulador(rs_aux!Ternro, confval, liq_anio, liq_Mes)
        End If
        Linea = Linea & conc_aux
        
    
        'Incremento el progreso
        Progreso = Progreso + IncPorc
        'Inserto progreso
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
     
        'Flog.writelline linea
        If conc_aux = 0 Or conc_aux = 0 Then
        
        Else
            ArchExp.Write Linea
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

Public Sub Asosykes(ByVal bpronro As Long, ByVal Parametros As String)
'13/03/2013 - Gonzalez Nicolás - Se creo nuevo formato de salida.
'05/09/2013 - Mauricio Zwenger - CAS-20813 - Se modificó el formato de salida



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
Dim directorio
Dim Sep
Dim UsaEncabezado
Dim Nombre_Arch
Dim carpeta
Dim Linea As String

Dim delimitador

Dim arr_lista
Dim Nroliq
Dim Empresa
Dim pos1
Dim pos2
Dim Lista_Pro

Dim liq_Mes
Dim liq_anio
Dim pliq_desc
Dim conc_aux
Dim confval
Dim confval2
Dim conftipo

Dim ListaConfetiq As String
Dim ListaConftipo As String
Dim ListaConfval As String
Dim ListaConfval2 As String


Dim ArrConfetiq
Dim ArrConftipo
Dim ArrConfval
Dim ArrConfval2
Dim a
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
StrSql = "SELECT pliqmes,pliqanio,pliqdesc FROM periodo WHERE pliqnro = " & Nroliq
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   liq_Mes = rs_aux!pliqmes
   liq_anio = rs_aux!pliqanio
   pliq_desc = rs_aux!pliqdesc
End If
rs_aux.Close


'-------------------------------------------------------------------------------------------
' Configuracion Confrep nro 393
'-------------------------------------------------------------------------------------------

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 393  AND (conftipo = 'CO' OR conftipo = 'AC')"
OpenRecordset StrSql, rs_aux

ListaConfetiq = "0"
ListaConftipo = "0"
ListaConfval = "0"
ListaConfval2 = "0"

If Not rs_aux.EOF Then
    Do While Not rs_aux.EOF
        If rs_aux!confval2 <> "" Then
            
            'Guardo lista de etiquetas
            If rs_aux!confetiq <> "" Then
                ListaConfetiq = ListaConfetiq & "@" & rs_aux!confetiq
            End If

            ListaConftipo = ListaConftipo & "@" & rs_aux!conftipo
            ListaConfval = ListaConfval & "@" & rs_aux!confval
            ListaConfval2 = ListaConfval2 & "@" & rs_aux!confval2
            
        End If
        
        'datosconfrep = datosconfrep & "@" & rs_aux!confetiq & "!" & rs_aux!conftipo & "!" & rs_aux!confval2
'        confval = rs_aux!confval
'        confval2 = rs_aux!confval2
'        conftipo = rs_aux!conftipo
        rs_aux.MoveNext
    Loop
Else
    Flog.writeline "Error de configuración de reporte. Al menos debe configurar un Concepto o un Acumulador."
End If
rs_aux.Close


'Armo Array con el contenido
ArrConfetiq = Split(ListaConfetiq, "@")
ArrConftipo = Split(ListaConftipo, "@")
ArrConfval = Split(ListaConfval, "@")
ArrConfval2 = Split(ListaConfval2, "@")


'----------------------------------------------------------------------------
' Directorio de exportacion
'----------------------------------------------------------------------------
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   directorio = Trim(rs_aux!sis_dirsalidas)
End If
rs_aux.Close
Flog.writeline directorio


'----------------------------------------------------------------------------
' Modelo
'----------------------------------------------------------------------------
StrSql = "SELECT * FROM modelo WHERE modnro = " & 356
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   If Not IsNull(rs_aux!modarchdefault) Then
      directorio = directorio & Trim(rs_aux!modarchdefault)
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

Nombre_Arch = directorio & "\expasosykes.csv"

Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch

Set fs = CreateObject("Scripting.FileSystemObject")
Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)

If Err.Number <> 0 Then
   Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
   Set carpeta = fs.CreateFolder(directorio)
   Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
End If


On Error GoTo 0
'--------------------------------------------------
If UsaEncabezado = -1 Then
    'IMPRIME ENCABEZADO
    '05/09/2013 - MDZ - CAS-20813
    'Linea = "" & delimitador
    Linea = Linea & "Badge" & delimitador
    Linea = Linea & "Apellido y Nombre" & delimitador
   
'    For a = 1 To UBound(ArrConfetiq)
'        Linea = Linea & "MONTO" & delimitador
'    Next
    For a = 1 To UBound(ArrConfetiq)
        Linea = Linea & ArrConfetiq(a) & delimitador
    Next
    
    'saco el ultimo delimitador
    Linea = Left(Linea, Len(Linea) - 1)
'    Linea = Linea & "PERIODO" & delimitador
'    Linea = Linea & "AÑO" & delimitador
    ArchExp.Write Linea
    ArchExp.writeline ""
    
'    Linea = "" & delimitador
'    Linea = Linea & "" & delimitador
'    Linea = Linea & "" & delimitador
'
'   'Armo las columnas con su etiqueta
'    For a = 1 To UBound(ArrConfetiq)
'        Linea = Linea & ArrConfetiq(a) & delimitador
'    Next
'
'    Linea = Linea & "" & delimitador
'    Linea = Linea & "" & delimitador
'    ArchExp.Write Linea
'    ArchExp.writeline ""
    
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
    
    'Entra si se seleccionó una empresa
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
       ' Linea = Format_StrNro(Right(rs_aux!empleg, 5), 8, True, 0) & delimitador
        
        '05/09/2013 - MDZ - CAS-20813
        '1er Columna | Columna Vacía
        Linea = "" '& delimitador
        
        '2da Columna | 5 últimos digitos del legajo, se completa con 0
        Linea = Linea & ultimosCaracteres(rs_aux!empleg, 5, True, "0") & delimitador
        
        '3er Columna |Apellido y Nombre
        Linea = Linea & rs_aux!terape & " " & rs_aux!terape2 & " " & rs_aux!ternom & delimitador
        
        
       
        '4ta columna Hasta cantidad de CO/AC configurados por confrep
        For a = 1 To UBound(ArrConftipo)
            If ArrConftipo(a) = "CO" Then
                'Si es un concepto llama a la funcion concepto
                conc_aux = concepto(rs_aux!Ternro, ArrConfval2(a), Lista_Pro)
            ElseIf ArrConftipo(a) = "AC" Then
                'Si es AC llama a la funcion acumulador
                conc_aux = acumulador(rs_aux!Ternro, ArrConfval(a), liq_anio, liq_Mes)
            End If
            Linea = Linea & """" & Format(conc_aux, "#,##0.00") & """" & delimitador
        Next
           
        'saco el ultimo delimitador
        Linea = Left(Linea, Len(Linea) - 1)
       
'        '7ma Columna | Nombre del Periodo
'        Linea = Linea & pliq_desc & delimitador
'
'        '8va Columna | Año
'        Linea = Linea & liq_anio & delimitador
    
        'Incremento el progreso
        Progreso = Progreso + IncPorc
        'Inserto progreso
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
     
        'Flog.writelline linea
        If conc_aux = 0 Or conc_aux = 0 Then
        
        Else
            ArchExp.Write Linea
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
Public Function ultimosCaracteres(ByVal Str As String, ByVal Longitud As Long, ByVal Completar As Boolean, ByVal Str_Completar As String)
    If Not EsNulo(Str) Then
        Str = Right(Str, Longitud)
        If Completar Then
            If Len(Str) < Longitud Then
                Str = String(Longitud - Len(Str), Str_Completar) & Str
            End If
        End If
        'mdz
        ultimosCaracteres = CStr("'" & Str)
        ultimosCaracteres = CStr(Str)
    Else
        ultimosCaracteres = ""
    End If
End Function
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
' Autor      : Carmen Quintero
' Fecha      : 14/01/2013
' Descripcion:
' Ultima Mod.:
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
    
    Nombre_Arch = PathFLog & "Exportación_Asosykes" & "-" & NroProcesoBatch & ".log"
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 386 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Asosykes(NroProcesoBatch, bprcparam)
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



