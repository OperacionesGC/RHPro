Attribute VB_Name = "ExportacionEmpleadoCCU"
Option Explicit

'Version: 1.01
'   Exportación de Empleados CCU

'Global Const Version = "1.01"
'Global Const FechaVersion = "29/03/2012"
'Global Const UltimaModificacion = "Carmen Quintero (14753) Se creo nuevo proceso de exportacion de empleados CCU."
'Global Const UltimaModificacion1 = ""


Global Const Version = "1.02"
Global Const FechaVersion = "15/09/2014"
Global Const UltimaModificacion = " Se agrega carpetas por usuarios - CAS-24538 - CCU - MEJORA EN SEGURIDAD EN IN-OUT"
Global Const UltimaModificacion1 = "Sebastian Stremel "

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
Global iduser

'----------------------------------------------------------

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de la Exportación de Empleados
' Autor      : Raul CHinestra
' Fecha      : 01/09/2006
' Ultima Mod.:
' Descripcion:
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

    On Error Resume Next
    
    Nombre_Arch = PathFLog & "Exportación_Empleados" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion:              " & UltimaModificacion
    Flog.writeline "                           " & UltimaModificacion1
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 368 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        iduser = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ExpEmp(NroProcesoBatch, bprcparam)
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


Public Sub ExpEmp(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte de Exportacion de Empleados
' Autor      : RCH
' Fecha      : 27/09/2006
' Modificado :
' --------------------------------------------------------------------------------------------

Dim ArregloParametros

Dim todos As Integer
Dim fecha_estruc As Date
Dim rs_emple As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_direccion As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Apellido As String
Dim nombre As String
Dim FecNac As String
Dim PaisNac As String
Dim FecIng As String
Dim Estciv As String
Dim Sexo As String
Dim FecAlt As String
Dim Nivest As String
Dim TipoDocu As String
Dim NroDocu As String
Dim Cuil As String
Dim Calle As String
Dim Nro  As String
Dim Piso As String
Dim Depto As String
Dim Torre As String
Dim Manzana As String
Dim CodPostal As String
Dim Entre As String
Dim Barrio As String
Dim Loca As String
Dim Prov As String
Dim Pais As String
Dim Tele As String
Dim Sucu As String
Dim Conv As String
Dim Cate As String
Dim Puesto As String
Dim Centro As String
Dim Caja As String
Dim Sindi As String
Dim Osocial As String
Dim PlanOsocial As String
Dim EstEmp As String
Dim Causa As String
Dim FecBaj As String
Dim Empre As String
Dim Remu As String
Dim Orga As String
Dim Gerencia As String
Dim Contrato As String
Dim GrupSeg As String

Dim directorio As String
Dim Nombre_Arch As String
Dim fs1
Dim carpeta
Dim totalEmpleados
Dim cantRegistros
Dim Linea As String
Dim Niveles As String
Dim Tidonro As Integer
Dim Sep As String
Dim UsaEncabezado As Boolean

On Error GoTo CE

TiempoAcumulado = GetTickCount

'----------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "."
'----------------------------------------------------------------------------

Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    Flog.writeline "If Len(Parametros) >= 1 Then"
    
    If Len(Parametros) >= 1 Then
    
        ArregloParametros = Split(Parametros, ".")
           
        todos = ArregloParametros(0)
        Flog.writeline "Parametro todos = " & todos
        
        fecha_estruc = ArregloParametros(1)
        Flog.writeline "Parametro fecha estructuras = " & fecha_estruc
    End If

Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline "terminó de levantar los parametros"


'----------------------------------------------------------------------------
' Directorio de exportacion
'----------------------------------------------------------------------------
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   directorio = Trim(rs_aux!sis_dirsalidas)
    If "\" <> CStr(Right(directorio, 1)) Then
        directorio = directorio & "\"
    End If
End If
rs_aux.Close

'---------------------------------------------------------------------------
'Agrego la carpeta por usuario
directorio = directorio & "porUsr\" & iduser
'---------------------------------------------------------------------------

'----------------------------------------------------------------------------
' Modelo
'----------------------------------------------------------------------------

StrSql = "SELECT * FROM modelo WHERE modnro = " & 311
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   If Not IsNull(rs_aux!modarchdefault) Then
      directorio = directorio & Trim(rs_aux!modarchdefault)
   Else
      Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
   End If
Else
   Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo 311. El archivo será generado en el directorio default"
End If



'Obtengo los datos del separador
Sep = rs_aux!modseparador
UsaEncabezado = rs_aux!modencab

Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep


On Error Resume Next

Nombre_Arch = directorio & "\exportacion-empleados.csv"
Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
Set fs = CreateObject("Scripting.FileSystemObject")
Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)

If Err.Number <> 0 Then
   Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
   Set carpeta = fs.CreateFolder(directorio)
   Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
End If

'desactivo el manejador de errores

On Error GoTo 0


'-------------------------------------------------------------------------------------------
' Busco los Niveles de Estudio Configurados en el Confrep nro 178
'-------------------------------------------------------------------------------------------

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 178 "
StrSql = StrSql & " AND conftipo = 'NIV'"
OpenRecordset StrSql, rs_aux
Niveles = "0," ' Se inicializa en 0 para que no de error en el caso de haber configurado mal el confrep
Do While Not rs_aux.EOF
    Niveles = Niveles & "'" & rs_aux!confval & "'"
    rs_aux.MoveNext
    If Not rs_aux.EOF Then
        Niveles = Niveles & ","
    End If
Loop
rs_aux.Close


'-------------------------------------------------------------------------------------------
' Busco el Tipo de Domicilio en el Confrep nro 178
'-------------------------------------------------------------------------------------------

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 178 "
StrSql = StrSql & " AND conftipo = 'DOM'"
OpenRecordset StrSql, rs_aux
Tidonro = 2 ' POr default coloco el tipo de domicilio particular
If Not rs_aux.EOF Then
    Tidonro = rs_aux!confval
End If
rs_aux.Close


If todos = 0 Then
    StrSql = "SELECT * FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
    StrSql = StrSql & " WHERE bpronro = " & bpronro
Else
    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " WHERE empest = -1 "
End If

OpenRecordset StrSql, rs_emple

cantRegistros = rs_emple.RecordCount
totalEmpleados = rs_emple.RecordCount

StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
            ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & bpronro

objConn.Execute StrSql, , adExecuteNoRecords

Flog.writeline "Empiezo a Recorrer los empleados seleccionados = "


'-------------------------------------------------------------------------------------------
' Si se imprime el Encabezado
'-------------------------------------------------------------------------------------------

If UsaEncabezado Then

    '-------------------------------------------------------------------------------------------
    ' Exporto a un CSV
    '-------------------------------------------------------------------------------------------
    
    Linea = "Legajo " & Sep & "Apellido" & Sep & "nombre" & Sep & "FecNac" & Sep & "PaisNac" & Sep & "FecIng"
    Linea = Linea & Sep & "Estciv" & Sep & "Sexo" & Sep & "FecAlt" & Sep & "Niv. Estudio" & Sep & "TipoDocu" & Sep & "NroDocu"
    Linea = Linea & Sep & "Cuil" & Sep & "Calle" & Sep & "Nro" & Sep & "Piso" & Sep & "Depto"
    Linea = Linea & Sep & "Torre" & Sep & "Manzana" & Sep & "CodPostal" & Sep & "Entre" & Sep & "Barrio"
    Linea = Linea & Sep & "Loca" & Sep & "Prov" & Sep & "Pais" & Sep & "Tele" & Sep & "Sucu"
    Linea = Linea & Sep & "Conv" & Sep & "Cate" & Sep & "Puesto" & Sep & "Centro" & Sep & "Caja" & Sep & "Sindi"
    Linea = Linea & Sep & "Osocial" & Sep & "PlanOsocial" & Sep & "EstEmp" & Sep & "Causa" & Sep & "FecBaj" & Sep & "Empre" & Sep & "Remu" & Sep & "Orga" & Sep & "Gerencia" & Sep & "Contrato" & Sep & "GrupSeg"

    ArchExp.Write Linea
    ArchExp.writeline ""
    
End If

Do While Not rs_emple.EOF
        

    '----------------------------------------------------------------
    ' Buscar el apellido y nombre
    '----------------------------------------------------------------
    
    StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_emple!Ternro
    
    OpenRecordset StrSql, rs_Tercero
    
    If Not rs_Tercero.EOF Then
           
            Flog.writeline "Empleado (ternro) = " & rs_emple!Ternro
            
            '----------------------------------------------------------------
            ' 1 - Legajo del Empleado
            '----------------------------------------------------------------
            
            Legajo = rs_emple("empleg")
            Flog.writeline " 1 -- Legajo = " & Legajo
            
                '----------------------------------------------------------------
                ' 2- Buscar el Apellido
                '----------------------------------------------------------------
             
                Apellido = rs_Tercero!terape & " " & rs_Tercero!terape2
                Flog.writeline " 2 -- Apellido = " & Apellido
                
                '----------------------------------------------------------------
                ' 3 - Buscar el Nombre
                '----------------------------------------------------------------
                
                nombre = rs_Tercero!ternom & " " & rs_Tercero!ternom2
                Flog.writeline " 3 -- Nombres = " & nombre
                
                '----------------------------------------------------------------
                ' 4 - Buscar la fecha de Nacimiento
                '----------------------------------------------------------------
                
                If Not IsNull(rs_Tercero!terfecnac) Then
                    FecNac = CStr(rs_Tercero!terfecnac)
                Else
                    FecNac = ""
                End If
                Flog.writeline " 4 -- Fec Nac = " & FecNac
                        
                '----------------------------------------------------------------
                ' 5 - Buscar el Pais de Nacimiento
                '----------------------------------------------------------------
                          
                If Not IsNull(rs_Tercero!paisnro) Then
                    StrSql = "SELECT paisdesc FROM pais WHERE paisnro = " & rs_Tercero!paisnro
                    OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                        PaisNac = rs_aux!paisdesc
                        Flog.writeline " 5 -- Pais Nac = " & PaisNac
                    Else
                        PaisNac = ""
                        Flog.writeline " 5 -- Pais Nac= NO encontrado"
                    End If
                    rs_aux.Close
                Else
                    PaisNac = ""
                    Flog.writeline " 5 -- Pais Nac= NO encontrado"
                End If
                
                '----------------------------------------------------------------
                ' 6 - Buscar la fecha de Ingreso al Pais
                '----------------------------------------------------------------
               
                If IsNull(rs_Tercero!terfecing) Then
                    FecIng = ""
                Else
                    If rs_Tercero!terfecing <> CDate("01/01/1900") Then
                       FecIng = rs_Tercero!terfecing
                    Else
                       FecIng = ""
                    End If
                End If
                 
                 Flog.writeline " 6 -- Fec Ing = " & FecIng
            
                '----------------------------------------------------------------
                ' 7 - Buscar el Estado Civil
                '----------------------------------------------------------------
                                  
                If Not IsNull(rs_Tercero!estcivnro) Then
                    StrSql = "SELECT estcivdesabr FROM estcivil WHERE estcivnro = " & rs_Tercero!estcivnro
                    OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                        Estciv = rs_aux!estcivdesabr
                        Flog.writeline " 7 -- Est Civil = " & Estciv
                    Else
                        Estciv = ""
                        Flog.writeline " 7 -- Est Civil = NO encontrado"
                    End If
                    rs_aux.Close
                Else
                    Estciv = ""
                    Flog.writeline " 7 -- Est Civil = NO encontrado"
                End If
                
                '----------------------------------------------------------------
                ' 8 - Buscar el sexo
                '----------------------------------------------------------------
                If rs_Tercero!tersex = -1 Then
                    Sexo = "M"
                Else
                    Sexo = "F"
                End If
                Flog.writeline " 8 -- Sexo = " & Sexo
                
                '----------------------------------------------------------------
                ' 9 - Buscar la fecha de Alta
                '----------------------------------------------------------------
                
                If IsNull(rs_emple!empfaltagr) Then
                    FecAlt = ""
                Else
                    FecAlt = rs_emple!empfaltagr
                End If
                Flog.writeline " 9 -- Fec Alt = " & FecAlt
                
                '----------------------------------------------------------------
                ' 10  Nivel de Estudio
                '----------------------------------------------------------------
                        
                StrSql = " SELECT * FROM cap_estformal " & _
                         " INNER JOIN nivest ON nivest.nivnro = cap_estformal.nivnro  " & _
                         " WHERE nivest.nivnro IN (" & Niveles & ")" & _
                         " AND cap_estformal.ternro= " & rs_Tercero!Ternro
                OpenRecordset StrSql, rs_aux
                Nivest = ""
                Do While Not rs_aux.EOF
                    Nivest = Nivest & rs_aux!nivdesc & " "
                    rs_aux.MoveNext
                Loop
                Flog.writeline " 10 -- Niv Estudio = " & Nivest
                rs_aux.Close
                        
                '----------------------------------------------------------------
                ' 11 , 12 - Tipo y Nro de Documento
                '----------------------------------------------------------------
                        
                StrSql = " SELECT ter_doc.nrodoc, tipodocu.tidsigla FROM tercero " & _
                         " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro <= 4) " & _
                         " INNER JOIN tipodocu ON (ter_doc.tidnro = tipodocu.tidnro ) " & _
                         " WHERE tercero.ternro= " & rs_Tercero!Ternro
                OpenRecordset StrSql, rs_aux
                If Not rs_aux.EOF Then
                    TipoDocu = rs_aux!tidsigla
                    NroDocu = CStr(rs_aux!NroDoc)
                    Flog.writeline " 11,12 -- Tipo y Nro Doc = " & TipoDocu & " " & NroDocu
                Else
                    TipoDocu = ""
                    NroDocu = ""
                    Flog.writeline " 11, 12 -- Tipo y Nro Doc = NO encontrado "
                End If
                rs_aux.Close
                
                '----------------------------------------------------------------
                ' 13 - CUIL
                '----------------------------------------------------------------
                        
                StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                         " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
                         " WHERE tercero.ternro= " & rs_Tercero!Ternro
                OpenRecordset StrSql, rs_aux
                If Not rs_aux.EOF Then
                    Cuil = Left(CStr(rs_aux!NroDoc), 13)
                    Flog.writeline " 13 -- CUIL = " & Cuil
                Else
                    Cuil = ""
                    Flog.writeline " 13 -- CUIL = NO encontrado "
                End If
                rs_aux.Close
                        
                '----------------------------------------------------------------
                ' Dirección del Empleado
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM detdom " & _
                         " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                         " LEFT JOIN zona ON zona.zonanro = detdom.zonanro " & _
                         " WHERE cabdom.ternro = " & rs_Tercero!Ternro & " AND " & _
                         " cabdom.tidonro =  " & Tidonro
                         
                OpenRecordset StrSql, rs_direccion
                If Not rs_direccion.EOF Then
                    
                    '----------------------------------------------------------------
                    ' 14 - Calle
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!Calle) Then
                        Calle = rs_direccion!Calle
                    Else
                        Calle = ""
                    End If
                    Flog.writeline " 14 -- Calle = " & Calle
                    
                    '----------------------------------------------------------------
                    ' 15 - Nro
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!Nro) Then
                        Nro = rs_direccion!Nro
                    Else
                        Nro = ""
                    End If
                    Flog.writeline " 15 -- Nro = " & Nro
                    
                    '----------------------------------------------------------------
                    ' 16 - Piso
                    '----------------------------------------------------------------
                                
                    If Not IsNull(rs_direccion!Piso) Then
                        Piso = rs_direccion!Piso
                    Else
                        Piso = ""
                    End If
                    Flog.writeline " 16 -- Piso = " & Piso
                    
                    '----------------------------------------------------------------
                    ' 17 - Departamento
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!oficdepto) Then
                        Depto = rs_direccion!oficdepto
                    Else
                        Depto = ""
                    End If
                    Flog.writeline " 17 -- Depto = " & Depto
                    
                    '----------------------------------------------------------------
                    ' 18 - Torre
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!Torre) Then
                        Torre = rs_direccion!Torre
                    Else
                        Torre = ""
                    End If
                    Flog.writeline " 18 -- Torre = " & Torre
                    
                    '----------------------------------------------------------------
                    ' 19 - Manzana
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!Manzana) Then
                        Manzana = rs_direccion!Manzana
                    Else
                        Manzana = ""
                    End If
                    Flog.writeline " 19 -- Manzana = " & Manzana
                    
                    '----------------------------------------------------------------
                    ' 20 - Codigo Postal
                    '----------------------------------------------------------------
                    If Not IsNull(rs_direccion!codigopostal) Then
                        CodPostal = rs_direccion!codigopostal
                    Else
                        CodPostal = ""
                    End If
                    Flog.writeline " 20 -- Cod. Postal = " & CodPostal
                    
                    '----------------------------------------------------------------
                    ' 21 - Entre Calles
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!entrecalles) Then
                        Entre = rs_direccion!entrecalles
                    Else
                        Entre = ""
                    End If
                    Flog.writeline " 21 -- Entre Calles = " & Entre
                    
                    '----------------------------------------------------------------
                    ' 22 - Barrio
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!Barrio) Then
                        Barrio = rs_direccion!Barrio
                    Else
                        Barrio = ""
                    End If
                    Flog.writeline " 22 -- Barrio = " & Barrio
                    
                    '----------------------------------------------------------------
                    ' 23 - Localidad
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!locnro) Then
                        StrSql = " SELECT locdesc FROM localidad " & _
                                " WHERE localidad.locnro= " & rs_direccion!locnro
                        OpenRecordset StrSql, rs_aux
                        If Not rs_aux.EOF Then
                            Loca = rs_aux!locdesc
                            Flog.writeline " 23 -- Localidad = " & Loca
                        Else
                            Loca = ""
                            Flog.writeline " 23 -- Localidad = NO encontrado "
                        End If
                        rs_aux.Close
                    Else
                        Loca = ""
                        Flog.writeline " 23 -- Localidad = NO encontrado "
                    End If
                    
                                        
                    '----------------------------------------------------------------
                    ' 24 - Provincia
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!provnro) Then
                        StrSql = " SELECT provdesc FROM provincia " & _
                                " WHERE provincia.provnro= " & rs_direccion!provnro
                        OpenRecordset StrSql, rs_aux
                        If Not rs_aux.EOF Then
                            Prov = rs_aux!provdesc
                            Flog.writeline " 24 -- Provincia = " & Prov
                        Else
                            Prov = ""
                            Flog.writeline " 24 -- Provincia = NO encontrado "
                        End If
                        rs_aux.Close
                    Else
                        Prov = ""
                        Flog.writeline " 24 -- Provincia = NO encontrado "
                    End If
                    
                    '----------------------------------------------------------------
                    ' 25 - Pais
                    '----------------------------------------------------------------
                    
                    If Not IsNull(rs_direccion!paisnro) Then
                        StrSql = " SELECT paisdesc FROM pais " & _
                                " WHERE pais.paisnro= " & rs_direccion!paisnro
                        OpenRecordset StrSql, rs_aux
                        If Not rs_aux.EOF Then
                            Pais = rs_aux!paisdesc
                            Flog.writeline " 25 -- Pais = " & Pais
                        Else
                            Pais = ""
                            Flog.writeline " 25 -- Pais = NO encontrado "
                        End If
                        rs_aux.Close
                    Else
                        Pais = ""
                        Flog.writeline " 25 -- Pais = NO encontrado "
                    End If
                    
                    '----------------------------------------------------------------
                    ' 26 - Telefono Particular
                    '----------------------------------------------------------------
                    
                    StrSql = " SELECT * FROM telefono " & _
                            " WHERE telefono.teldefault = -1 AND telefono.domnro= " & rs_direccion!domnro
                    OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                        Tele = rs_aux!telnro
                        Flog.writeline " 26 -- Telefono = " & Tele
                    Else
                        Tele = ""
                        Flog.writeline " 26 -- Telefono = NO encontrado "
                    End If
                    rs_aux.Close
        
                Else
                    Calle = ""
                    Nro = ""
                    Piso = ""
                    Depto = ""
                    Torre = ""
                    Manzana = ""
                    CodPostal = ""
                    Barrio = ""
                    Loca = ""
                    Prov = ""
                    Pais = ""
                    Tele = ""
                    
                    Flog.writeline " 15 -- Domicilio = NO encontrado "
                End If
                
                
                '----------------------------------------------------------------
                ' 27 - Sucursal
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 1 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Sucu = rs_Estructura!estrdabr
                    Flog.writeline " 27 -- Sucursal = " & Sucu
                Else
                    Sucu = ""
                    Flog.writeline " 27 -- Sucursal = NO encontrado "
                End If
                rs_Estructura.Close
        
                '----------------------------------------------------------------
                ' 28 - Convenio
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 19 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Conv = rs_Estructura!estrdabr
                    Flog.writeline " 28 -- Convenio = " & Conv
                Else
                    Conv = ""
                    Flog.writeline " 28 -- Convenio = NO encontrado "
                End If
                rs_Estructura.Close
        
                '----------------------------------------------------------------
                ' 29 - Categoria
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 3 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Cate = rs_Estructura!estrdabr
                    Flog.writeline " 29 -- Categoria = " & Cate
                Else
                    Cate = ""
                    Flog.writeline " 29 -- Categoria = NO encontrado "
                End If
                rs_Estructura.Close
                
                
                '----------------------------------------------------------------
                ' 30 - Puesto
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 4 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Puesto = rs_Estructura!estrdabr
                    Flog.writeline " 30 -- Puesto = " & Puesto
                Else
                    Puesto = ""
                    Flog.writeline " 30 -- Puesto = NO encontrado "
                End If
                rs_Estructura.Close
                
                
                '----------------------------------------------------------------
                ' 31 - Centro de Costo
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 5 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Centro = rs_Estructura!estrdabr
                    Flog.writeline " 31 -- Centro = " & Centro
                Else
                    Centro = ""
                    Flog.writeline " 31 -- Centro = NO encontrado "
                End If
                rs_Estructura.Close
                
                '----------------------------------------------------------------
                ' 32 - Caja de Jubilación (AFJP)
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 15 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Caja = rs_Estructura!estrdabr
                    Flog.writeline " 32 -- Caja de Jubilación = " & Caja
                Else
                    Sindi = ""
                    Flog.writeline " 32 -- Caja de Jubilación = NO encontrado "
                End If
                rs_Estructura.Close
                
                
                '----------------------------------------------------------------
                ' 33 - Sindicato
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 16 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Sindi = rs_Estructura!estrdabr
                    Flog.writeline " 33 -- Sindicato = " & Sindi
                Else
                    Sindi = ""
                    Flog.writeline " 33 -- Sindicato = NO encontrado "
                End If
                rs_Estructura.Close
                
                '----------------------------------------------------------------
                ' 34 - Obra Social elegida
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 17 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Osocial = rs_Estructura!estrdabr
                    Flog.writeline " 34 -- Obra Social = " & Osocial
                Else
                    Osocial = ""
                    Flog.writeline " 34 -- Obra Social = NO encontrado "
                End If
                rs_Estructura.Close
                
                
                '----------------------------------------------------------------
                ' 35 - PLan de Obra Social elegida
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 23 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    PlanOsocial = rs_Estructura!estrdabr
                    Flog.writeline " 35 -- Plan de Obra Social = " & PlanOsocial
                Else
                    PlanOsocial = ""
                    Flog.writeline " 35 -- PLan de Obra Social = NO encontrado "
                End If
                rs_Estructura.Close
                
                '----------------------------------------------------------------
                ' 36 - Estado Empleado
                '----------------------------------------------------------------
                If rs_emple!empest = -1 Then
                    EstEmp = "Activo"
                Else
                    EstEmp = "Inactivo"
                End If
                Flog.writeline " 36 -- Estado Empleado = " & EstEmp
                
                '----------------------------------------------------------------
                ' 37, 38 - Causa y Fecha de Baja
                '----------------------------------------------------------------
                If EstEmp = "Inactivo" Then
                
                    StrSql = " SELECT * FROM fases "
                    StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro "
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " fases.empleado = " & rs_Tercero!Ternro
                    StrSql = StrSql & " AND fases.fasnro = (select max(fasnro) from fases where empleado = " & rs_Tercero!Ternro & ")"
                    StrSql = StrSql & " ORDER BY fases.altfec desc "
                    OpenRecordset StrSql, rs_Estructura
                    If Not rs_Estructura.EOF Then
                        Causa = rs_Estructura!caudes
                        FecBaj = rs_Estructura!bajfec
                        Flog.writeline " 37, 38 -- Causa y Fecha de Baja = " & Causa & " " & FecBaj
                    Else
                        Causa = ""
                        FecBaj = ""
                        Flog.writeline " 37, 38 -- Causa y Fecha de Baja = NO encontrado "
                    End If
                    rs_Estructura.Close 'NG - Cerre conexion
                Else
                    Causa = ""
                    FecBaj = ""
                End If
                
                
                '----------------------------------------------------------------
                ' 39 - Empresa
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 10 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Empre = rs_Estructura!estrdabr
                    Flog.writeline " 39 -- Empresa = " & Empre
                Else
                    Empre = ""
                    Flog.writeline " 39 -- Empresa = NO encontrado "
                End If
                rs_Estructura.Close
                
                '----------------------------------------------------------------
                ' 40 - Remuneración
                '----------------------------------------------------------------
                
                If Not IsNull(rs_emple!empremu) Then
                    Remu = rs_emple!empremu
                Else
                    Remu = ""
                End If
                
                Flog.writeline " 40 -- Remuneración = " & Remu
                
                '----------------------------------------------------------------
                ' 41 - Modelo de Organización
                '----------------------------------------------------------------
                
                If Not IsNull(rs_emple!tplatenro) Then
                    StrSql = " SELECT tplatedesabr FROM adptemplate " & _
                            " WHERE adptemplate.tplatenro= " & rs_emple!tplatenro
                    OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                        Orga = rs_aux!tplatedesabr
                        Flog.writeline " 42 -- Organización = " & Orga
                    Else
                        Orga = ""
                        Flog.writeline " 42 -- Organización = NO encontrado "
                    End If
                    rs_aux.Close
                Else
                    Orga = ""
                    Flog.writeline " 42 -- Organización = NO encontrado "
                End If
            
                '----------------------------------------------------------------
                ' 42 - Gerencia
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 6 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Gerencia = rs_Estructura!estrdabr
                    Flog.writeline " 43 -- Gerencia = " & Gerencia
                Else
                    Gerencia = ""
                    Flog.writeline " 43 -- Gerencia = NO encontrado "
                End If
                rs_Estructura.Close
                
                
                '----------------------------------------------------------------
                ' 43 - Contrato
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 18 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Contrato = rs_Estructura!estrdabr
                    Flog.writeline " 44 -- Contrato = " & Contrato
                Else
                    Contrato = ""
                    Flog.writeline " 44 -- Contrato = NO encontrado "
                End If
                rs_Estructura.Close
                
                '----------------------------------------------------------------
                ' 44 - Grupo de Seguridad
                '----------------------------------------------------------------
        
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Tercero!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 7 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_estruc) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(fecha_estruc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    GrupSeg = rs_Estructura!estrdabr
                    Flog.writeline " 45 -- Grupo Seguridad = " & GrupSeg
                Else
                    GrupSeg = ""
                    Flog.writeline " 45 -- Grupo Seguridad = NO encontrado "
                End If
                rs_Estructura.Close
                
                '----------------------------------------------------------------
                ' Exporto cada Empleado
                '----------------------------------------------------------------
               
                Linea = Legajo & Sep
                Linea = Linea & Apellido & Sep
                Linea = Linea & nombre & Sep
                Linea = Linea & FecNac & Sep
                Linea = Linea & PaisNac & Sep
                Linea = Linea & FecIng & Sep
                
                Linea = Linea & Estciv & Sep
                Linea = Linea & Sexo & Sep
                Linea = Linea & FecAlt & Sep
                Linea = Linea & Nivest & Sep
                Linea = Linea & TipoDocu & Sep
                Linea = Linea & NroDocu & Sep
                
                Linea = Linea & Cuil & Sep
                Linea = Linea & Calle & Sep
                Linea = Linea & Nro & Sep
                Linea = Linea & Piso & Sep
                Linea = Linea & Depto & Sep
                
                Linea = Linea & Torre & Sep
                Linea = Linea & Manzana & Sep
                Linea = Linea & CodPostal & Sep
                Linea = Linea & Entre & Sep
                Linea = Linea & Barrio & Sep
                
                Linea = Linea & Loca & Sep
                Linea = Linea & Prov & Sep
                Linea = Linea & Pais & Sep
                Linea = Linea & Tele & Sep
                Linea = Linea & Sucu & Sep
        
                Linea = Linea & Conv & Sep
                Linea = Linea & Cate & Sep
                Linea = Linea & Puesto & Sep
                Linea = Linea & Centro & Sep
                Linea = Linea & Caja & Sep
                Linea = Linea & Sindi & Sep
                
                Linea = Linea & Osocial & Sep
                Linea = Linea & PlanOsocial & Sep
                Linea = Linea & EstEmp & Sep
                Linea = Linea & Causa & Sep
                Linea = Linea & FecBaj & Sep
                Linea = Linea & Empre & Sep
                Linea = Linea & Remu & Sep
                Linea = Linea & Orga & Sep
                Linea = Linea & Gerencia & Sep
                Linea = Linea & Contrato & Sep
                Linea = Linea & GrupSeg & ""
                       
                ArchExp.Write Linea
                ArchExp.writeline ""
         
     Else
        Flog.writeline "No se encontró el tercero"
     End If
    
       'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
              
        cantRegistros = cantRegistros - 1
           
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & bpronro
        objConn.Execute StrSql, , adExecuteNoRecords
    
    rs_emple.MoveNext
Loop

'----------------------------------------------------------------
' Borrar los Empleados de la tabla batch_proceso
'----------------------------------------------------------------

StrSql = " DELETE FROM batch_empleado "
StrSql = StrSql & " WHERE bpronro = " & bpronro
objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub

