Attribute VB_Name = "MdlIntTiger"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "14/05/2007" 'Martin Ferraro - Version Inicial

'Const Version = 1.02
'Const FechaVersion = "14/06/2007"
'Martin Ferraro -   Permitir Fecha nula ""
'                   Calificacion profesional 49

'Const Version = 1.03
'Const FechaVersion = "15/06/2007"
'Martin Ferraro - Licencia Vacaciones Buscaba tipo 20, se corrigio por dos
'                 Cambio la busqueda de fecha de baja

'Const Version = 1.04
'Const FechaVersion = "30/11/2007"
''Martin Ferraro - En las estructuras 29, 30 y 56 se busca el tipo de codigo

Global Const Version = "1.05" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

'----------------------------------------------------------------

Global ValorNum_Confrep(100) As String
Global ValorAlf_Confrep(100) As String
Global ArrPar
Global Formato_Fecha As String
Global Fecha_Nula As String
Global Fecha_Vacia As String
'-----------------------------------------------
'Variables de Planificacion
Global IdUser As String
Global Schednro As Long
Global dia As String
Global Hora As String
Global Nuevo_bprcparam As String
Global RealizarPlanificacion As Boolean
'-----------------------------------------------



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Exportacion Tiger.
' Autor      : Martin Ferraro
' Fecha      : 14/05/2007
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
    
    Nombre_Arch = PathFLog & "Interf_Tiger" & "-" & NroProcesoBatch & ".log"
    
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
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 174 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    RealizarPlanificacion = False
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ExportTiger(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        '---------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------
        If RealizarPlanificacion Then
            Flog.writeline
            Flog.writeline
            Flog.writeline "Alarma recurrente. Setear estado del proceso = 'Pendiente'."
                            
            'Control si ya existia un proceso planificado
            StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 174 "
            StrSql = StrSql & " AND bprcfecha =" & ConvFecha(dia)
            StrSql = StrSql & " AND bprchora ='" & Hora & "'"
            StrSql = StrSql & " AND bprcestado ='Pendiente'"
            StrSql = StrSql & " ORDER BY bprcparam"
            OpenRecordset StrSql, rs_batch_proceso
            If rs_batch_proceso.EOF Then
                'No existia, lo creo
                StrSql = "INSERT INTO batch_proceso"
                StrSql = StrSql & " (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
                StrSql = StrSql & " bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados,bprcurgente) "
                StrSql = StrSql & " VALUES (174," & ConvFecha(dia) & ", '" & IdUser & "','" & Hora & "' "
                StrSql = StrSql & " ,null,null"
                StrSql = StrSql & " , '" & Nuevo_bprcparam & "', 'Pendiente', null , null, null, null, 0, null,0)"
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        '---------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------
    End If
    
    
Fin:
    Flog.Close
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


Public Sub ExportTiger(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de Exportacion Tiger
' Autor      : Martin Ferraro
' Fecha      : 14/05/2007
' --------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim ternro As Long
Dim tipo As Integer
Dim FecDesde As Date
Dim FecHasta As Date
Dim Planificar As Boolean

'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------
Dim Directorio As String
Dim Separador As String
Dim SeparadorDecimal As String
Dim DescripcionModelo As String
Dim Archivo As String
Dim fExport
Dim carpeta
Dim Aux_Linea As String
Dim EstructuraNro As Long
Dim EstructuraDesc As String
Dim Legajo As Long
Dim terape As String
Dim ternom As String
Dim NacionacNro As Long
Dim NacionacDesabr As String
Dim FechaNac As String
Dim FechaIng As String
Dim Sexo As String
Dim EstCivilNro As Long
Dim EstCivilDesc As String
Dim I As Long
Dim TidNro As String
Dim TidNom As String
Dim DocNro As String
Dim InstNro As String
Dim InstDes As String
Dim ValorNov As Double
Dim Cuil As String
Dim CI As String
Dim Calle As String
Dim NroCalle As String
Dim Piso As String
Dim Oficdepto As String
Dim Codigopostal As String
Dim ProvinciaNro As String
Dim ProvinciaDesc As String
Dim LocalidadNro As String
Dim LocalidadDesc As String
Dim Telefono As String
Dim NivEstNro As Long
Dim NivEstDesc As String
Dim CausaNro As Long
Dim CausaDesc As String
Dim FechaBaja As String
Dim FechaIngCat As String
Dim FechaIngRec As String
Dim FechaUltIng As String
Dim FechaIngCont As String
Dim ImpTopeOS As Double
Dim CodContr As String
Dim CodModContr As String
Dim CBU As String
Dim CtaNro As String
Dim FormaPagoNro As String
Dim FormaPagoDesc As String
Dim BancoNro As String
Dim BancoDesc As String
Dim DesdeLicVac As String
Dim DesdeLicMaternidad As String
Dim HastaLicMaternidad As String

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset

'Inicio codigo ejecutable
On Error GoTo E_ExportTiger

'-------------------------------------------------------------------------------------------------
'Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, "@")
        If UBound(ArrPar) = 4 Then
            ternro = IIf(ArrPar(0) = "", 0, ArrPar(0))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Tercero = " & ternro
            
            tipo = CInt(ArrPar(1))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Tipo = " & tipo
            
            FecDesde = CDate(ArrPar(2))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Fecha Desde = " & FecDesde
        
            FecHasta = CDate(ArrPar(3))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Fecha Hasta = " & FecHasta
            
            Planificar = ArrPar(4)
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Planificar = " & Planificar
            
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Numero de parametros erroneo."
            Exit Sub
        End If
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    Exit Sub
End If
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte."

'Inicializo el array de configuracion
For I = 1 To UBound(ValorNum_Confrep)
    ValorNum_Confrep(I) = "0"
    ValorAlf_Confrep(I) = ""
Next I

StrSql = "SELECT * FROM confrep WHERE repnro = 199 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte."
Else
    Do While Not rs_Consult.EOF
        If UBound(ValorNum_Confrep) >= rs_Consult!confnrocol Then
            ValorNum_Confrep(rs_Consult!confnrocol) = IIf(EsNulo(rs_Consult!confval), "0", rs_Consult!confval)
            ValorAlf_Confrep(rs_Consult!confnrocol) = IIf(EsNulo(rs_Consult!confval2), "", rs_Consult!confval2)
            Flog.writeline Espacios(Tabulador * 1) & "Columna " & rs_Consult!confnrocol & " " & rs_Consult!confetiq & " VNUM = " & ValorNum_Confrep(rs_Consult!confnrocol) & " VALF = " & ValorAlf_Confrep(rs_Consult!confnrocol)
        End If
    
        rs_Consult.MoveNext
    Loop
End If
rs_Consult.Close
Flog.writeline

'Validaciones
If EsNulo(ValorAlf_Confrep(1)) Then
    ValorAlf_Confrep(1) = "dd/mm/yyyy"
End If
If EsNulo(ValorAlf_Confrep(2)) Then
    '14/06/2007 - Martin Ferraro - Permitir Fecha_Nula ""
    'ValorAlf_Confrep(2) = "31/12/9999"
    ValorAlf_Confrep(2) = ""
End If

Formato_Fecha = ValorAlf_Confrep(1)
Fecha_Nula = ValorAlf_Confrep(2)
Schednro = ValorNum_Confrep(3)

'-------------------------------------------------------------------------------------------------
'Configuracion del Directorio de salida
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio de salida."
StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Trim(rs_Consult!sis_direntradas)
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el registro de la tabla sistema nro 1"
    Exit Sub
End If
rs_Consult.Close
Flog.writeline
    
    
'-------------------------------------------------------------------------------------------------
'Configuracion del Modelo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Modelo Interface."
StrSql = "SELECT * FROM modelo WHERE modnro = 908"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Directorio & Trim(rs_Consult!modarchdefault)
    Separador = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, ",")
    SeparadorDecimal = IIf(Not IsNull(rs_Consult!modsepdec), rs_Consult!modsepdec, ".")
    DescripcionModelo = rs_Consult!moddesc
    
    Flog.writeline Espacios(Tabulador * 1) & "Modelo 908 " & rs_Consult!moddesc
    Flog.writeline Espacios(Tabulador * 1) & "Directorio de Exportacion : " & Directorio
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo 908."
    Exit Sub
End If
rs_Consult.Close
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Planificacion
'-------------------------------------------------------------------------------------------------
If Planificar Then

    Flog.writeline Espacios(Tabulador * 0) & "Calculando Planificacion."
    Flog.writeline
    If Schednro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. Falta configurar el numero de planificador en la columna 3 del confrep 199."
    Else
        RealizarPlanificacion = True
        Call FechaHora(Schednro)
        
        'La planificacion es desde la fecha calculada (dia) a la fecha hasta anterior (FecHasta del filtro)
        If CLng(DateDiff("d", FecHasta, dia)) < 0 Then
            Nuevo_bprcparam = "@3@" & dia & "@" & dia & "@-1"
        Else
            Nuevo_bprcparam = "@3@" & FecHasta & "@" & dia & "@-1"
        End If
        
    End If
    
End If



'-------------------------------------------------------------------------------------------------
'Busqueda de empleados
'-------------------------------------------------------------------------------------------------
If tipo <> 3 Then
    'Si no estan marcados todos los empleados, los seleccionados fueron insertados en batch_empleado
    StrSql = "SELECT empleado.ternro, empleado.empleg, empleado.ternom, empleado.ternom2, empleado.terape, empleado.terape2, batch_empleado.progreso, empleado.empfaltagr"
    StrSql = StrSql & " FROM batch_empleado"
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro"
    StrSql = StrSql & " AND batch_empleado.bpronro = " & bpronro
    StrSql = StrSql & " ORDER BY batch_empleado.progreso"
Else
    'Buscando todos los empleados
    StrSql = " SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2, empleado.empfaltagr"
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " ORDER BY empleg"
End If

OpenRecordset StrSql, rs_Empleados

'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay empleados"
    CEmpleadosAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Empleados: " & CEmpleadosAProc
End If
IncPorc = (100 / CEmpleadosAProc)
        
        
'-------------------------------------------------------------------------------------------------
'Creacion del archivo
'-------------------------------------------------------------------------------------------------
If Not rs_Empleados.EOF Then
    'Seteo el nombre del archivo generado
    Archivo = Directorio & "\Exp_tiger_" & Format(Date, "dd-mm-yyyy") & "_" & bpronro & ".txt"
    Set fs = CreateObject("Scripting.FileSystemObject")
    'Activo el manejador de errores
    On Error Resume Next
    Set fExport = fs.CreateTextFile(Archivo, True)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs.CreateFolder(Directorio)
        Set fExport = fs.CreateTextFile(Archivo, True)
    End If
    On Error GoTo E_ExportTiger
    Flog.writeline Espacios(Tabulador * 0) & "Archivo Creado: " & "Exp_tiger_" & Date & "_" & bpronro & ".txt"
End If
  
        
'-------------------------------------------------------------------------------------------------
'Comienzo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el Procesamiento de Empleados"
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Do While Not rs_Empleados.EOF
    
    ternro = rs_Empleados!ternro
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO EMPLEADO: " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!ternom
    
    
    Legajo = rs_Empleados!empleg
    terape = IIf(EsNulo(rs_Empleados!terape), "", rs_Empleados!terape)
    terape = terape & IIf(EsNulo(rs_Empleados!terape2), "", " " & rs_Empleados!terape2)
    ternom = IIf(EsNulo(rs_Empleados!ternom), "", rs_Empleados!ternom)
    ternom = ternom & IIf(EsNulo(rs_Empleados!ternom2), "", " " & rs_Empleados!ternom2)
    FechaIng = Fecha_Nula
    If Not EsNulo(rs_Empleados!empfaltagr) Then FechaIng = Format_Fecha(rs_Empleados!empfaltagr, Formato_Fecha, Fecha_Nula)

    'Buscado Datos del tercero
    StrSql = "SELECT nacionalidad.nacionaldes, nacionalidad.nacionalnro, tercero.terfecing"
    StrSql = StrSql & " , tercero.tersex, tercero.terfecnac"
    StrSql = StrSql & " , estcivil.estcivnro, estcivil.estcivdesabr"
    StrSql = StrSql & " , ter_doc.nrodoc, ci.nrodoc cedula"
    StrSql = StrSql & " FROM tercero "
    StrSql = StrSql & " LEFT JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro"
    StrSql = StrSql & " LEFT JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro"
    StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = tercero.ternro"
    StrSql = StrSql & " AND ter_doc.tidnro = 10 "
    StrSql = StrSql & " LEFT JOIN ter_doc ci ON ci.ternro = tercero.ternro"
    StrSql = StrSql & " AND ci.tidnro = 4 "
    StrSql = StrSql & " WHERE tercero.ternro = " & ternro
    OpenRecordset StrSql, rs_Consult
    'Valores Iniciales
    NacionacNro = 1
    NacionacDesabr = "Argentina"
    FechaNac = Fecha_Nula
    Sexo = ""
    EstCivilNro = 0
    EstCivilDesc = ""
    Cuil = "0"
    CI = "0"
    If Not rs_Consult.EOF Then
        NacionacNro = IIf(EsNulo(rs_Consult!nacionalnro), 1, rs_Consult!nacionalnro)
        NacionacDesabr = IIf(EsNulo(rs_Consult!nacionaldes), "Argentina", rs_Consult!nacionaldes)
        If Not EsNulo(rs_Consult!terfecnac) Then FechaNac = Format_Fecha(rs_Consult!terfecnac, Formato_Fecha, Fecha_Nula)
        Sexo = IIf(rs_Consult!tersex = -1, "V", "M")
        EstCivilNro = IIf(EsNulo(rs_Consult!estcivnro), 0, rs_Consult!estcivnro)
        EstCivilDesc = IIf(EsNulo(rs_Consult!estcivdesabr), "", rs_Consult!estcivdesabr)
        If Not EsNulo(rs_Consult!nrodoc) Then Cuil = Replace(rs_Consult!nrodoc, "-", "")
        If Not EsNulo(rs_Consult!cedula) Then CI = Replace(rs_Consult!cedula, "-", "")
    End If
    rs_Consult.Close
    
    'Empresa
    Call BuscarEstructura(10, ternro, FecHasta, 1, "Sofrecom Argentina S.A.", EstructuraNro, EstructuraDesc)
    Aux_Linea = EstructuraNro & Separador & EstructuraDesc
    
    'Legajo
    Aux_Linea = Aux_Linea & Separador & Legajo
    
    'Tipo de personal
    Call BuscarEstructura(22, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Filial
    Call BuscarEstructura(1, ternro, FecHasta, 1, "Adm. Central", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Apellido y Nombres
    Aux_Linea = Aux_Linea & Separador & terape & ", " & ternom
    
    'Nacionalidad
    Aux_Linea = Aux_Linea & Separador & NacionacNro & Separador & NacionacDesabr
    
    'Fecha de Nacimiento
    Aux_Linea = Aux_Linea & Separador & FechaNac
    
    'Datos del documento
    Call BuscarDoc(ternro, "1", "DNI", "0", "1", "RNP", TidNro, TidNom, DocNro, InstNro, InstDes)
    Aux_Linea = Aux_Linea & Separador & TidNro & Separador & TidNom
    Aux_Linea = Aux_Linea & Separador & InstNro & Separador & InstDes
    Aux_Linea = Aux_Linea & Separador & DocNro
    
    'Sexo
    Aux_Linea = Aux_Linea & Separador & Sexo
    
    'Estado Civil
    Aux_Linea = Aux_Linea & Separador & EstCivilNro & Separador & EstCivilDesc
         
    'Contrato
    Call BuscarEstructura(18, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Nivel de estudio
    NivEstNro = 0
    NivEstDesc = ""
    StrSql = " SELECT nivest.nivnro, nivcodext, nivest.nivdesc, cap_estformal.capfechas "
    StrSql = StrSql & " FROM cap_estformal "
    StrSql = StrSql & " INNER JOIN nivest ON cap_estformal.nivnro = nivest.nivnro "
    StrSql = StrSql & " WHERE cap_estformal.ternro = " & ternro
    'StrSql = StrSql & " AND cap_estformal.capcomp = -1 "
    'StrSql = StrSql & " AND cap_estformal.capfechas is not null "
    StrSql = StrSql & " AND cap_estformal.capfechas <= " & ConvFecha(FecHasta)
    StrSql = StrSql & " ORDER BY cap_estformal.capfechas DESC "
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        NivEstNro = IIf(EsNulo(rs_Consult!nivnro), 0, rs_Consult!nivnro)
        NivEstDesc = IIf(EsNulo(rs_Consult!nivdesc), "", rs_Consult!nivdesc)
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & NivEstNro & Separador & NivEstDesc
    
    'Datos del Domicilio
    Calle = ""
    NroCalle = "0"
    Piso = "0"
    Oficdepto = ""
    Codigopostal = ""
    ProvinciaNro = "0"
    ProvinciaDesc = ""
    LocalidadNro = "0"
    LocalidadDesc = ""
    Telefono = ""
    StrSql = " SELECT detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal"
    StrSql = StrSql & " , detdom.provnro, detdom.locnro, provincia.provdesc, localidad.locdesc, telefono.telnro"
    StrSql = StrSql & " From cabdom"
    StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
    StrSql = StrSql & " LEFT JOIN localidad ON localidad.locnro = detdom.locnro"
    StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro"
    StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = cabdom.domnro"
    StrSql = StrSql & " AND telefono.teldefault = -1"
    StrSql = StrSql & " Where cabdom.domdefault = -1"
    StrSql = StrSql & " AND cabdom.ternro = " & ternro
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!Calle) Then Calle = rs_Consult!Calle
        If Not EsNulo(rs_Consult!nro) Then NroCalle = rs_Consult!nro
        If Not EsNulo(rs_Consult!Piso) Then Piso = rs_Consult!Piso
        If Not EsNulo(rs_Consult!Oficdepto) Then Oficdepto = rs_Consult!Oficdepto
        If Not EsNulo(rs_Consult!Codigopostal) Then Codigopostal = rs_Consult!Codigopostal
        If Not EsNulo(rs_Consult!provnro) Then ProvinciaNro = rs_Consult!provnro
        If Not EsNulo(rs_Consult!provdesc) Then ProvinciaDesc = rs_Consult!provdesc
        If Not EsNulo(rs_Consult!locnro) Then LocalidadNro = rs_Consult!locnro
        If Not EsNulo(rs_Consult!locdesc) Then LocalidadDesc = rs_Consult!locdesc
        If Not EsNulo(rs_Consult!telnro) Then Telefono = rs_Consult!telnro
    End If
    rs_Consult.Close
    
    'Domicilio
    Aux_Linea = Aux_Linea & Separador & Calle & Separador & NroCalle
    Aux_Linea = Aux_Linea & Separador & Piso & Separador & Oficdepto
    
    'Localidad
    Aux_Linea = Aux_Linea & Separador & LocalidadNro & Separador & LocalidadDesc
    
    'CP
    Aux_Linea = Aux_Linea & Separador & Codigopostal
    
    'Provincia
    Aux_Linea = Aux_Linea & Separador & ProvinciaNro & Separador & ProvinciaDesc

    'Telefono
    Aux_Linea = Aux_Linea & Separador & Telefono

    'Actividad SIJP
    '30/11/2007 - Martin Ferraro - Buscar tipo de codigo SIJP de la estructura
    'Call BuscarEstructura(29, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Call BuscarEstructuraTipoCod(29, 1, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Fecha de alta en el sistema - Fecha de ingresos
    Aux_Linea = Aux_Linea & Separador & FechaIng
    
    'Fecha de baja - Busco fase cuya fecha de baja sea menor al hasta y no exista otra alta en el rango
    CausaNro = 0
    CausaDesc = ""
    FechaBaja = Fecha_Nula
    StrSql = "SELECT fases.bajfec, causa.caunro, causa.caudes"
    StrSql = StrSql & " FROM fases"
    StrSql = StrSql & " LEFT JOIN causa ON causa.caunro = fases.caunro"
    StrSql = StrSql & " WHERE fases.Empleado = " & ternro
    StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(FecHasta)
    StrSql = StrSql & " ORDER BY fases.altfec Desc"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        
        'Busco una fase activa mayor a la encontrada dentro del rango
        StrSql = "SELECT fases.fasnro"
        StrSql = StrSql & " From fases"
        StrSql = StrSql & " Where fases.Empleado = " & ternro
        StrSql = StrSql & " AND fases.altfec > " & ConvFecha(rs_Consult!bajfec)
        StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(FecHasta)
        OpenRecordset StrSql, rs_Aux
        If rs_Aux.EOF Then
            If Not EsNulo(rs_Consult!caunro) Then CausaNro = rs_Consult!caunro
            If Not EsNulo(rs_Consult!caudes) Then CausaDesc = rs_Consult!caudes
            If Not EsNulo(rs_Consult!bajfec) Then FechaBaja = Format_Fecha(rs_Consult!bajfec, Formato_Fecha, Fecha_Nula)
        End If
        rs_Aux.Close
        
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & FechaBaja
    
    'Fecha Ingreso a Categoria
    FechaIngCat = Fecha_Nula
    StrSql = " SELECT his_estructura.htetdesde"
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 3 AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(FecHasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(FecHasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!htetdesde) Then FechaIngCat = Format_Fecha(rs_Consult!htetdesde, Formato_Fecha, Fecha_Nula)
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & FechaIngCat
    
    'Cod Talle 3
    Call BuscarEstructura(45, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Fecha de ingreso reconocida - Fecha de fase marcada como reconocida
    FechaIngRec = Fecha_Nula
    StrSql = "SELECT fases.altfec"
    StrSql = StrSql & " FROM fases"
    StrSql = StrSql & " WHERE fases.Empleado = " & ternro
    StrSql = StrSql & " AND fases.fasrecofec = -1"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!altfec) Then FechaIngRec = Format_Fecha(rs_Consult!altfec, Formato_Fecha, Fecha_Nula)
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & FechaIngRec
    
    'Fecha Ultimo Ingreso menor al alta
    FechaUltIng = Fecha_Nula
    StrSql = "SELECT fases.altfec"
    StrSql = StrSql & " FROM fases"
    StrSql = StrSql & " WHERE fases.Empleado = " & ternro
    'StrSql = StrSql & " AND fases.real = -1 "
    'StrSql = StrSql & " AND (" & ConvFecha(FecDesde) & " <= fases.altfec"
    StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(FecHasta)
    StrSql = StrSql & " ORDER BY fases.altfec Desc"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!altfec) Then FechaUltIng = Format_Fecha(rs_Consult!altfec, Formato_Fecha, Fecha_Nula)
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & FechaUltIng
    
    'Fecha Egreso Transitorio
    Aux_Linea = Aux_Linea & Separador & FechaBaja
    
    'Causa de Egreso
    Aux_Linea = Aux_Linea & Separador & CausaNro & Separador & CausaDesc
    
    'Cod de sector de trabajo
    Call BuscarEstructura(2, ternro, FecHasta, 471, "Division Informatica", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cod de Jurisdiccion
    Call BuscarEstructura(48, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Personal en convenio
    Call BuscarEstructura(32, ternro, FecHasta, 402, "Fuera de Convenio", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cod Centro de Costos 1
    Aux_Linea = Aux_Linea & Separador & "0"
    
    '% en Centro de Costos 1
    Aux_Linea = Aux_Linea & Separador & "100"
    
    'Cuenta Bancaria
    CBU = ""
    CtaNro = "0"
    FormaPagoNro = "0"
    FormaPagoDesc = ""
    BancoNro = "0"
    BancoDesc = ""
    StrSql = " SELECT ctabancaria.ctabnro, ctabancaria.ctabcbu,"
    StrSql = StrSql & " formapago.fpagdescabr, formapago.fpagnro,"
    StrSql = StrSql & " banc.estrnro, banc.estrdabr"
    StrSql = StrSql & " FROM ctabancaria"
    StrSql = StrSql & " INNER JOIN formapago ON ctabancaria.fpagnro = formapago.fpagnro"
    StrSql = StrSql & " AND formapago.fpagbanc = -1"
    StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
    StrSql = StrSql & " INNER JOIN estructura banc ON banc.estrnro = banco.estrnro "
    StrSql = StrSql & " WHERE ctabancaria.ternro  = " & ternro
    StrSql = StrSql & " AND ctabancaria.ctabestado = -1"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!ctabcbu) Then CBU = rs_Consult!ctabcbu
        If Not EsNulo(rs_Consult!ctabnro) Then CtaNro = rs_Consult!ctabnro
        If Not EsNulo(rs_Consult!fpagnro) Then FormaPagoNro = rs_Consult!fpagnro
        If Not EsNulo(rs_Consult!fpagdescabr) Then FormaPagoDesc = rs_Consult!fpagdescabr
        If Not EsNulo(rs_Consult!Estrnro) Then BancoNro = rs_Consult!Estrnro
        If Not EsNulo(rs_Consult!Estrdabr) Then BancoDesc = rs_Consult!Estrdabr
    End If
    rs_Consult.Close
    
    'Forma de Pago
     Aux_Linea = Aux_Linea & Separador & FormaPagoNro & Separador & FormaPagoDesc
    
    'Banco y sucursal de pago
     Aux_Linea = Aux_Linea & Separador & BancoNro & Separador & BancoDesc
    
    'Nro de cuenta
    Aux_Linea = Aux_Linea & Separador & CtaNro
    
    'Cod de Convenio
    Call BuscarEstructura(19, ternro, FecHasta, 347, "Fuera de Convenio", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cod de Categoria
    Call BuscarEstructura(3, ternro, FecHasta, 534, "Empleado", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Calificacion Profesional
    '14/06/2007 - Martin Ferraro - Cambio tipo de estructura 48 por 49
    'Call BuscarEstructura(48, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Call BuscarEstructura(49, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Tipo de Remuneracion
    Aux_Linea = Aux_Linea & Separador & "2"
    
    'Cod de no liq autom
    Aux_Linea = Aux_Linea & Separador & "0"
    
    'Cod tipo ropa de trabajo
    Call BuscarEstructura(46, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Imp Basico de convenio
    Aux_Linea = Aux_Linea & Separador & "0"
    
    'Busqueda de las 19 primeras novedades
    For I = 10 To 19
        If Len(ValorAlf_Confrep(I)) <> 0 Then
            Call BuscarNovedad(ValorAlf_Confrep(I), ValorNum_Confrep(I), ternro, FecDesde, FecHasta, 0, ValorNov)
            Aux_Linea = Aux_Linea & Separador & Replace(CStr(ValorNov), ".", SeparadorDecimal)
        Else
            Aux_Linea = Aux_Linea & Separador & "0"
        End If
    Next I
    
    'Cod de sindicato
    Call BuscarEstructura(16, ternro, FecHasta, 549, "Fuera de Sindicato", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Marca de afiliado a sindicato
    Aux_Linea = Aux_Linea & Separador & "0"
    
    'Cuil
    Aux_Linea = Aux_Linea & Separador & Cuil
    
    'Cod de caja de jubilacion
    Call BuscarEstructura(15, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Nro de Jubilacion
    Aux_Linea = Aux_Linea & Separador & "0"
    
    'Cod de obra social
    Call BuscarEstructura(17, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Marca seguro de vida titular
    '30/11/2007 - Martin Ferraro - Buscar tipo de codigo 38 de la estructura
    'Call BuscarEstructura(56, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Call BuscarEstructuraTipoCod(56, 38, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Imp Tope Obra Social
    ImpTopeOS = 0
    StrSql = " SELECT plprecio "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " INNER JOIN replica_estr ON estructura.estrnro=replica_estr.estrnro "
    StrSql = StrSql & " INNER JOIN planos ON replica_estr.origen = planos.plnro"
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 23 AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(FecHasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(FecHasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!plprecio) Then ImpTopeOS = rs_Consult!plprecio
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & Replace(CStr(ImpTopeOS), ".", SeparadorDecimal)
    
    'Servicio Fijo 1
    Call BuscarEstructura(15, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cod plan OS
    Call BuscarEstructura(23, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Marca liq. imp. gan.
    Aux_Linea = Aux_Linea & Separador & "1"
    
    'Tipo de cuenta bancaria
    Call BuscarEstructura(55, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Nivel para sal fam
    Aux_Linea = Aux_Linea & Separador & "3"
    
    'Busco la fecha de inicio de licencia de vacaciones que se encuentre en el rango
    DesdeLicVac = Fecha_Nula
    StrSql = " SELECT elfechadesde "
    StrSql = StrSql & " FROM emp_lic"
    StrSql = StrSql & " WHERE Empleado = " & ternro
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " ( ( elfechadesde <= " & ConvFecha(FecDesde) & " ) AND ( elfechahasta >= " & ConvFecha(FecDesde) & " ) )"
    StrSql = StrSql & " OR"
    StrSql = StrSql & " ( ( elfechadesde >= " & ConvFecha(FecDesde) & " ) AND ( elfechadesde <= " & ConvFecha(FecHasta) & " ) )"
    StrSql = StrSql & " )"
    StrSql = StrSql & " AND tdnro = 2"
    StrSql = StrSql & " AND licestnro = 2"
    StrSql = StrSql & " ORDER BY elfechadesde"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!elfechadesde) Then DesdeLicVac = Format_Fecha(rs_Consult!elfechadesde, Formato_Fecha, Fecha_Nula)
    End If
    rs_Consult.Close
    'Fecha Inicio Vacaciones
    Aux_Linea = Aux_Linea & Separador & DesdeLicVac
    
    'Cod Tarea efectiva
    Call BuscarEstructura(4, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cedula de identidad
    Aux_Linea = Aux_Linea & Separador & CI
    
    'Busqueda de las novedades
    For I = 20 To 23
        If Len(ValorAlf_Confrep(I)) <> 0 Then
            Call BuscarNovedad(ValorAlf_Confrep(I), ValorNum_Confrep(I), ternro, FecDesde, FecHasta, 0, ValorNov)
            Aux_Linea = Aux_Linea & Separador & Replace(CStr(ValorNov), ".", SeparadorDecimal)
        Else
            Aux_Linea = Aux_Linea & Separador & "0"
        End If
    Next I
    
    'Fecha Inicio Contrato
    FechaIngCont = Fecha_Nula
    CodContr = ""
    StrSql = " SELECT his_estructura.htetdesde, estr_cod.nrocod"
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " LEFT JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " AND tcodnro = 1"
    StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 18 AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(FecHasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(FecHasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!htetdesde) Then FechaIngCont = Format_Fecha(rs_Consult!htetdesde, Formato_Fecha, Fecha_Nula)
        If Not EsNulo(rs_Consult!nrocod) Then CodContr = rs_Consult!nrocod
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & FechaIngCont
    
    'Fecha inicio lic maternidad - Fecha de inicio de la primer licencia 11 dentro del rango
    DesdeLicMaternidad = Fecha_Nula
    StrSql = " SELECT elfechadesde "
    StrSql = StrSql & " FROM emp_lic"
    StrSql = StrSql & " WHERE Empleado = " & ternro
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " ( ( elfechadesde <= " & ConvFecha(FecDesde) & " ) AND ( elfechahasta >= " & ConvFecha(FecDesde) & " ) )"
    StrSql = StrSql & " OR"
    StrSql = StrSql & " ( ( elfechadesde >= " & ConvFecha(FecDesde) & " ) AND ( elfechadesde <= " & ConvFecha(FecHasta) & " ) )"
    StrSql = StrSql & " )"
    StrSql = StrSql & " AND tdnro = 11"
    StrSql = StrSql & " AND licestnro = 2"
    StrSql = StrSql & " ORDER BY elfechadesde"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!elfechadesde) Then DesdeLicMaternidad = Format_Fecha(rs_Consult!elfechadesde, Formato_Fecha, Fecha_Nula)
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & DesdeLicMaternidad

    'Fecha  fin   lic maternidad - Fecha de fin de la ultima licencia 11 dentro del rango
    HastaLicMaternidad = Fecha_Nula
    StrSql = " SELECT elfechahasta "
    StrSql = StrSql & " FROM emp_lic"
    StrSql = StrSql & " WHERE Empleado = " & ternro
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " ( ( elfechadesde <= " & ConvFecha(FecDesde) & " ) AND ( elfechahasta >= " & ConvFecha(FecDesde) & " ) )"
    StrSql = StrSql & " OR"
    StrSql = StrSql & " ( ( elfechadesde >= " & ConvFecha(FecDesde) & " ) AND ( elfechadesde <= " & ConvFecha(FecHasta) & " ) )"
    StrSql = StrSql & " )"
    StrSql = StrSql & " AND tdnro = 11"
    StrSql = StrSql & " AND licestnro = 2"
    StrSql = StrSql & " ORDER BY elfechahasta DESC"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!elfechahasta) Then HastaLicMaternidad = Format_Fecha(rs_Consult!elfechahasta, Formato_Fecha, Fecha_Nula)
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & HastaLicMaternidad
    
    'Cod grupo cargas sociales
    Aux_Linea = Aux_Linea & Separador & "1"
    
    'Grupo fiscal
    Aux_Linea = Aux_Linea & Separador & "1"
    
    'Cod Prestador Osocial
    Call BuscarEstructura(47, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cod Caracter de servicio
    Aux_Linea = Aux_Linea & Separador & "1"
    
    'Cod Sector Estadistico
    Call BuscarEstructura(50, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Situacion de revista
    '30/11/2007 - Martin Ferraro - Buscar tipo de codigo SIJP de la estructura
    'Call BuscarEstructura(30, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Call BuscarEstructuraTipoCod(30, 1, ternro, FecHasta, 0, "", EstructuraNro, EstructuraDesc)
    Aux_Linea = Aux_Linea & Separador & EstructuraNro & Separador & EstructuraDesc
    
    'Cod SIJP Contrato
    Aux_Linea = Aux_Linea & Separador & CodContr
    
    'Cod SIJP Modalidad de contrato
    CodModContr = ""
    StrSql = " SELECT estr_cod.nrocod"
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " AND tcodnro = 1"
    StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 31 AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(FecHasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(FecHasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!nrocod) Then CodModContr = rs_Consult!nrocod
    End If
    rs_Consult.Close
    Aux_Linea = Aux_Linea & Separador & CodModContr
    
    'CBU
    Aux_Linea = Aux_Linea & Separador & CBU
    
    'Imprimo la linea
    fExport.writeline Aux_Linea
    
    'Si el empleado estaba guardado en batch, entonces lo borro
    If tipo <> 3 Then
        Flog.writeline Espacios(Tabulador * 0) & "Borrando de batch_empleado."
        StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & bpronro & " And ternro = " & ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    '---------------------------------------------------------------------------------------------------------------
    'ACTUALIZO EL PROGRESO------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Empleados.MoveNext
    
Loop


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close


Set rs_Empleados = Nothing
Set rs_Consult = Nothing
Set rs_Aux = Nothing

Exit Sub

E_ExportTiger:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: ExportTiger"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


Sub BuscarEstructura(ByVal Tenro As Long, ByVal ternro As Long, ByVal Fecha As Date, ByVal EstrnroDefault As Long, ByVal EstrdabrDefault As String, ByRef Estrnro As Long, ByRef Estrdabr As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de busqueda de estructura tipo Tenro del ternro a la fecha
' Autor      : Martin Ferraro
' Fecha      : 15/05/2007
' --------------------------------------------------------------------------------------------

Dim rs_Estructura As New ADODB.Recordset
On Error GoTo E_BuscarEstructura

    Estrnro = EstrnroDefault
    Estrdabr = EstrdabrDefault
    
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = " & Tenro & " AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Estrnro = IIf(EsNulo(rs_Estructura!Estrnro), EstrnroDefault, rs_Estructura!Estrnro)
        Estrdabr = IIf(EsNulo(rs_Estructura!Estrdabr), EstrdabrDefault, rs_Estructura!Estrdabr)
    End If
    rs_Estructura.Close
    
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

Exit Sub

E_BuscarEstructura:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarEstructura"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Sub BuscarEstructuraTipoCod(ByVal Tenro As Long, ByVal TipoCod As Long, ByVal ternro As Long, ByVal Fecha As Date, ByVal EstrnroDefault As Long, ByVal EstrdabrDefault As String, ByRef Estrnro As Long, ByRef Estrdabr As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de busqueda de tipo de cod de estructura tipo Tenro del ternro a la fecha
' Autor      : Martin Ferraro
' Fecha      : 30/11/2007
' --------------------------------------------------------------------------------------------

Dim rs_Estructura As New ADODB.Recordset
On Error GoTo E_BuscarEstructura

    Estrnro = EstrnroDefault
    Estrdabr = EstrdabrDefault
    
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr, estr_cod.nrocod"
    StrSql = StrSql & " FROM his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " LEFT JOIN estr_cod ON estr_cod.estrnro = estructura.estrnro AND tcodnro = " & TipoCod
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND"
    StrSql = StrSql & " his_estructura.tenro = " & Tenro & " AND"
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") AND"
    StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Estrnro = IIf(EsNulo(rs_Estructura!Estrnro), EstrnroDefault, rs_Estructura!Estrnro)
        Estrdabr = IIf(EsNulo(rs_Estructura!nrocod), EstrdabrDefault, rs_Estructura!nrocod)
    End If
    rs_Estructura.Close
    
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

Exit Sub

E_BuscarEstructura:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarEstructura"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Sub BuscarNovedad(ByVal concepto As String, ByVal Parametro As Long, ByVal ternro As Long, ByVal Inicio As Date, ByVal Fin As Date, ByVal ValorDefault As Double, ByRef Valor As Double)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de busqueda de novedad
' Autor      : Martin Ferraro
' Fecha      : 15/05/2007
' --------------------------------------------------------------------------------------------

Dim rs_Novedad As New ADODB.Recordset
On Error GoTo E_BuscarNovedad
    
    Valor = 0
    'Novedades con vigencia
    StrSql = "SELECT novemp.* FROM novemp "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro"
    StrSql = StrSql & " AND concepto.conccod = '" & concepto & "'"
    StrSql = StrSql & " AND novemp.tpanro = " & Parametro
    StrSql = StrSql & " AND novemp.empleado = " & ternro
    StrSql = StrSql & " AND novemp.nevigencia = -1 "
    StrSql = StrSql & " AND (novemp.nevigencia = -1 "
    StrSql = StrSql & " AND novemp.nedesde <= " & ConvFecha(Fin)
    StrSql = StrSql & " AND (novemp.nehasta >= " & ConvFecha(Inicio)
    StrSql = StrSql & " OR novemp.nehasta is null )) "
    StrSql = StrSql & " ORDER BY novemp.nedesde, novemp.nehasta "
    OpenRecordset StrSql, rs_Novedad
    If Not rs_Novedad.EOF Then
    
        Do While Not rs_Novedad.EOF
            Valor = Valor + rs_Novedad!nevalor
            rs_Novedad.MoveNext
        Loop
        
        rs_Novedad.Close
    Else
    
        rs_Novedad.Close
    
        'Novedades sin vigencia
        StrSql = "SELECT novemp.* FROM novemp "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro"
        StrSql = StrSql & " AND concepto.conccod = '" & concepto & "'"
        StrSql = StrSql & " AND novemp.tpanro = " & Parametro
        StrSql = StrSql & " AND novemp.empleado = " & ternro
        StrSql = StrSql & " AND nevigencia = 0"
        StrSql = StrSql & " ORDER BY nedesde, nehasta "
        OpenRecordset StrSql, rs_Novedad
        If Not rs_Novedad.EOF Then
            Valor = rs_Novedad!nevalor
        Else
            Valor = ValorDefault
        End If
        
        rs_Novedad.Close
    End If
    
    
If rs_Novedad.State = adStateOpen Then rs_Novedad.Close
Set rs_Novedad = Nothing

Exit Sub

E_BuscarNovedad:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarNovedad"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="

    HuboError = True
End Sub


Sub BuscarDoc(ByVal ternro As Long, ByVal TidNroDefault As String, ByVal TidNomDefault As String, ByVal DocNroDefault As String, ByVal InstNroDefault As String, ByVal InstDesDefault As String, ByRef TidNro As String, ByRef TidNom As String, ByRef DocNro As String, ByRef InstNro As String, ByRef InstDes As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de busqueda de doc ternro
' Autor      : Martin Ferraro
' Fecha      : 15/05/2007
' --------------------------------------------------------------------------------------------

Dim rs_Documento As New ADODB.Recordset
On Error GoTo E_BuscarDoc

    TidNro = TidNroDefault
    TidNom = TidNomDefault
    DocNro = DocNroDefault
    InstNro = InstNroDefault
    InstDes = InstDesDefault
    
    StrSql = "SELECT ter_doc.nrodoc, tipodocu.tidnro, tipodocu.tidnom, tipodocu.tidsigla, institucion.instnro, institucion.instabre, ternro "
    StrSql = StrSql & "From ter_doc "
    StrSql = StrSql & "INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
    StrSql = StrSql & "INNER JOIN institucion ON institucion.instnro = tipodocu.instnro "
    StrSql = StrSql & "Where ter_doc.tidnro <= 5 "
    StrSql = StrSql & "AND  ter_doc.ternro = " & ternro
    StrSql = StrSql & "ORDER BY tipodocu.tidnro "

    OpenRecordset StrSql, rs_Documento
    If Not rs_Documento.EOF Then
        TidNro = IIf(EsNulo(rs_Documento!TidNro), TidNroDefault, rs_Documento!TidNro)
        TidNom = IIf(EsNulo(rs_Documento!tidsigla), TidNomDefault, rs_Documento!tidsigla)
        If Not EsNulo(rs_Documento!nrodoc) Then DocNro = Replace(rs_Documento!nrodoc, "-", "")
        InstNro = IIf(EsNulo(rs_Documento!InstNro), InstNroDefault, rs_Documento!InstNro)
        InstDes = IIf(EsNulo(rs_Documento!instabre), InstDesDefault, rs_Documento!instabre)
    End If
    rs_Documento.Close
    
If rs_Documento.State = adStateOpen Then rs_Documento.Close
Set rs_Documento = Nothing

Exit Sub

E_BuscarDoc:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarDoc"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub

Public Function Format_Fecha(ByVal Str As String, ByVal Formato As String, ByVal Nulo As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Convierte el string a fecha segun el formato de salida
' Autor      : Martin Ferraro
' Fecha      : 23/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

On Error GoTo E_Format_Fecha

    If Not EsNulo(Trim(Str)) Then
        Format_Fecha = Format(Trim(Str), Formato)
    Else
        Format_Fecha = Nulo
    End If
    
Exit Function

E_Format_Fecha:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: Format_Fecha"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
    
End Function


Public Sub FechaHora(ByVal l_schednro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que calcula la fecha y hora de la proxima ejecucion del proceso.
' Autor      : FGZ
' Fecha      : 17/12/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim l_frectipnro
Dim l_schedhora
Dim l_alesch_fecini
Dim l_alesch_fecfin
Dim l_alesch_frecrep
   
Dim rs As New ADODB.Recordset

On Error GoTo E_FechaHora

    StrSql = "SELECT frectipnro, alesch_fecini, schedhora, alesch_frecrep, alesch_fecfin "
    StrSql = StrSql & "FROM ale_sched "
    StrSql = StrSql & "WHERE schednro = " & l_schednro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        l_frectipnro = rs!frectipnro
        l_schedhora = rs!schedhora
        l_alesch_fecini = rs!alesch_fecini
        l_alesch_fecfin = rs!alesch_fecfin
        l_alesch_frecrep = rs!alesch_frecrep
        If (DateValue(Date) >= DateValue(l_alesch_fecini)) And (DateValue(Date) <= DateValue(l_alesch_fecfin)) Then
            ' Diariamiente
            If l_frectipnro = 1 Then
                Flog.writeline Espacios(Tabulador * 1) & "Planificacion Diaria."
                If FormatDateTime(Time, 4) > FormatDateTime(l_schedhora, 4) Then
                    'dia = CStr(CDate(Date + 1))
                    dia = CStr(DateAdd("d", Date, 1))
                    Hora = l_schedhora & ":00"
                Else
                    dia = Date
                    Hora = l_schedhora & ":00"
                End If
                Flog.writeline Espacios(Tabulador * 2) & "Dia: " & dia
                Flog.writeline Espacios(Tabulador * 2) & "Hora: " & Hora
            Else
                ' Mensualmente
                If l_frectipnro = 3 Then
                Flog.writeline Espacios(Tabulador * 1) & "Planificacion Mensual."
                    If Int(Day(Date)) < Int(l_alesch_frecrep) Then
                        'Debo programar para este mes para dia indicado
                        '/*controlar fecha*/
                        dia = DateValue(l_alesch_frecrep & "/" & Month(Date) & "/" & Year(Date))
                        Hora = l_schedhora & ":00"
                    Else
                        'Es el dia de la planificacion
                        If Int(Day(Date)) = Int(l_alesch_frecrep) Then
                            If Int(Left(FormatDateTime(Time, 4), 2) & Mid(Time, 4, 2)) + 1 > Int(Left(FormatDateTime(l_schedhora, 4), 2) & Mid(l_schedhora, 4, 2)) Then
                                'se paso de la hora, entonces planifico el mes siguiente
                                '/*controlar fecha*/
                                'dia = DateValue(l_alesch_frecrep & "/" & Int(Month(Date)) + 1 & "/" & Year(Date))
                                dia = DateValue(DateAdd("m", 1, l_alesch_frecrep & "/" & Int(Month(Date)) & "/" & Year(Date)))
                                Hora = l_schedhora & ":00"
                            Else
                                'Solo modifica la hora
                                dia = DateValue(l_alesch_frecrep & "/" & Month(Date) & "/" & Year(Date))
                                Hora = l_schedhora & ":00"
                            End If
                        Else
                            'Planificar para el mes siguiente
                                '/*controlar fecha*/
                                dia = DateValue(DateAdd("m", 1, l_alesch_frecrep & "/" & Int(Month(Date)) & "/" & Year(Date)))
                                'dia = DateValue(l_alesch_frecrep & "/" & Int(Month(Date)) + 1 & "/" & Year(Date))
                                Hora = l_schedhora & ":00"
                        End If
                    End If
                    Flog.writeline Espacios(Tabulador * 2) & "Dia: " & dia
                    Flog.writeline Espacios(Tabulador * 2) & "Hora: " & Hora
                Else
                    ' Semanalmente
                    If l_frectipnro = 2 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Planificacion Semanal."
                        If Weekday(Date) < Int(l_alesch_frecrep) Then
                            dia = DateValue(Date + l_alesch_frecrep - Weekday(Date))
                            Hora = l_schedhora & ":00"
                        Else
                            If Weekday(Date) > Int(l_alesch_frecrep) Then
                                dia = DateValue(Date + 7 - (Weekday(Date) - l_alesch_frecrep))
                                Hora = l_schedhora & ":00"
                            Else
                                If Int(Left(FormatDateTime(Time, 4), 2) & Mid(Time, 4, 2)) + 1 > Int(Left(FormatDateTime(l_schedhora, 4), 2) & Mid(l_schedhora, 4, 2)) Then
                                    dia = DateValue(Date + 7 - (Weekday(Date) - l_alesch_frecrep))
                                    Hora = l_schedhora
                                Else
                                    dia = Date
                                    Hora = l_schedhora & ":00"
                                End If
                            End If
                        End If
                        Flog.writeline Espacios(Tabulador * 2) & "Dia: " & dia
                        Flog.writeline Espacios(Tabulador * 2) & "Hora: " & Hora
                    Else
                        'Temporal
                        If l_frectipnro = 4 Then
                            Flog.writeline Espacios(Tabulador * 1) & "Planificacion Temporal."
                            dia = DateValue(Now + l_alesch_frecrep)
                            Hora = FormatDateTime(Now + l_alesch_frecrep, 4) & ":00"
                            Flog.writeline Espacios(Tabulador * 2) & "Dia: " & dia
                            Flog.writeline Espacios(Tabulador * 2) & "Hora: " & Hora
                        End If
                    End If
                End If
            End If
            ' Fecha siguiente fuera del tope maximo
            If DateValue(dia) > DateValue(l_alesch_fecfin) Then
                Flog.writeline Espacios(Tabulador * 1) & "ATENCION: No se puede activar la Planificacion, ya que la fecha de la proxima ejecucion esta fuera del rango de vigencia del schedule asociado. Verifique el mismo."
                RealizarPlanificacion = False
            End If
        Else
            'Fecha fuera de rango
            RealizarPlanificacion = False
        End If
    Else
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Schedule desconocido " & l_schednro
        Flog.writeline
        RealizarPlanificacion = False
    End If
    
    'cierro y libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
Exit Sub

E_FechaHora:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: FechaHora"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
    
End Sub

