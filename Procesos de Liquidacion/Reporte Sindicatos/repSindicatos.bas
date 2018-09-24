Attribute VB_Name = "repSindicatos"
Option Explicit

'Global Const Version = "1.0"
'Global Const FechaModificacion = "23/12/2015"
'Global Const UltimaModificacion = " " 'Version Inicial - Dimatz Rafael - CAS 33516 - Nuevo Reporte SEC

'Global Const Version = "1.1"
'Global Const FechaModificacion = "14/01/2016"
'Global Const UltimaModificacion = " " 'Gonzalez Nicolás - CAS-33516 - RH PRO (Producto) - Nuevo reporte SEC.
                                      'Se cambio nombre del proyecto VB y del ejecutable además de nombres internos/preparación para multiples reportes en un procesos
                                      'Correcciones varias en generar_SEC
'Global Const Version = "1.2"
'Global Const FechaModificacion = "20/01/2016"
'Global Const UltimaModificacion = "Nuevo reporte FAECYS" 'Gonzalez Nicolás - CAS-33516 - RH PRO (Producto) - Nuevo reporte FAECYS.
                                  'Nuevo reporte FAECYS

'Global Const Version = "1.3"
'Global Const FechaModificacion = "26/01/2016"
'Global Const UltimaModificacion = "Nuevo reporte FAECYS" 'Gonzalez Nicolás - CAS-33516 - RH PRO (Producto) - Nuevo reporte FAECYS.
                                                         'Se corrigen querys para obtener tipos de cód.

Global Const Version = "1.4"
Global Const FechaModificacion = "28/01/2016"
Global Const UltimaModificacion = " " 'Gonzalez Nicolás - CAS-35372 - RHPro (Producto) - ARG - NOM - Bug reportes legales con sit de revista
                                     ' Corrección de cód. a informar para cuando hay mas de 3 Situaciones de Revista
Global ArchExp

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Private Type TConfrepCO
    ConcCod As String
    Conccod_txt As String
    ConcNro As Long
End Type

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
Global EmpErrores As Boolean
Global Arrcod()
Global Arrdia()

Private Sub Main()

Dim NombreArchivo As String
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim Parametros As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim fechadesde
Dim fechahasta
Dim arr
Dim Empresa
Dim param
Dim Ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsleg  As New ADODB.Recordset
Dim Sql_leg As String
Dim legemp As Long

Dim cliqnro

Dim Retencion
Dim EmpAnt
Dim profecpago
Dim CantRegistros
Dim actualReg

Dim TiempoInicialProceso
Dim TiempoAcumulado
Dim PID As String
Dim bprcparam As String

Dim I As Integer
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
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "Reporte Sindicatos" & "-" & NroProceso & ".log"
    
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
    On Error GoTo 0

    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = " & 460
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        Flog.writeline "Ingresa a clasificar parametros"
        Parametros = rs!bprcparam
        
        If Not IsNull(Parametros) Then
            Flog.writeline "Parametros: " & Parametros
            ArrParametros = Split(Parametros, "@")
            Select Case CLng(ArrParametros(0))
                Case 1: 'SEC
                    Flog.writeline "------------------------------------------------------------------------------ "
                    Flog.writeline "------------------------------------- SEC ------------------------------------"
                    Flog.writeline "------------------------------------------------------------------------------ "
                    Call Generar_SEC(ArrParametros)
                Case 2: 'FAECYS
                    Flog.writeline "------------------------------------------------------------------------------ "
                    Flog.writeline "----------------------------------- FAECYS ----------------------------------- "
                    Flog.writeline "------------------------------------------------------------------------------ "
                    Call Generar_FAECYS(ArrParametros)
                Case Else
                    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Parametos de generación incorrectos."
                    Exit Sub
            End Select
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. La cantidad de parametros no es la esperada."
            Exit Sub
       End If
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
       
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & " Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & " Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub
Private Sub Generar_SEC(ByVal ArrParametros As Variant)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : Dimatz Rafael
' Fecha      : 03/12/2015
' Ultima Mod :
' Descripcion: Generacion de Reporte SEC
' 14/01/2016 - NG - Se modificó el nombre del sub a Generar_Sec + Correcciones varias
' ---------------------------------------------------------------------------------------------

'variables recibidas de parametro
Dim l_filtro As String

Dim l_concnro As Long
Dim l_empresa
Dim l_conceptonombre As String
Dim l_orden As String
Dim StrSql2 As String

Dim Nombre_Arch As String
Dim fs1, fs2

Dim Sep As String
Dim SepDec As String

Dim CantRegistros As Long
Dim Empresa As String

Dim Acumulador As Long

Dim l_Lista_ternro As String
Dim l_Lista_ternro_temp
Dim I As Long

Dim l_proaprob As Integer
Dim l_listaproc As String

Dim Cuil As String
Dim l_tipdoc As String
Dim l_nrodoc As String
Dim Nombre
Dim Apellido
Dim l_apellido As String

Dim l_tenro1 As Integer
Dim l_estrnro1 As Integer
Dim l_tenro2 As Integer
Dim l_estrnro2 As Integer
Dim l_tenro3 As Integer
Dim l_estrnro3 As Integer

Dim l_legdesde
Dim l_leghasta
Dim l_desde
Dim l_hasta
Dim l_pliqdesde
Dim l_pliqhasta
Dim l_empestrnro As Long
Dim l_pronro
Dim TipoCat
Dim Categoria
Dim l_categoria
Dim cliqnro

Dim RemuTotal
Dim l_RemuTotal
Dim CuotaSind
Dim l_CuotaSind
Dim Comercio

Dim l_Terape
Dim l_Ternom
Dim l_Ternom2
Dim l_Ternro
Dim l_Empleg
Dim Sucursal
Dim l_Sucursal
Dim Aux_Cod
Dim Aux_Cod_sitr
Dim Aux_Cod_sitr1
Dim Aux_diainisr1
Dim Aux_Cod_sitr2
Dim Aux_diainisr2
Dim Aux_Cod_sitr3
Dim Aux_diainisr3
Dim Fecha

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim l_rs As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset
    
    On Error GoTo ME_Local
    
    l_legdesde = ArrParametros(1)
    l_leghasta = ArrParametros(2)
    
    l_empestrnro = ArrParametros(3)
            
    l_desde = ArrParametros(4)
    l_hasta = ArrParametros(5)
    
    l_proaprob = ArrParametros(6)
    l_listaproc = ArrParametros(7)
    
    'Estructura 1
    l_tenro1 = ArrParametros(8)
    l_estrnro1 = ArrParametros(9)
    'Estructura 2
    l_tenro2 = ArrParametros(10)
    l_estrnro2 = ArrParametros(11)
    'Estructura 3
    l_tenro3 = ArrParametros(12)
    l_estrnro3 = ArrParametros(13)
    
    'Fecha seleccionada en el filtro
    If IsDate(ArrParametros(14)) Then
        Fecha = ArrParametros(14)
    Else
        Fecha = Date
    End If
    
    
           
    l_filtro = "((empleg >=" & l_legdesde & ") AND (empleg <= " & l_leghasta & ")) "
    If l_proaprob = 1 Then
            l_filtro = l_filtro & " AND (empest = 0 OR empest= -1)"
    Else
            l_filtro = l_filtro & " AND empest = " & l_proaprob
    End If
        
    On Error GoTo ME_Local
    
    Set fs2 = CreateObject("Scripting.FileSystemObject")
'    On Error Resume Next
    
    'Periodo Desde
    StrSql = "SELECT pliqdesde FROM periodo WHERE pliqnro = " & l_desde
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        l_pliqdesde = rs("pliqdesde")
    End If
    
    'Periodo Hasta
    StrSql = "SELECT pliqhasta FROM periodo WHERE pliqnro = " & l_hasta
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        l_pliqhasta = rs("pliqhasta")
    End If
    
    'Empresa
    StrSql = "SELECT empnom FROM empresa WHERE estrnro = " & l_empestrnro
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        l_empresa = rs("empnom")
    Else
        l_empresa = ""
    End If
      
    Categoria = 0
    RemuTotal = 0
    Sucursal = 0
    CuotaSind = 0
    'De aquí obtiene las columnas acumuladores para el reporte en confrep
    StrSql = "SELECT * FROM confrepAdv "
    StrSql = StrSql & " WHERE repnro = 500 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No hay configuración del reporte. "
        Exit Sub
    Else
        Do While Not rs.EOF
            Select Case CLng(rs("confnrocol"))
                Case 1: 'TE
                    Categoria = rs("confval")
                Case 2: 'AC
                    RemuTotal = rs("confval")
                Case 3: 'TE
                    Sucursal = rs("confval")
                Case 4: 'TE
                    CuotaSind = rs("confval")
                Case 5: 'EST
                    Comercio = rs("confval")
            End Select
            rs.MoveNext
        Loop
    End If
    
    'Pone a funcionar, de los filtros seleccionados, los que se eligen en ordenamiento
    If l_tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
'        Comentado por Carmen Quintero 26/03/2013
         StrSql = " SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1, "
         StrSql = StrSql & " estact2.tenro tenro2, estact2.estrnro estrnro2, estact3.tenro tenro3, estact3.estrnro estrnro3, cliqnro "
         StrSql = StrSql & " FROM cabliq "
         If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
         Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & l_tenro1 & " AND estact1.estrnro = " & Comercio
         StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_pliqhasta) & "))"
         If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
         StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_pliqdesde) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_pliqhasta) & "))"
         If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro =" & l_estrnro2
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & l_tenro3
         StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_pliqdesde) & " AND (estact3.htethasta IS NULL OR estact3.htethasta>=" & ConvFecha(l_pliqhasta) & "))"
         If l_estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
            StrSql = StrSql & " AND estact3.estrnro =" & l_estrnro3
         End If
         
         StrSql = StrSql & " WHERE " & l_filtro
            
    ElseIf l_tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel de estructura
'        Comentado por Carmen Quintero 26/03/2013
         StrSql = "SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1, "
         StrSql = StrSql & " estact2.tenro tenro2, estact2.estrnro estrnro2, cliqnro "
         StrSql = StrSql & " FROM cabliq "
         If l_listaproc = "" Or l_listaproc = "0" Then
           StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
         Else
           StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & l_tenro1 & " AND estact1.estrnro = " & Comercio
         StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_pliqhasta) & "))"
         If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
         StrSql = StrSql & " AND (estact2.htetdesde <= " & ConvFecha(l_pliqdesde) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >= " & ConvFecha(l_pliqhasta) & "))"
         If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro = " & l_estrnro2
         End If
         
         StrSql = StrSql & " WHERE " & l_filtro
                      
    ElseIf l_tenro1 <> 0 Then  ' Cuando solo seleccionamos el primer nivel de estructura
        StrSql = "SELECT DISTINCT empleado.ternro, empleg,  estact1.tenro tenro1, estact1.estrnro estrnro1, cliqnro "
        StrSql = StrSql & " FROM cabliq "
        If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & l_tenro1 & " AND estact1.estrnro = " & Comercio
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_pliqhasta) & "))"
        If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
        End If
        
        StrSql = StrSql & " WHERE " & l_filtro

    Else
        StrSql = " SELECT DISTINCT empleado.ternro, empleg, cliqnro "
        StrSql = StrSql & " FROM cabliq "
        If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.estrnro = " & Comercio
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_pliqhasta) & "))"
        StrSql = StrSql & " WHERE " & l_filtro
        
    End If
       Flog.writeline "Consulta general: " & StrSql
    'Busco los empleados
    OpenRecordset StrSql, l_rs
    
    ' _________________________________________________________________________
    'Flog.writeline "  SQL para control de los empleados pertenecientes al filtro seleccionado. "
    'Flog.writeline "    " & StrSql
    'Flog.writeline " "

    If l_rs.EOF Then
        CantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Procesar."
        'Flog.writeline "No se encontraron Empleados para el Reporte."
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
    Else
        
        CantRegistros = l_rs.RecordCount
        IncPorc = (99 / CantRegistros)
            
            Do Until l_rs.EOF 'Para todos los que cumplen con el filtro elegido + SEC por las dudas
                cliqnro = l_rs("cliqnro")
                l_Ternro = l_rs("ternro")
                l_Empleg = l_rs("empleg")
                                        
                'Busco el CUIL del empleado
                StrSql = " SELECT nrodoc "
                StrSql = StrSql & " FROM ter_doc "
                StrSql = StrSql & " WHERE ternro = " & l_rs("ternro")
                StrSql = StrSql & " AND tidnro = 10 "
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró el cuit para el empleado.ternro(" & l_rs("ternro") & ")"
                    Cuil = " "
                Else
                    Cuil = Left(CStr(rs2("nrodoc")), 13)
                End If
                rs2.Close
                Cuil = Replace(Cuil, "-", "")
                
                'Apellido y Nombre
                StrSql = "SELECT ternom,ternom2, terape "
                StrSql = StrSql & "FROM empleado "
                StrSql = StrSql & "WHERE empleado.ternro = " & l_rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                  l_Terape = ""
                  l_Ternom = ""
                  l_Ternom2 = ""
                Else
                  l_Terape = rs2("terape")
                  l_Ternom = rs2("ternom")
                  l_Ternom2 = rs2("ternom2")
                End If
                rs2.Close
                
                'Remuneracion Total
                StrSql = " SELECT almonto"
                StrSql = StrSql & " From acu_liq"
                StrSql = StrSql & " Where acunro = " & RemuTotal
                StrSql = StrSql & " AND cliqnro = " & cliqnro
                OpenRecordset StrSql, rs
                If Not rs.EOF Then
                    l_RemuTotal = rs!almonto
                Else
                    l_RemuTotal = 0
                End If
                
                
                
                'Categoria | Retorna cod. asociado a la estructura para cada empleado a la fecha seleccionada en el filtro
                
                'StrSql = " SELECT nrocod FROM estr_cod"
                'StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
                'StrSql = StrSql & " WHERE (tipocod.tcodnro = 197)"
                'StrSql = StrSql & " AND estrnro = " & Categoria
                
                StrSql = " SELECT his_estructura.ternro,estrdabr,his_estructura.estrnro,nrocod"
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
                StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = 197"
                StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
                StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha)
                StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(Fecha) & ")"
                StrSql = StrSql & " AND his_estructura.tenro = " & Categoria
                StrSql = StrSql & " AND his_estructura.ternro=" & l_rs("ternro")
                OpenRecordset StrSql, rs
                l_categoria = 0
                If Not rs.EOF Then
                    If Not EsNulo(rs("nrocod")) Then
                        l_categoria = rs("nrocod")
                    End If
                End If

                
                'Descuento Cuota Sindical
                'l_CuotaSind
                StrSql = " SELECT estrdabr "
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                StrSql = StrSql & " AND htetdesde <= " & ConvFecha(l_pliqhasta)
                StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(l_pliqhasta) & ") "
                StrSql = StrSql & " AND his_estructura.tenro = " & CuotaSind
                StrSql = StrSql & " AND his_estructura.ternro = " & l_rs("ternro")
                OpenRecordset StrSql, rs
                
                l_CuotaSind = ""
                'Flog.writeline "Query Cuota Sindical " & StrSql
                If Not rs.EOF Then
                   l_CuotaSind = rs!estrdabr
                End If
                              
                'Sucursal
                StrSql = " SELECT estrdabr "
                StrSql = StrSql & " FROM his_estructura"
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                StrSql = StrSql & " AND htetdesde <= " & ConvFecha(l_pliqhasta) & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(l_pliqhasta) & ")"
                StrSql = StrSql & " AND his_estructura.tenro = " & Sucursal
                StrSql = StrSql & " AND his_estructura.ternro = " & l_rs("ternro")
                       
                OpenRecordset StrSql, rs

                l_Sucursal = ""

                If Not rs.EOF Then
                   l_Sucursal = rs!estrdabr
                Else
                   Flog.writeline "Error al obtener los datos de la sucursal"
                '   GoTo MError
                End If
                
                '----------------------------------------------------------------------------------------------------------------------------------
                Flog.writeline "Buscar Situacion de Revista Actual"
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                'Buscar Situacion de Revista Actual
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & l_rs("ternro") & " AND "
                StrSql = StrSql & " his_estructura.tenro = 30 AND "
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(l_pliqhasta) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(l_pliqdesde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                         
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
    
                Flog.writeline "inicializo"
                Aux_Cod_sitr1 = ""
                Aux_diainisr1 = 0
                Aux_Cod_sitr2 = ""
                Aux_diainisr2 = 0
                Aux_Cod_sitr3 = ""
                Aux_diainisr3 = 0
                Select Case rs_Estructura.RecordCount
                Case 0:
                        'EAM- Si no tiene situación de revista, busca si tiene alguna estructura 'REV' en el confrep
                        'Aux_Cod_sitr1 = Buscar_SituacionRevistaConfig(l_rs("Ternro"), fechadesde, fechahasta)
                        Aux_diainisr1 = 1
                        If Aux_Cod_sitr1 <> 0 Then
                            Flog.writeline "Se asignó la situación de revista del confrep: " & Aux_Cod_sitr1
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        Else
                            Flog.writeline "no hay situaciones de revista"
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        End If
                Case 1:
                    'Aux_Cod_sitr1 = rs_Estructura!estrcodext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 198"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = CStr(rs_Estr_cod!nrocod)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If rs_Estructura!htetdesde < l_pliqdesde Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = Day(rs_Estructura!htetdesde)
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > Day(l_pliqhasta) Then
                        Aux_diainisr1 = CStr(Day(l_pliqhasta))
                    End If
                    
                    Aux_Cod_sitr = Aux_Cod_sitr1
                    Flog.writeline "hay 1 situaciones de revista"
                Case 2:
                    'Primer situacion
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 198"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = CStr(rs_Estr_cod!nrocod)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If rs_Estructura!htetdesde < l_pliqdesde Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = Day(rs_Estructura!htetdesde)
                    End If

                    If CInt(Aux_diainisr1) > Day(l_pliqhasta) Then
                        Aux_diainisr1 = CStr(Day(l_pliqhasta))
                    End If
                    
                    'siguiente situacion
                    rs_Estructura.MoveNext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 198"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        If Aux_Cod_sitr1 <> CStr(rs_Estr_cod!nrocod) Then
                            If Not rs_Estr_cod.EOF Then
                                Aux_Cod_sitr2 = CStr(rs_Estr_cod!nrocod)
                            Else
                                Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                Aux_Cod_sitr2 = 1
                            End If
                            Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                            If CInt(Aux_diainisr2) > Day(l_pliqhasta) Then
                                Aux_diainisr2 = CStr(Day(l_pliqhasta))
                            End If
                            Aux_Cod_sitr = Aux_Cod_sitr2
                            Flog.writeline "hay 2 situaciones de revista"
                        Else
                            'Si es la misma sit de revista ==> le asigno la anterior
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        End If
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr2 = 1
                    End If
                    
                Case 3:
                    'Primer situacion (1)
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 198"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = CStr(rs_Estr_cod!nrocod)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If rs_Estructura!htetdesde < l_pliqdesde Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = Day(rs_Estructura!htetdesde)
                    End If

                    If CInt(Aux_diainisr1) > Day(l_pliqhasta) Then
                        Aux_diainisr1 = CStr(Day(l_pliqhasta))
                    End If
                    
                    'siguiente situacion (2)
                    rs_Estructura.MoveNext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 198"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        If Aux_Cod_sitr1 <> CStr(rs_Estr_cod!nrocod) Then
                            If Not rs_Estr_cod.EOF Then
                                Aux_Cod_sitr2 = CStr(rs_Estr_cod!nrocod)
                            Else
                                Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                Aux_Cod_sitr2 = 1
                            End If
                            Aux_diainisr2 = Day(rs_Estructura!htetdesde)
    
                            If CInt(Aux_diainisr2) > Day(l_pliqhasta) Then
                                Aux_diainisr2 = CStr(Day(l_pliqhasta))
                            End If
                        Else
                            'Si es la misma sit de revista ==> le asigno la anterior
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        End If
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr2 = 1
                    End If
                    
                    'siguiente situacion (3)
                    rs_Estructura.MoveNext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 198"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        If Aux_Cod_sitr2 <> CStr(rs_Estr_cod!nrocod) Then
                            If Not rs_Estr_cod.EOF Then
                                If Aux_Cod_sitr2 <> "" Then
                                    Aux_Cod_sitr3 = CStr(rs_Estr_cod!nrocod)
                                Else
                                    Aux_Cod_sitr2 = CStr(rs_Estr_cod!nrocod)
                                End If
                            Else
                                Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                Aux_Cod_sitr3 = 1
                            End If
                            If Aux_Cod_sitr3 <> "" Then 'Agregado ver 1.37 - JAZ
                                Aux_diainisr3 = Day(rs_Estructura!htetdesde)
                            Else
                                Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                            End If
                            'FGZ - 08/07/2005
                            If Aux_diainisr3 <> "" Then 'Agregado ver 1.37 - JAZ
                                If CInt(Aux_diainisr3) > Day(l_pliqhasta) Then
                                    Aux_diainisr3 = CStr(Day(l_pliqhasta))
                                End If
                                Aux_Cod_sitr = Aux_Cod_sitr3
                            Else
                                If CInt(Aux_diainisr2) > Day(l_pliqhasta) Then
                                    Aux_diainisr2 = CStr(Day(l_pliqhasta))
                                End If
                                Aux_Cod_sitr = Aux_Cod_sitr2
                            End If
                            Flog.writeline "hay 3 situaciones de revista"
                        Else 'FGZ - 11/01/2012 --------------------------------------------
                            'Si es la misma sit de revista ==> le asigno la anterior
                            If Aux_Cod_sitr2 <> "" Then 'Modificado ver 1.36 - JAZ
                                Aux_Cod_sitr = Aux_Cod_sitr2
                            Else
                                Aux_Cod_sitr = Aux_Cod_sitr1
                            End If
                        End If
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr3 = 1
                        If Aux_Cod_sitr3 <> "" Then 'Agregado ver 1.37 - JAZ
                            Aux_diainisr3 = Day(rs_Estructura!htetdesde)
                        Else
                            Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                        End If
                        'FGZ - 08/07/2005
                        If Aux_diainisr3 <> "" Then 'Agregado ver 1.37 - JAZ
                            If CInt(Aux_diainisr3) > Day(l_pliqhasta) Then
                                Aux_diainisr3 = CStr(Day(l_pliqhasta))
                            End If
                            Aux_Cod_sitr = Aux_Cod_sitr3
                        Else
                            If CInt(Aux_diainisr2) > Day(l_pliqhasta) Then
                                Aux_diainisr2 = CStr(Day(l_pliqhasta))
                            End If
                            Aux_Cod_sitr = Aux_Cod_sitr2
                        End If
                        Flog.writeline "hay 3 situaciones de revista"
                End If
               Case Else
                   If Not rs_Estructura.EOF Then
                      Dim k
                      k = 0
                      Do While Not rs_Estructura.EOF
                         If (k = 0) Then
                              ReDim Preserve Arrcod(k)
                              ReDim Preserve Arrdia(k)
                              
                              'Arrcod(k) = rs_Estructura!estrcodext 'NG - v 1.4
                              
                              StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                              StrSql = StrSql & " AND tcodnro = 198"
                              OpenRecordset StrSql, rs_Estr_cod
                              If Not rs_Estr_cod.EOF Then
                                Arrcod(k) = Left(CStr(rs_Estr_cod!nrocod), 2)
                              Else
                                Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                Arrcod(k) = 1
                              End If

                                                            
                              'Arrdia(k) = Day(rs_Estructura!htetdesde)
                              'Agregado 02/10/2014
                              If rs_Estructura!htetdesde < l_pliqdesde Then
                                 Arrdia(k) = 1
                              Else
                                 Arrdia(k) = Day(rs_Estructura!htetdesde)
                              End If
                              'fin
                              k = k + 1
                         Else
                         
                              'NG - v 1.4------------------------------------------------------------------------
                              Aux_Cod = ""
                              StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                              StrSql = StrSql & " AND tcodnro = 198"
                              OpenRecordset StrSql, rs_Estr_cod
                              If Not rs_Estr_cod.EOF Then
                                Aux_Cod = Left(CStr(rs_Estr_cod!nrocod), 2)
                              Else
                                Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                Aux_Cod = 1
                              End If
                              'NG - v 1.4------------------------------------------------------------------------
                              
                         
                             If Arrcod(k - 1) <> Aux_Cod Then
                                  ReDim Preserve Arrcod(k)
                                  ReDim Preserve Arrdia(k)
                                  'Arrcod(k) = rs_Estructura!estrcodext 'NG - v 1.4
                                  Arrcod(k) = Aux_Cod
                                  If rs_Estructura!htetdesde < l_pliqdesde Then
                                      Arrdia(k) = 1
                                  Else
                                      Arrdia(k) = Day(rs_Estructura!htetdesde)
                                  End If
                                  k = k + 1
                             End If
                         End If
        
                         rs_Estructura.MoveNext
                      Loop
                    End If
                         
                     If UBound(Arrcod) >= 2 Then
                        Aux_Cod_sitr3 = Arrcod(UBound(Arrcod))
                        Aux_Cod_sitr2 = Arrcod(UBound(Arrcod) - 1)
                        Aux_Cod_sitr1 = Arrcod(UBound(Arrcod) - 2)
                        Aux_diainisr3 = Arrdia(UBound(Arrdia))
                        Aux_diainisr2 = Arrdia(UBound(Arrdia) - 1)
                        Aux_diainisr1 = Arrdia(UBound(Arrdia) - 2)
                     End If
                    
                     If UBound(Arrcod) = 1 Then
                        Aux_Cod_sitr3 = ""
                        Aux_Cod_sitr2 = Arrcod(UBound(Arrcod))
                        Aux_Cod_sitr1 = Arrcod(UBound(Arrcod) - 1)
                        Aux_diainisr3 = ""
                        Aux_diainisr2 = Arrdia(UBound(Arrdia))
                        Aux_diainisr1 = Arrdia(UBound(Arrdia) - 1)
                     End If
                    
                     If UBound(Arrcod) = 0 Then
                        Aux_Cod_sitr3 = ""
                        Aux_Cod_sitr2 = ""
                        Aux_Cod_sitr1 = Arrcod(UBound(Arrcod))
                        Aux_diainisr3 = ""
                        Aux_diainisr2 = ""
                        Aux_diainisr1 = Arrdia(UBound(Arrdia))
                     End If
                     
                     Aux_Cod_sitr = Arrcod(UBound(Arrcod))
                    'fin
                     Flog.writeline "hay + de 3 situaciones de revista"
'----------------------------------------------------------------------------------------------------------------------------------
               End Select
                'No puede haber situaciones de revista iguales consecutivas.
                'Antes ese caso, me quedo con la primera de las iguales y consecutivas
                If Aux_Cod_sitr3 = Aux_Cod_sitr2 Then
                    'Elimino la situacion de revista 3
                    Aux_Cod_sitr3 = ""
                    Aux_diainisr3 = ""
                End If
                If Aux_Cod_sitr2 = Aux_Cod_sitr1 Then
                    'Elimino la situacion de revista 2 y la 3 la pongo en la 2
                    Aux_Cod_sitr2 = Aux_Cod_sitr3
                    Aux_diainisr2 = Aux_diainisr3
                    
                    Aux_Cod_sitr3 = ""
                    Aux_diainisr3 = ""
                End If
                
'-------------------------------------------------- Reporte Faecys ----------------------------------------------------------------
                'Ingreso
                'Media jornada:
                'Licencia
'---------------------------------------------- Fin Reporte Faecys ----------------------------------------------------------------
                
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                CantRegistros = CantRegistros - 1
                Progreso = Progreso + IncPorc
                Progreso = 2
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                StrSql = StrSql & ", bprcempleados ='" & CStr(CantRegistros) & "' WHERE bpronro = " & NroProceso
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                StrSql = " INSERT INTO rep_SEC "
                StrSql = StrSql & "(bpronro, cuil, terape, ternom, ternom2, empresa, "
                StrSql = StrSql & " RemuTotal, Categoria, CuotaSind, Sucursal,ternro,empleg,"
                StrSql = StrSql & " SituacionLab1,SituacionLab2,SituacionLab3,DiaDesde1,DiaDesde2,DiaDesde3,"
                StrSql = StrSql & " pliqdesde,pliqhasta,listaproc) "
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuil & "'"
                StrSql = StrSql & ",'" & l_Terape & "'"
                StrSql = StrSql & ",'" & l_Ternom & "'"
                StrSql = StrSql & ",'" & l_Ternom2 & "'"
                StrSql = StrSql & ",'" & l_empresa & "'"
                StrSql = StrSql & "," & Replace(l_RemuTotal, ",", ".")
                StrSql = StrSql & "," & l_categoria
                StrSql = StrSql & ",'" & l_CuotaSind & "'"
                StrSql = StrSql & ",'" & l_Sucursal & "'"
                StrSql = StrSql & "," & l_Ternro
                StrSql = StrSql & "," & l_Empleg
                StrSql = StrSql & ",'" & Aux_Cod_sitr1 & "'"
                StrSql = StrSql & ",'" & Aux_Cod_sitr2 & "'"
                StrSql = StrSql & ",'" & Aux_Cod_sitr3 & "'"
                StrSql = StrSql & "," & IIf(Aux_diainisr1 = "", 0, Aux_diainisr1)
                StrSql = StrSql & "," & IIf(Aux_diainisr2 = "", 0, Aux_diainisr2)
                StrSql = StrSql & "," & IIf(Aux_diainisr3 = "", 0, Aux_diainisr3)
                StrSql = StrSql & ",'" & l_pliqdesde & "'"
                StrSql = StrSql & ",'" & l_pliqhasta & "'"
                StrSql = StrSql & ",'" & l_listaproc & "'"
                StrSql = StrSql & ")"

                objConn.Execute StrSql, , adExecuteNoRecords
                
                l_rs.MoveNext
            Loop
        End If

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub


Private Sub Generar_FAECYS(ByVal ArrParametros As Variant)
' ---------------------------------------------------------------------------------------------
' Descripcion: Generacion de Reporte FAECYS (CABA)
' Autor      : Gonzalez Nicolás
' Fecha      : 14/01/2016
' Ultima Mod : 26/01/2016 - Se modificaron querys que obtienen cód. asociados a estructuras
' ---------------------------------------------------------------------------------------------
Dim Filtro As String
Dim Orden As String
Dim Empnom As String
Dim CantRegistros As Long
Dim Proaprob As Integer
Dim Listaproc As String
Dim Cuil As String
Dim Tenro1 As Integer
Dim Estrnro1 As Integer
Dim Tenro2 As Integer
Dim Estrnro2 As Integer
Dim Tenro3 As Integer
Dim Estrnro3 As Integer
Dim Legdesde
Dim Leghasta
Dim Desde
Dim Hasta
Dim pliqdesde
Dim pliqhasta
Dim Empestrnro As Long
Dim CategoriaCod
Dim Categoria As String
Dim cliqnro
Dim RemuTotal
Dim Ternom
Dim Ternom2
Dim Ternro
Dim Terape
Dim Empleg
Dim Sucursal
Dim SucursalTernro As Long
Dim SucursalEstrnro As Long
Dim SucursalCod
Dim SucCalle
Dim SucNro
Dim SucLocnro
Dim SucProvnro
Dim SucCodLoc As String
Dim SucCodProv As String
Dim Auxiliar As String
Dim Pulgas As String
Dim CodCentraliza As String
Dim Fecha
Dim FecIngreso As Date
Dim MensajeErr As String
Dim EmpTernro As Long
Dim MediaJor As String
'Confrep --------------------------
Dim TenroSucursal
Dim TipoCodSuc
Dim TipoCodSucCen

Dim TenroCategoria
Dim TipoCodCat

Dim TenroRegHor
Dim TipoCodRegHor
Dim AC_RemuTotal
Dim listaTipoLic As String
Dim Licencia As String
'----------------------------------
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Longitud As Long


Longitud = 20 'SE UTILIZA PARA IMPRIMIR EL FLOG
    
On Error GoTo ME_Local


'CONTROLO PARÁMETROS
Legdesde = ArrParametros(1)
Leghasta = ArrParametros(2)

Empestrnro = ArrParametros(3)
        
Desde = ArrParametros(4)
Hasta = ArrParametros(5)

Proaprob = ArrParametros(6)
Listaproc = ArrParametros(7)

'Estructura 1
Tenro1 = ArrParametros(8)
Estrnro1 = ArrParametros(9)
'Estructura 2
Tenro2 = ArrParametros(10)
Estrnro2 = ArrParametros(11)
'Estructura 3
Tenro3 = ArrParametros(12)
Estrnro3 = ArrParametros(13)

'Fecha seleccionada en el filtro
If IsDate(ArrParametros(14)) Then
    Fecha = ArrParametros(14)
Else
    Fecha = Date
End If

'Orden
Orden = " ORDER BY empleado." & Trim(ArrParametros(15))

           
Filtro = "((empleg >=" & Legdesde & ") AND (empleg <= " & Leghasta & ")) "
If Proaprob = 1 Then
    Filtro = Filtro & " AND (empest = 0 OR empest= -1)"
Else
    Filtro = Filtro & " AND empest = " & Proaprob
End If




'----------------------------------------------------------------------
'Periodo Desde
StrSql = "SELECT pliqdesde FROM periodo WHERE pliqnro = " & Desde
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If Not rs.EOF Then
    pliqdesde = rs("pliqdesde")
End If
    
'----------------------------------------------------------------------
'Periodo Hasta
StrSql = "SELECT pliqhasta FROM periodo WHERE pliqnro = " & Hasta
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If Not rs.EOF Then
    pliqhasta = rs("pliqhasta")
End If
    
'----------------------------------------------------------------------
'DATOS EMPRESA
StrSql = "SELECT empnom FROM empresa WHERE estrnro = " & Empestrnro
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Empnom = rs("empnom")
Else
    Empnom = ""
End If


'----------------------------------------------------------------------
'CONFREP
'----------------------------------------------------------------------
'SUCURSAL
TenroSucursal = 0
TipoCodSuc = 0
TipoCodSucCen = 0
'CATEGORIA
TenroCategoria = 0
TipoCodCat = 0

TenroCategoria = 0
AC_RemuTotal = 0
TenroRegHor = 0
TipoCodRegHor = 0
listaTipoLic = ""


'De aquí obtiene las columnas acumuladores para el reporte en confrep
StrSql = "SELECT confnrocol,confval,confval2,confval3 FROM confrepAdv "
StrSql = StrSql & " WHERE repnro = 503 "
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    Flog.writeline "No hay configuración del reporte. "
    Exit Sub
Else
    Do While Not rs.EOF
        Select Case CLng(rs!confnrocol)
            Case 1: 'SUCURSAL
                TenroSucursal = rs!confval
                TipoCodSuc = rs!confval2
                TipoCodSucCen = rs!confval3
            Case 2: 'CATEGORIA
                TenroCategoria = rs!confval
                TipoCodCat = rs!confval2
            Case 3: 'REGIMEN HORARIO
                TenroRegHor = rs!confval
                TipoCodRegHor = rs!confval2
            Case 4: 'AC
                AC_RemuTotal = rs!confval
            Case 5: 'LICENCIAS
                listaTipoLic = rs!confval
        End Select
        rs.MoveNext
    Loop
End If
    
'Pone a funcionar, de los filtros seleccionados, los que se eligen en ordenamiento
If Tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
    StrSql = " SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1 "
    StrSql = StrSql & ",empleado.ternom,empleado.ternom2,empleado.terape,empleado.terape2 "
    StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2, estact3.tenro tenro3, estact3.estrnro estrnro3, cliqnro "
    StrSql = StrSql & " FROM cabliq "
    If Listaproc = "" Or Listaproc = "0" Then
       StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
    Else
       StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & Listaproc & ") "
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1 & " AND estact1.estrnro = " & Empestrnro
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(pliqhasta) & "))"
    If Estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
       StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2
    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(pliqdesde) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(pliqhasta) & "))"
    If Estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
       StrSql = StrSql & " AND estact2.estrnro =" & Estrnro2
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & Tenro3
    StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(pliqdesde) & " AND (estact3.htethasta IS NULL OR estact3.htethasta>=" & ConvFecha(pliqhasta) & "))"
    If Estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
       StrSql = StrSql & " AND estact3.estrnro =" & Estrnro3
    End If
    
    StrSql = StrSql & " WHERE " & Filtro
            
ElseIf Tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel de estructura
    StrSql = "SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1 "
    StrSql = StrSql & ",empleado.ternom,empleado.ternom2,empleado.terape,empleado.terape2 "
    StrSql = StrSql & ",estact2.tenro tenro2, estact2.estrnro estrnro2, cliqnro "
    StrSql = StrSql & " FROM cabliq "
    If Listaproc = "" Or Listaproc = "0" Then
      StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
    Else
      StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & Listaproc & ") "
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1 & " AND estact1.estrnro = " & Empestrnro
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(pliqhasta) & "))"
    If Estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
       StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2
    StrSql = StrSql & " AND (estact2.htetdesde <= " & ConvFecha(pliqdesde) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >= " & ConvFecha(pliqhasta) & "))"
    If Estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
       StrSql = StrSql & " AND estact2.estrnro = " & Estrnro2
    End If
    
    StrSql = StrSql & " WHERE " & Filtro
                      
ElseIf Tenro1 <> 0 Then  ' Cuando solo seleccionamos el primer nivel de estructura
    StrSql = "SELECT DISTINCT empleado.ternro, empleg,  estact1.tenro tenro1, estact1.estrnro estrnro1, cliqnro "
    StrSql = StrSql & ",empleado.ternom,empleado.ternom2,empleado.terape,empleado.terape2 "
    StrSql = StrSql & " FROM cabliq "
    If Listaproc = "" Or Listaproc = "0" Then
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
    Else
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & Listaproc & ") "
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1 & " AND estact1.estrnro = " & Empestrnro
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(pliqhasta) & "))"
    If Estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
        StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
    End If
    
    StrSql = StrSql & " WHERE " & Filtro

Else
    StrSql = " SELECT DISTINCT empleado.ternro, empleg, cliqnro "
    StrSql = StrSql & ",empleado.ternom,empleado.ternom2,empleado.terape,empleado.terape2 "
    StrSql = StrSql & " FROM cabliq "
    If Listaproc = "" Or Listaproc = "0" Then
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
    Else
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & Listaproc & ") "
    End If
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.estrnro = " & Empestrnro
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(pliqdesde) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(pliqhasta) & "))"
    StrSql = StrSql & " WHERE " & Filtro
End If

StrSql = StrSql & Orden

'Flog.writeline StrSql
OpenRecordset StrSql, rs
If rs.EOF Then
    CantRegistros = 1
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Procesar."
    StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    Exit Sub
Else
    CantRegistros = rs.RecordCount
    IncPorc = (99 / CantRegistros)
    Do Until rs.EOF
    
        Pulgas = ""
        
        cliqnro = rs("cliqnro")
        Ternro = rs("ternro")
        Empleg = rs("empleg")
        
        'NOMBRES
        Terape = rs("terape")
        If Not EsNulo(rs!terape2) Then
            Terape = Terape & " " & rs("terape2")
        End If
        'APELLIDOS
        Ternom = rs("ternom")
        If Not EsNulo(rs!Ternom2) Then
            Ternom = Ternom & " " & rs("ternom2")
        End If
        
        Flog.writeline ""
        Flog.writeline "****************************************************************************************************"
        Flog.writeline " EMPLEADO " & "(" & Empleg & ") : " & UCase(Terape) & ", " & UCase(Ternom)
        Flog.writeline "****************************************************************************************************"
        
        '---------------------------------------------------------------------
        'SUCURSAL | Datos de Sucursal + Tipo de código + Direccion
        '---------------------------------------------------------------------
        Sucursal = ""
        SucursalTernro = 0
        SucursalCod = Empty
        MensajeErr = ""
        'StrSql = " SELECT sucursal.ternro,estructura.estrdabr, his_estructura.estrnro,nrocod "
        StrSql = " SELECT sucursal.ternro,estructura.estrdabr, his_estructura.estrnro "
        StrSql = StrSql & ", (SELECT nrocod FROM estr_cod INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro WHERE estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & TipoCodSuc & ") nrocod"
        StrSql = StrSql & " FROM his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
        'StrSql = StrSql & " LEFT JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & TipoCodSuc
        'StrSql = StrSql & " LEFT JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
        
        StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(Fecha) & ")"
        StrSql = StrSql & " AND his_estructura.tenro = " & TenroSucursal
        StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            SucursalTernro = rs2!Ternro
            SucursalEstrnro = rs2!Estrnro
            Sucursal = rs2!estrdabr
            SucursalCod = IIf(EsNulo(rs2!nrocod), Empty, rs2!nrocod)
        Else
            Pulgas = "Sucursal"
            MensajeErr = "No se encontraron datos a la fecha : " & Fecha
        End If
        
        If MensajeErr = "" Then
            If IsEmpty(SucursalCod) Then
                MensajeErr = "No se encontró tipo de código asociado a la estructura : " & Sucursal
                MensajeErr = MensajeErr & ". Ver configuración del reporte columna 1 valor 2 (Tipos de Código)"
            Else
                'CONTROLO QUE SEA UN NÚMERO
                If Not IsNumeric(SucursalCod) Then
                    MensajeErr = "EL código debe ser numérico."
                    MensajeErr = MensajeErr & ". Ver configuración del reporte columna 1 valor 2 (Tipos de Código)"
                End If
            End If
        End If
        'Flog.writeline Espacios(Tabulador * 1) & "SUCURSAL: " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucursalCod)
        Flog.writeline Format_StrLR("SUCURSAL", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucursalCod)
        
        '---------------------------------------------------------------------
        'DATOS DEL DOMICILIO DE LA SUCURSAL (DEFAULT)
        '---------------------------------------------------------------------
        SucCalle = ""
        SucNro = ""
        MensajeErr = ""
        SucLocnro = 0
        SucProvnro = 0
        SucCodLoc = ""
        SucCodProv = ""
        Auxiliar = ""
        If SucursalTernro <> 0 Then
            StrSql = "SELECT calle,nro,locnro,provnro "
            StrSql = StrSql & " FROM cabdom "
            StrSql = StrSql & " INNER JOIN detdom on detdom.domnro = cabdom.domnro"
            StrSql = StrSql & " WHERE domdefault = -1 AND cabdom.Ternro = " & SucursalTernro
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                SucCalle = rs2!Calle
                SucNro = rs2!nro
                SucLocnro = IIf(EsNulo(rs2!locnro), 0, rs2!locnro)
                SucProvnro = IIf(EsNulo(rs2!provnro), 0, rs2!provnro)
            Else
                MensajeErr = "No se encontró domicilio default a para la sucursal : " & Sucursal
            End If
            'Flog.writeline Espacios(Tabulador * 2) & "CALLE y N°: " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucCalle & " " & SucNro)
            Flog.writeline Format_StrLR("CALLE y N°", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucCalle & " " & SucNro)
            If MensajeErr = "" Then
                'CODIGO DE LOCALIDAD
                If SucLocnro <> 0 Then
                    StrSql = "SELECT locdesc,loccodext FROM localidad WHERE locnro = " & SucLocnro
                    OpenRecordset StrSql, rs2
                    If Not rs2.EOF Then
                        SucCodLoc = IIf(EsNulo(rs2!loccodext), Empty, rs2!loccodext)
                        Auxiliar = rs2!locdesc
                    End If
                    
                    If EsNulo(SucCodLoc) = True And Auxiliar = "" Then
                        MensajeErr = "No se encontró localidad en el domicilio configurado."
                    ElseIf EsNulo(SucCodLoc) = True Then
                        MensajeErr = "No se encontró el código externo de la localidad : " & Auxiliar
                    End If
                Else
                    MensajeErr = "No se encontró localidad en el domicilio configurado."
                End If
                'Flog.writeline Espacios(Tabulador * 2) & "CÓDIGO DE LOCALIDAD: " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucCodLoc)
                Flog.writeline Format_StrLR("CÓDIGO DE LOCALIDAD", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucCodLoc)
                
                Auxiliar = ""
                MensajeErr = ""
                'CODIGO DE PROVINCIA
                If SucProvnro <> 0 Then
                    StrSql = "SELECT provdesc,provcodext FROM provincia WHERE provnro = " & SucProvnro
                    OpenRecordset StrSql, rs2
                    If Not rs2.EOF Then
                        SucCodProv = IIf(EsNulo(rs2!provcodext), Empty, rs2!provcodext)
                        Auxiliar = rs2!provdesc
                    End If
                    
                    If EsNulo(SucCodProv) = True And Auxiliar = "" Then
                        MensajeErr = "No se encontró provincia en el domicilio configurado."
                    ElseIf EsNulo(SucCodProv) = True Then
                        MensajeErr = "No se encontró el código externo de la provincia : " & Auxiliar
                    End If
                Else
                    MensajeErr = "No se encontró provincia en el domicilio configurado."
                End If
                'Flog.writeline Espacios(Tabulador * 2) & "CÓDIGO DE PROVINCIA: " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucCodProv)
                Flog.writeline Format_StrLR("CÓDIGO DE PROVINCIA", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, SucCodProv)
            
            End If
        End If
        
        
        
        '---------------------------------------------------------------------
        'CENTRALIZA | SI/NO
        '---------------------------------------------------------------------
        CodCentraliza = ""
        MensajeErr = ""
        Auxiliar = ""
        StrSql = "SELECT nrocod FROM  estr_cod"
        StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
        StrSql = StrSql & " WHERE estr_cod.tcodnro =  " & TipoCodSucCen
        StrSql = StrSql & " AND estr_cod.Estrnro = " & SucursalEstrnro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            If EsNulo(rs2!nrocod) Then
                CodCentraliza = "NO"
                MensajeErr = "No se encontró tipo de código asociado a la estructura : " & Sucursal
                MensajeErr = MensajeErr & ". Ver configuración del reporte columna 1 valor 3 (Tipo de Código Centraliza)."
                MensajeErr = MensajeErr & "Se informa NO como default."
            Else
                If UCase(rs2!nrocod) = "SI" Or UCase(rs2!nrocod) = "NO" Then
                    CodCentraliza = UCase(rs2!nrocod)
                Else
                    MensajeErr = "Código inválido. Se debe configurar SI/NO."
                    MensajeErr = MensajeErr & "Revisar tipo de código " & TipoCodSucCen & " para la estructura : " & Sucursal
                    MensajeErr = MensajeErr & "Se informa NO como default"
                    CodCentraliza = "NO"
                End If
            End If
        Else
            MensajeErr = "No se encontró tipo de código asociado a la estructura : " & Sucursal
            MensajeErr = MensajeErr & ". Ver configuración del reporte columna 1 valor 3 (Tipo de Código Centraliza)."
            MensajeErr = MensajeErr & "Se informa NO como default"
            CodCentraliza = "NO"
        End If
        'Flog.writeline Espacios(Tabulador * 2) & "CENTRALIZA: " & IIf(Not EsNulo(MensajeErr), MensajeErr, CodCentraliza)
        Flog.writeline Format_StrLR("CENTRALIZA", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, CodCentraliza)
        
        '---------------------------------------------------------------------
        'CUIL
        '---------------------------------------------------------------------
        MensajeErr = ""
        Cuil = ""
        StrSql = " SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc "
        StrSql = StrSql & " WHERE ternro = " & Ternro
        StrSql = StrSql & " AND tidnro = 10 "
        OpenRecordset StrSql, rs2
        If rs2.EOF Then
            MensajeErr = "No se encontró el CUIL para el empleado."
        Else
            If EsNulo(rs2("nrodoc")) = True Then
                MensajeErr = "No se encontró el CUIL para el empleado."
            Else
                Cuil = Left(CStr(rs2("nrodoc")), 13)
                Cuil = Replace(Cuil, "-", "")
            End If
        End If
        'Flog.writeline "CUIL: " & IIf(Not EsNulo(MensajeErr), MensajeErr, Cuil)
        Flog.writeline Format_StrLR("CUIL", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, Cuil)
        

        '---------------------------------------------------------------------
        'Ingreso | FECHA DESDE DE LA FASE MARCADA COMO ALTA RECONOCIDA
        '---------------------------------------------------------------------
        MensajeErr = ""
        StrSql = "SELECT altfec FROM fases "
        StrSql = StrSql & " WHERE  fasrecofec = -1 and empleado = " & Ternro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            FecIngreso = rs2!altfec
        Else
            MensajeErr = "No se encontró fase con fecha de alta reconocida."
            FecIngreso = Date
        End If
        'Flog.writeline "INGRESO: " & IIf(Not EsNulo(MensajeErr), MensajeErr, FecIngreso)
        Flog.writeline Format_StrLR("INGRESO", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, FecIngreso)
        
        
              
        '---------------------------------------------------------------------
        'Categoria | Retorna cod. asociado a la estructura para cada empleado a la fecha seleccionada en el filtro
        '---------------------------------------------------------------------
        Categoria = ""
        CategoriaCod = Empty
        MensajeErr = ""
        'StrSql = " SELECT his_estructura.ternro,estrdabr,his_estructura.estrnro,nrocod"
        StrSql = " SELECT his_estructura.ternro,estrdabr,his_estructura.estrnro"
        StrSql = StrSql & ", (SELECT nrocod FROM estr_cod INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro WHERE estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & TipoCodCat & ") nrocod"
        StrSql = StrSql & " From his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
        'StrSql = StrSql & " LEFT JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & TipoCodCat
        'StrSql = StrSql & " LEFT JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha)
        StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(Fecha) & ")"
        StrSql = StrSql & " AND his_estructura.tenro = " & TenroCategoria
        StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            Categoria = rs2!estrdabr
            If Not EsNulo(rs2("nrocod")) Then
                CategoriaCod = IIf(EsNulo(rs2!nrocod), Empty, rs2!nrocod)
            End If
        Else
            If Pulgas <> "" Then
                Pulgas = Pulgas & "<br>Categoria"
            Else
                Pulgas = "<br>Categoria"
            End If
            MensajeErr = "No se encontraron datos a la fecha : " & Fecha
        End If
        
        If MensajeErr = "" Then
            If IsEmpty(CategoriaCod) Then
                MensajeErr = "No se encontró tipo de código asociado a la estructura : " & Categoria
                MensajeErr = MensajeErr & ". Ver configuración del reporte columna 2 valor 2 (Tipo de Código)"
            Else
                'CONTROLO QUE SEA UN NÚMERO
                If Not IsNumeric(CategoriaCod) Then
                    MensajeErr = "EL código debe ser numérico."
                    MensajeErr = MensajeErr & ". Ver configuración del reporte columna 2 valor 2 (Tipo de Código)"
                End If
            End If
        End If
        'Flog.writeline "CATEGORIA: " & IIf(Not EsNulo(MensajeErr), MensajeErr, CategoriaCod)
        Flog.writeline Format_StrLR("CATEGORIA", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, CategoriaCod)
        
        
        '---------------------------------------------------------------------
        'Remuneraciòn
        '---------------------------------------------------------------------
        MensajeErr = ""
        StrSql = " SELECT almonto"
        StrSql = StrSql & " FROM acu_liq"
        StrSql = StrSql & " WHERE acunro = " & AC_RemuTotal
        StrSql = StrSql & " AND cliqnro = " & cliqnro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            RemuTotal = rs2!almonto
        Else
            RemuTotal = 0
            MensajeErr = "No se encontro valor para el AC: " & AC_RemuTotal
            MensajeErr = MensajeErr & ". Ver configuración del reporte columna 4 valor 1."
        End If
        'Flog.writeline "REMUNERACIÓN: " & IIf(Not EsNulo(MensajeErr), MensajeErr, RemuTotal)
        Flog.writeline Format_StrLR("REMUNERACIÓN", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, RemuTotal)
        
        
        '---------------------------------------------------------------------
        'Media Jornada | S/N  (TE: Tipo de Jornada)
        '---------------------------------------------------------------------
        MensajeErr = ""
        MediaJor = Empty
        'StrSql = " SELECT his_estructura.ternro,estrdabr,his_estructura.estrnro,nrocod"
        StrSql = " SELECT his_estructura.ternro,estrdabr,his_estructura.estrnro"
        StrSql = StrSql & ", (SELECT nrocod FROM estr_cod INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro WHERE estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & TipoCodRegHor & ") nrocod"
        StrSql = StrSql & " From his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
        'StrSql = StrSql & " LEFT JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & TipoCodRegHor
        'StrSql = StrSql & " LEFT JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha)
        StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(Fecha) & ")"
        StrSql = StrSql & " AND his_estructura.tenro = " & TenroRegHor
        StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            If EsNulo(rs2!nrocod) Then
                MediaJor = "N"
                MensajeErr = "No se encontró tipo de código asociado a la estructura : " & rs2!estrdabr
                MensajeErr = MensajeErr & ". Ver configuración del reporte columna 3 valor 2 (Tipo de Código Media jornada)."
                MensajeErr = MensajeErr & "Se informa N como default."
            Else
                If (UCase(rs2!nrocod) = "SI" Or UCase(rs2!nrocod) = "NO") Or UCase(rs2!nrocod) = "S" Or UCase(rs2!nrocod) = "N" Then
                    MediaJor = Left(UCase(rs2!nrocod), 1)
                Else
                    MensajeErr = "Código inválido. Se debe configurar S/N o SI/NO."
                    MensajeErr = MensajeErr & "Revisar tipo de código " & TipoCodRegHor & " para la estructura : " & rs2!estrdabr
                    MensajeErr = MensajeErr & "Se informa N como default"
                    MediaJor = "N"
                End If
            End If
        Else
            
            If Pulgas <> "" Then
                Pulgas = Pulgas & "<br>Regimen Horario"
            Else
                Pulgas = "<br>Regimen Horario"
            End If
            MediaJor = "N"
            MensajeErr = "No se encontró estructura de tipo Regimen horario a la fecha : " & Fecha
            MensajeErr = MensajeErr & ". Ver configuración del reporte columna 3 valor 2 (Tipo de Código Media jornada)."
            MensajeErr = MensajeErr & "Se informa N como default"
        End If
        'Flog.writeline Espacios(Tabulador * 2) & "MEDIA JORNADA: " & IIf(Not EsNulo(MensajeErr), MensajeErr, MediaJor)
        Flog.writeline Format_StrLR("MEDIA JORNADA", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, MediaJor)
        
        '---------------------------------------------------------------------
        'Licencia | S/N (Si tiene al menos 1 en el mes del periodo desde
        '---------------------------------------------------------------------
        Licencia = "N"
        MensajeErr = ""
        If listaTipoLic <> "0" Then
            StrSql = "SELECT empleado FROM emp_lic WHERE (empleado = " & Ternro & " )"
            StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(pliqhasta)
            StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(pliqdesde)
            StrSql = StrSql & " AND tdnro IN (" & listaTipoLic & ")"
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                Licencia = "S"
            Else
                MensajeErr = "No se encontraron licencias entre " & pliqdesde & " y " & pliqhasta & " para los tipos -> " & listaTipoLic
            End If
        Else
            MensajeErr = "No hay licencias configuradas. Ver configuración del reporte columna 5 valor 1. Se informa N por default."
        End If
        'Flog.writeline Espacios(Tabulador * 2) & "LICENCIA: " & IIf(Not EsNulo(MensajeErr), MensajeErr, Licencia)
        Flog.writeline Format_StrLR("LICENCIA", Longitud, "R", True, " ") & ": " & IIf(Not EsNulo(MensajeErr), MensajeErr, Licencia)
                
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        CantRegistros = CantRegistros - 1
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(CantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " INSERT INTO rep_faecys "
        StrSql = StrSql & "(bpronro , Ternro, Empleg, faedescsuc, faecodsuc, faecentral, faecuil"
        StrSql = StrSql & ", faeape, faenom, faeing, faecodcat, faerem, faemedjor, faelic"
        StrSql = StrSql & " ,faecalle,faenum,faecodloc,faecodprov,faeempresa,pliqdesde,faepulgas)"
        StrSql = StrSql & " VALUES"
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & ",'" & Ternro & "'"
        StrSql = StrSql & ",'" & Empleg & "'"
        StrSql = StrSql & ",'" & Sucursal & "'" ' faedescsuc
        StrSql = StrSql & ",'" & SucursalCod & "'" 'faecodsuc
        StrSql = StrSql & ",'" & CodCentraliza & "'" 'CodCentraliza
        StrSql = StrSql & ",'" & Cuil & "'" 'faecuil
        StrSql = StrSql & ",'" & Terape & "'" ' faeape
        StrSql = StrSql & ",'" & Ternom & "'" 'faenom
        StrSql = StrSql & "," & ConvFecha(FecIngreso) 'faeing
        StrSql = StrSql & ",'" & CategoriaCod & "'" 'faecodcat
        StrSql = StrSql & "," & CDbl(RemuTotal) 'faerem
        StrSql = StrSql & ",'" & MediaJor & "'" 'faemedjor
        StrSql = StrSql & ",'" & Licencia & "'" 'faelic
        StrSql = StrSql & ",'" & SucCalle & "'" 'faecalle
        StrSql = StrSql & ",'" & SucNro & "'" 'faenum
        StrSql = StrSql & ",'" & SucCodLoc & "'" 'faecodloc
        StrSql = StrSql & ",'" & SucCodProv & "'" 'faecodprov
        StrSql = StrSql & ",'" & Empnom & "'" 'Nombre empresa
        StrSql = StrSql & "," & ConvFecha(pliqdesde) 'Periodo desde
        StrSql = StrSql & ",'" & Pulgas & "'" 'Pulgas
        
        
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline ""
        Flog.writeline "------------------------------> " & "Registro Insertado"
        rs.MoveNext
    Loop
    Flog.writeline ""
    Flog.writeline "----------------------------------------------------------------------------------------------------"
    Flog.writeline "----------------------------------------------------------------------------------------------------"
End If


Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub
Public Function Format_StrLR(ByVal Str, ByVal Longitud As Long, ByVal Posicion As String, ByVal Completar As Boolean, ByVal Str_Completar As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y lo completa
'              con el caracter pasado por parametro a la izq/der hasta la longitud (si completar es TRUE)
' Autor      : Gonzalez Nicolás
' Fecha      : 19/01/2016
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
If Not EsNulo(Str) Then
    Str = Left(Str, Longitud)
    If Completar Then
        If Len(Str) < Longitud Then
            If UCase(Posicion) = "R" Then
                Str = Str & String(Longitud - Len(Str), Str_Completar)
            Else
                Str = String(Longitud - Len(Str), Str_Completar) & Str
            End If
        End If
    End If
    'Corta el string según Tope
    Format_StrLR = UCase(Str)
Else
    If Completar Then
        Format_StrLR = String(Longitud, " ")
    Else
        Format_StrLR = ""
    End If
End If

End Function
