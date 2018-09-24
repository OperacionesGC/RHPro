Attribute VB_Name = "ExpRenatreCupon"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "08/01/2010" 'Elizabeth Gisela Oviedo
'Global Const UltimaModificacion = "" 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "10/03/2010" 'Elizabeth Gisela Oviedo
'Global Const UltimaModificacion = "" 'Toma un período de 3 meses a partir de la fecha inicial

'Global Const Version = "1.03"
'Global Const FechaModificacion = "22/03/2010" 'Elizabeth Gisela Oviedo
'Global Const UltimaModificacion = "" 'Toma un período de liquidacion 3 meses a partir de la fecha inicial
                                     ' No toma en cuenta la fecha de ingreso del empleado.

'Global Const Version = "1.04"
'Global Const FechaModificacion = "23/03/2010" 'Elizabeth Gisela Oviedo
'Global Const UltimaModificacion = "" 'Cambio de formato de exportación de fecha (01-02-2010 por 01/02/2010), coma final
                                     
'Global Const Version = "1.05"
'Global Const FechaModificacion = "25/03/2010" 'Elizabeth Gisela Oviedo
'Global Const UltimaModificacion = "" 'filtro de empleados seleccionados

'Global Const Version = "1.06"
'Global Const FechaModificacion = "13/04/2010" 'Elizabeth Gisela Oviedo
'Global Const UltimaModificacion = "" 'filtro de empleados seleccionados

Global Const Version = "1.07"
Global Const FechaModificacion = "02/06/2010" 'Elizabeth Gisela Oviedo
Global Const UltimaModificacion = "" 'Eliminacion de filtro por fases en empleados
                                     

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

'NUEVAS
Global EmpErrores As Boolean

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer

Global errorConfrep As Boolean

Global TipoCols4(200)
Global CodCols4(200)
Global TipoCols5(200)
Global CodCols5(200)

Global mes1 As String
Global mesPorc1 As String
Global mes2 As String
Global mesPorc2 As String
Global mes3 As String
Global mes4 As String
Global mesPorc3 As String
Global mes5 As String
Global mesPorc4 As String
Global mes6 As String


Global mesPeriodo As Integer
Global anioPeriodo As Integer
Global mesAnterior1 As Integer
Global mesAnterior2 As Integer
Global anioAnterior1 As Integer
Global anioAnterior2 As Integer

Global cantColumna4
Global cantColumna5

Global estrnomb1
Global estrnomb2
Global estrnomb3
Global testrnomb1
Global testrnomb2
Global testrnomb3

Global tprocNro As Integer
Global tprocDesc As String
Global proDesc As String
Global concnro As Integer
Global Conccod As String
Global concabr As String
Global tconnro As Integer
Global tconDesc As String
Global concimp As Integer
Global concpuente As Integer
Global fecEstr As String
Global Formato As Integer
Global Modelo As Long
Global TituloRep As String
Global descDesde
Global descHasta
Global FechaHasta
Global FechaDesde
Global ArchExp
Global UsaEncabezado As Integer
Global Encabezado As Boolean


Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion
' Autor      : FAF
' Fecha      : 11/04/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim Empresa As Long
Dim Tenro As Long
Dim Estrnro As Long
Dim Fecha As String
Dim informa_fecha As Integer
Dim nroModelo As Integer

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    'Abro la conexion
    
    Nombre_Arch = PathFLog & "ExportacionRenatre" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
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
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
   
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 260"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       Empresa = CLng(ArrParametros(0))
       Fecha = ArrParametros(1)
       Tenro = CLng(ArrParametros(2))
       Estrnro = CLng(ArrParametros(3))
       informa_fecha = CInt(ArrParametros(4))
       nroModelo = CInt(ArrParametros(5))
       Call Generar_Archivo(Empresa, Fecha, Tenro, Estrnro, informa_fecha, nroModelo)
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
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


Private Sub Generar_Archivo(ByVal Empnro As Long, ByVal fecha_alta As String, ByVal tipoestr As Long, ByVal estructura As Long, ByVal informa_fecha_alta As Integer, ByVal nroModelo As Integer)

' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      :
' Fecha      :
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim Directorio As String
Dim Carpeta
Dim fs1



Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim ternro_e As Long

Dim Cuil As String
Dim apynom As String
Dim f_ingreso As String
Dim tarea As String
Dim Trab As String
Dim per1 As String
Dim Dias1 As String
Dim Rem1 As String
Dim per2 As String
Dim Dias2 As String
Dim Rem2 As String
Dim per3 As String
Dim Dias3 As String
Dim Rem3 As String
Dim conv As String
Dim f_cese As String
Dim Cuit As String
Dim razon As String
Dim per_mes_1 As String
Dim per_mes_2 As String
Dim per_mes_3 As String
Dim per_anio_1 As String
Dim per_anio_2 As String
Dim per_anio_3 As String
Dim conf_fuera_conv
Dim conf_acum_dias
Dim conf_acum_rem
Dim NroReporte
Dim tipCodRenatre
Dim linea
Dim separador As String
Dim ternroempr
Dim fecha_limite As String

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    On Error GoTo ME_Local
      
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        'Directorio = Trim(rs!sis_dirsalidas) & "\ExpRenatre"
        Directorio = Trim(rs!sis_dirsalidas)
    End If
    
     
    StrSql = "SELECT * FROM modelo WHERE modnro = " & nroModelo
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Directorio & Trim(rs!modarchdefault)
        separador = IIf(Not IsNull(rs!modseparador), rs!modseparador, ",")
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & nroModelo & " " & rs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de importación :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & nroModelo
        Exit Sub
    End If
    
    Nombre_Arch = Directorio & "\renatre.csv"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    'desactivo el manejador de errores
    On Error Resume Next
    
    Set Carpeta = fs.GetFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & " no existe. Se creará."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio " & Directorio & ". Verifique los derechos de acceso o puede crearlo."
            HuboErrores = True
            GoTo Fin
        End If
    End If
    
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
    
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio
        HuboErrores = True
        GoTo Fin
    End If
    
    On Error GoTo ME_Local
    
    'Configuracion del Reporte
    NroReporte = 272
   
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND conftipo= 'CON'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Reporte"
        Exit Sub
    Else
        conf_fuera_conv = rs!confval
    End If
    
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND conftipo= 'REM'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        conf_acum_rem = rs!confval
    End If
    
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND conftipo= 'DIA'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        conf_acum_dias = rs!confval
    End If
    
    '------------------------------------------------------------------
    'Codigo RENATRE
    '------------------------------------------------------------------
    tipCodRenatre = 0
    StrSql = "SELECT tcodnro FROM tipocod "
    StrSql = StrSql & " WHERE tcodnom = 'RENATRE'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Codigo para renatre"
        Exit Sub
    Else
        tipCodRenatre = rs!tcodnro
    End If
    rs.Close

    per_mes_1 = Month(fecha_alta)
    per_anio_1 = Year(fecha_alta)
    

    If per_mes_1 = 12 Then
        per_anio_2 = per_anio_1 + 1
        per_mes_2 = 1
    Else
        per_anio_2 = per_anio_1
        per_mes_2 = per_mes_1 + 1
    End If
        
    If per_mes_2 = 12 Then
        per_anio_3 = per_anio_2 + 1
        per_mes_3 = 1
    Else
        per_anio_3 = per_anio_2
        per_mes_3 = per_mes_2 + 1
    End If
        
    per1 = Format(per_mes_1, "00") & "/" & Format(per_anio_1, "0000")
    per2 = Format(per_mes_2, "00") & "/" & Format(per_anio_2, "0000")
    per3 = Format(per_mes_3, "00") & "/" & Format(per_anio_3, "0000")
        
    
    fecha_limite = DateAdd("d", 90, fecha_alta)

    '------------------------------------------------------------------
    'Busco los datos
    '------------------------------------------------------------------
    StrSql = " SELECT distinct e.ternro, e.terape, e.ternom, e.empleg  "
    StrSql = StrSql & " FROM empleado e "
    StrSql = StrSql & " INNER JOIN cabliq c  on (c.empleado=e.ternro)"
    StrSql = StrSql & " INNER JOIN proceso  pr on (c.pronro=pr.pronro) "
    StrSql = StrSql & " INNER JOIN periodo  pe on (pe.pliqnro=pr.pliqnro) "
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON e.ternro = empresa.ternro "
    StrSql = StrSql & " AND empresa.tenro = 10 AND empresa.htetdesde <= " & ConvFecha(fecha_alta)
    StrSql = StrSql & " AND (empresa.htethasta >= " & ConvFecha(fecha_alta)
    StrSql = StrSql & " OR empresa.htethasta IS NULL) AND empresa.estrnro = " & Empnro
    If tipoestr <> 0 Or estructura <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura ON e.ternro = his_estructura.ternro "
        StrSql = StrSql & " AND his_estructura.tenro = " & tipoestr
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta)
        StrSql = StrSql & "      OR his_estructura.htethasta IS NULL) "
        If estructura <> 0 Then
            StrSql = StrSql & " AND his_estructura.estrnro = " & estructura
        End If
    End If
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " ((pe.pliqanio= " & per_anio_1 & " and pe.pliqmes=" & per_mes_1 & ") or"
    StrSql = StrSql & "  (pe.pliqanio= " & per_anio_2 & " and pe.pliqmes=" & per_mes_2 & ") or"
    StrSql = StrSql & "  (pe.pliqanio= " & per_anio_3 & " and pe.pliqmes=" & per_mes_3 & "))"
    
    
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        ternro_e = IIf(IsNull(rs!ternro), "", rs!ternro)
        apynom = IIf(IsNull(rs!terape), "", rs!terape & " " & IIf(IsNull(rs!ternom), "", rs!ternom))
        Legajo = IIf(IsNull(rs!empleg), "", rs!empleg)

        
        ' CUIL
        StrSql = " SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc "
        StrSql = StrSql & " WHERE ter_doc.ternro = " & ternro_e
        StrSql = StrSql & " AND ter_doc.tidnro = 10"
        OpenRecordset StrSql, rs2
        Cuil = ""
        If Not rs2.EOF And Not IsNull(rs2!Nrodoc) Then
            Cuil = Replace(rs2!Nrodoc, "-", "")
            Cuil = Replace(Cuil, ",", "")
            Cuil = Replace(Cuil, ".", "")
        End If
        rs2.Close
        
        
        ' FECHA INGRESO
        StrSql = " SELECT altfec "
        StrSql = StrSql & " FROM fases "
        StrSql = StrSql & " WHERE Empleado = " & rs!ternro
        StrSql = StrSql & "      AND altfec is not null and altfec<>'' "
        StrSql = StrSql & " order by fasrecofec,altfec desc "
        OpenRecordset StrSql, rs2
        f_ingreso = ""
        If Not rs2.EOF And Not IsNull(rs2!altfec) Then
            f_ingreso = Format(Day(rs2!altfec), "00") & "/" & Format(Month(rs2!altfec), "00") & "/" & Format(Year(rs2!altfec), "0000")
        End If
        rs2.Close
        
        ' FECHA cese
        StrSql = " SELECT bajfec "
        StrSql = StrSql & " FROM fases "
        StrSql = StrSql & " WHERE Empleado = " & rs!ternro
        StrSql = StrSql & "      AND altfec is not null and altfec<>'' "
        StrSql = StrSql & " order by altfec desc "
      
        OpenRecordset StrSql, rs2
        f_cese = ""
        If Not rs2.EOF And Not IsNull(rs2!bajfec) Then
            If rs2!bajfec <> "" And Not IsNull(rs2!bajfec) Then
                f_cese = Format(Day(rs2!bajfec), "00") & "/" & Format(Month(rs2!bajfec), "00") & "/" & Format(Year(rs2!bajfec), "0000")
            End If
        End If
        rs2.Close
       
        
        ' tarea realizada
        StrSql = " SELECT nrocod "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 3"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        tarea = ""
        If Not rs2.EOF Then
            If IsNull(rs2!nrocod) Then
                tarea = ""
            Else
                tarea = CStr(rs2!nrocod)
            End If
        End If
        rs2.Close
    
        
        ' Tipo Contrato
        StrSql = " SELECT nrocod "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 18"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        
        Trab = ""
        If Not rs2.EOF Then
            If IsNull(rs2!nrocod) Then
                Trab = ""
            Else
                Trab = CStr(rs2!nrocod)
            End If
        End If
        rs2.Close
       

        ' Período liquidacion
'        StrSql = " SELECT pe.pliqanio, pe.pliqmes "
'        StrSql = StrSql & " FROM cabliq c "
'        StrSql = StrSql & " INNER JOIN proceso  pr on (c.pronro=pr.pronro) "
'        StrSql = StrSql & " inner join periodo  pe on (pe.pliqnro=pr.pliqnro) "
'        StrSql = StrSql & " WHERE c.Empleado = " & rs!ternro
'        StrSql = StrSql & " group by pe.pliqanio, pe.pliqmes "
'        StrSql = StrSql & " order by pliqanio desc, pliqmes desc "
'
'        OpenRecordset StrSql, rs2
'        per_mes_1 = ""
'        per_mes_2 = ""
'        per_mes_3 = ""
'        per_anio_1 = ""
'        per_anio_2 = ""
'        per_anio_1 = ""
'        Do Until (rs2.EOF Or (per_mes_3 <> ""))
'            If Not (IsNull(rs2!pliqanio)) And Not (IsNull(rs2!pliqmes)) Then
'                If per_mes_1 = "" Then
'                    per_mes_1 = Format(CStr(rs2!pliqmes), "00")
'                    per_anio_1 = Format(CStr(rs2!pliqanio), "0000")
'                Else
'                    If per_mes_2 = "" Then
'                        per_mes_2 = Format(CStr(rs2!pliqmes), "00")
'                        per_anio_2 = Format(CStr(rs2!pliqanio), "0000")
'                    Else
'                        If per_mes_3 = "" Then
'                            per_mes_3 = Format(CStr(rs2!pliqmes), "00")
'                            per_anio_3 = Format(CStr(rs2!pliqanio), "0000")
'                        End If
'                    End If
'                End If
'            End If
'            rs2.MoveNext
'        Loop
'        rs2.Close
'
'        per1 = Format(per_mes_1, "00") & "/" & Format(per_anio_1, "0000")
'        per2 = Format(per_mes_2, "00") & "/" & Format(per_anio_2, "0000")
'        per3 = Format(per_mes_3, "00") & "/" & Format(per_anio_3, "0000")
        
        
        ' dias
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_1
        StrSql = StrSql & "   and ammes  = " & per_mes_1
        StrSql = StrSql & "   and acunro = " & conf_acum_dias
        OpenRecordset StrSql, rs2
        
        Dias1 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Dias1 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close
        

        ' dias
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_2
        StrSql = StrSql & "   and ammes  = " & per_mes_2
        StrSql = StrSql & "   and acunro = " & conf_acum_dias
        OpenRecordset StrSql, rs2
        
        Dias2 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Dias2 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close

        ' dias
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_3
        StrSql = StrSql & "   and ammes  = " & per_mes_3
        StrSql = StrSql & "   and acunro = " & conf_acum_dias
        OpenRecordset StrSql, rs2
        
        Dias3 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Dias3 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close


        ' Remuneración
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_1
        StrSql = StrSql & "   and ammes  = " & per_mes_1
        StrSql = StrSql & "   and acunro = " & conf_acum_rem
        OpenRecordset StrSql, rs2
        
        Rem1 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Rem1 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close

        ' Remuneración
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_2
        StrSql = StrSql & "   and ammes  = " & per_mes_2
        StrSql = StrSql & "   and acunro = " & conf_acum_rem
        OpenRecordset StrSql, rs2
        
        Rem2 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Rem2 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close
        
        ' Remuneración
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_3
        StrSql = StrSql & "   and ammes  = " & per_mes_3
        StrSql = StrSql & "   and acunro = " & conf_acum_rem
        OpenRecordset StrSql, rs2
        
        Rem3 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Rem3 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close

        
        ' Convenio
        StrSql = " SELECT estr_cod.estrnro "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro
        StrSql = StrSql & " AND his_estructura.tenro = 19"
        StrSql = StrSql & " AND his_estructura.estrnro = " & conf_fuera_conv
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
                
        OpenRecordset StrSql, rs2
        conv = "NO"
        If Not rs2.EOF Then
            If IsNull(rs2!Estrnro) Then
                conv = "NO"
            Else
                conv = "SI"
            End If
        End If
        rs2.Close
        
        ' Empresa
        StrSql = " SELECT empnom,ternro  "
        StrSql = StrSql & " From Empresa "
        StrSql = StrSql & " Where Estrnro = " & Empnro
        OpenRecordset StrSql, rs2
        razon = ""
        If Not rs2.EOF And Not IsNull(rs2!empnom) Then
            razon = rs2!empnom
            ternroempr = rs2!ternro
        End If
        rs2.Close
        
        ' CUIT
        StrSql = " SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc "
        StrSql = StrSql & " WHERE ter_doc.ternro = " & ternroempr
        StrSql = StrSql & " AND ter_doc.tidnro = 6"
        OpenRecordset StrSql, rs2
        Cuit = ""
        If Not rs2.EOF And Not IsNull(rs2!Nrodoc) Then
            Cuit = Replace(rs2!Nrodoc, "-", "")
            Cuit = Replace(Cuit, ",", "")
            Cuit = Replace(Cuit, ".", "")
        End If
        rs2.Close
            
            
        'RENGLON

        'Cuil , apynom_trabajador, fecha_de_ingreso,tarea_realizada,Trab
        'periodo1,Dias_periodo1,Rem_suj_aportes_1,
        'periodo2,Dias_periodo2,Rem_suj_aportes_2,
        'periodo3,Dias_periodo3,Rem_suj_aportes_3,
        'conv_corresp,fecha_de_cese,Cuit,razon_social,Legajo
        
        
        linea = """" & Cuil & """" & separador & """" & _
                apynom & """" & separador & """" & _
                f_ingreso & """" & separador & """" & _
                tarea & """" & separador & """" & _
                Trab & """" & separador & """" & _
                per1 & """" & separador & """" & _
                Dias1 & """" & separador & """" & _
                Rem1 & """" & separador & """" & _
                per2 & """" & separador & """" & _
                Dias2 & """" & separador & """" & _
                Rem2 & """" & separador & """" & _
                per3 & """" & separador & """" & _
                Dias3 & """" & separador & """" & _
                Rem3 & """" & separador & """" & _
                conv & """" & separador & """" & _
                f_cese & """" & separador & """" & _
                Cuit & """" & separador & """" & _
                razon & """" & separador & """" & _
                Legajo & """" & separador


                
        ArchExp.Write linea
        ArchExp.writeline
        
        rs.MoveNext
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    Loop
    
            
    ArchExp.Close
    
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
  

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub



Private Sub Generar_Archivo_Old(ByVal Empnro As Long, ByVal fecha_alta As String, ByVal tipoestr As Long, ByVal estructura As Long, ByVal informa_fecha_alta As Integer, ByVal nroModelo As Integer)

' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      :
' Fecha      :
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim Directorio As String
Dim Carpeta
Dim fs1



Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim ternro_e As Long

Dim Cuil As String
Dim apynom As String
Dim f_ingreso As String
Dim tarea As String
Dim Trab As String
Dim per1 As String
Dim Dias1 As String
Dim Rem1 As String
Dim per2 As String
Dim Dias2 As String
Dim Rem2 As String
Dim per3 As String
Dim Dias3 As String
Dim Rem3 As String
Dim conv As String
Dim f_cese As String
Dim Cuit As String
Dim razon As String
Dim per_mes_1 As String
Dim per_mes_2 As String
Dim per_mes_3 As String
Dim per_anio_1 As String
Dim per_anio_2 As String
Dim per_anio_3 As String
Dim conf_fuera_conv
Dim conf_acum_dias
Dim conf_acum_rem
Dim NroReporte
Dim tipCodRenatre
Dim linea
Dim separador As String
Dim fecha_limite As String


Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    On Error GoTo ME_Local
      
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        'Directorio = Trim(rs!sis_dirsalidas) & "\ExpRenatre"
        Directorio = Trim(rs!sis_dirsalidas)
    End If
    
     
    StrSql = "SELECT * FROM modelo WHERE modnro = " & nroModelo
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Directorio & Trim(rs!modarchdefault)
        separador = IIf(Not IsNull(rs!modseparador), rs!modseparador, ",")
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & nroModelo & " " & rs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de importación :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & nroModelo
        Exit Sub
    End If
    
    Nombre_Arch = Directorio & "\renatre.txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    'desactivo el manejador de errores
    On Error Resume Next
    
    Set Carpeta = fs.GetFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & " no existe. Se creará."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio " & Directorio & ". Verifique los derechos de acceso o puede crearlo."
            HuboErrores = True
            GoTo Fin
        End If
    End If
    
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
    
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio
        HuboErrores = True
        GoTo Fin
    End If
    
    On Error GoTo ME_Local
    
    'Configuracion del Reporte
    NroReporte = 272
   
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND conftipo= 'CON'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Reporte"
        Exit Sub
    Else
        conf_fuera_conv = rs!confval
    End If
    
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND conftipo= 'REM'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        conf_acum_rem = rs!confval
    End If
    
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND conftipo= 'DIA'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        conf_acum_dias = rs!confval
    End If
    
    '------------------------------------------------------------------
    'Codigo RENATRE
    '------------------------------------------------------------------
    tipCodRenatre = 0
    StrSql = "SELECT tcodnro FROM tipocod "
    StrSql = StrSql & " WHERE tcodnom = 'RENATRE'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Codigo para renatre"
        Exit Sub
    Else
        tipCodRenatre = rs!tcodnro
    End If
    rs.Close

    fecha_limite = DateAdd("d", 90, fecha_alta)

    '------------------------------------------------------------------
    'Busco los datos
    '------------------------------------------------------------------
    StrSql = " SELECT distinct empleado.ternro, empleado.terape, empleado.ternom, empleado.empleg  "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro "
    StrSql = StrSql & " AND empresa.tenro = 10 AND empresa.htetdesde <= " & ConvFecha(fecha_alta)
    StrSql = StrSql & " AND (empresa.htethasta >= " & ConvFecha(fecha_alta)
    StrSql = StrSql & " OR empresa.htethasta IS NULL) AND empresa.estrnro = " & Empnro
    If tipoestr <> 0 Or estructura <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
        StrSql = StrSql & " AND his_estructura.tenro = " & tipoestr
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta)
        StrSql = StrSql & "      OR his_estructura.htethasta IS NULL) "
        If estructura <> 0 Then
            StrSql = StrSql & " AND his_estructura.estrnro = " & estructura
        End If
    End If
    StrSql = StrSql & " WHERE empleado.empest = -1"
    If CInt(informa_fecha_alta) = -1 Then
        StrSql = StrSql & " AND empleado.empfaltagr >= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND empleado.empfaltagr <= " & ConvFecha(fecha_limite)
    End If
    
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        ternro_e = IIf(IsNull(rs!ternro), "", rs!ternro)
        apynom = IIf(IsNull(rs!terape), "", rs!terape & " " & IIf(IsNull(rs!ternom), "", rs!ternom))
        Legajo = IIf(IsNull(rs!empleg), "", rs!empleg)

        
        ' CUIL
        StrSql = " SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc "
        StrSql = StrSql & " WHERE ter_doc.ternro = " & ternro_e
        StrSql = StrSql & " AND ter_doc.tidnro = 10"
        OpenRecordset StrSql, rs2
        Cuil = ""
        If Not rs2.EOF And Not IsNull(rs2!Nrodoc) Then
            Cuil = Replace(rs2!Nrodoc, "-", "")
            Cuil = Replace(Cuil, ",", "")
            Cuil = Replace(Cuil, ".", "")
        End If
        rs2.Close
        
        ' CUIT
        StrSql = " SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc "
        StrSql = StrSql & " WHERE ter_doc.ternro = " & rs!ternro
        StrSql = StrSql & " AND ter_doc.tidnro = 6"
        OpenRecordset StrSql, rs2
        Cuit = ""
        If Not rs2.EOF And Not IsNull(rs2!Nrodoc) Then
            Cuit = Replace(rs2!Nrodoc, "-", "")
            Cuit = Replace(Cuit, ",", "")
            Cuit = Replace(Cuit, ".", "")
        End If
        rs2.Close
        
        ' FECHA INGRESO
        StrSql = " SELECT altfec "
        StrSql = StrSql & " FROM fases "
        StrSql = StrSql & " WHERE Empleado = " & rs!ternro
        StrSql = StrSql & "      AND altfec is not null and altfec<>'' "
        StrSql = StrSql & " order by fasrecofec,altfec desc "
        OpenRecordset StrSql, rs2
        f_ingreso = ""
        If Not rs2.EOF And Not IsNull(rs2!altfec) Then
            f_ingreso = Format(Day(rs2!altfec), "00") & Format(Month(rs2!altfec), "00") & Format(Year(rs2!altfec), "0000")
        End If
        rs2.Close
        
        ' FECHA cese
        StrSql = " SELECT bajfec "
        StrSql = StrSql & " FROM fases "
        StrSql = StrSql & " WHERE Empleado = " & rs!ternro
        StrSql = StrSql & "      AND altfec is not null and altfec<>'' "
        StrSql = StrSql & " order by altfec desc "
      
        OpenRecordset StrSql, rs2
        f_cese = ""
        If Not rs2.EOF And Not IsNull(rs2!bajfec) Then
            If rs2!bajfec <> "" And Not IsNull(rs2!bajfec) Then
                f_cese = Format(Day(rs2!bajfec), "00") & Format(Month(rs2!bajfec), "00") & Format(Year(rs2!bajfec), "0000")
            End If
        End If
        rs2.Close
       
        
        ' tarea realizada
        StrSql = " SELECT nrocod "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 3"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        tarea = ""
        If Not rs2.EOF Then
            If IsNull(rs2!nrocod) Then
                tarea = ""
            Else
                tarea = CStr(rs2!nrocod)
            End If
        End If
        rs2.Close
    
        
        ' Tipo Contrato
        StrSql = " SELECT nrocod "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 18"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        
        Trab = ""
        If Not rs2.EOF Then
            If IsNull(rs2!nrocod) Then
                Trab = ""
            Else
                Trab = CStr(rs2!nrocod)
            End If
        End If
        rs2.Close
       

        ' Período liquidacion
        StrSql = " SELECT pe.pliqanio, pe.pliqmes "
        StrSql = StrSql & " FROM cabliq c "
        StrSql = StrSql & " INNER JOIN proceso  pr on (c.pronro=pr.pronro) "
        StrSql = StrSql & " inner join periodo  pe on (pe.pliqnro=pr.pliqnro) "
        StrSql = StrSql & " WHERE c.Empleado = " & rs!ternro
        StrSql = StrSql & " group by pe.pliqanio, pe.pliqmes "
        StrSql = StrSql & " order by pliqanio desc, pliqmes desc "
       
        OpenRecordset StrSql, rs2
        per_mes_1 = ""
        per_mes_2 = ""
        per_mes_3 = ""
        per_anio_1 = ""
        per_anio_2 = ""
        per_anio_1 = ""
        Do Until (rs2.EOF Or (per_mes_3 <> ""))
            If Not (IsNull(rs2!pliqanio)) And Not (IsNull(rs2!pliqmes)) Then
                If per_mes_1 = "" Then
                    per_mes_1 = Format(CStr(rs2!pliqmes), "00")
                    per_anio_1 = Format(CStr(rs2!pliqanio), "0000")
                Else
                    If per_mes_2 = "" Then
                        per_mes_2 = Format(CStr(rs2!pliqmes), "00")
                        per_anio_2 = Format(CStr(rs2!pliqanio), "0000")
                    Else
                        If per_mes_3 = "" Then
                            per_mes_3 = Format(CStr(rs2!pliqmes), "00")
                            per_anio_3 = Format(CStr(rs2!pliqanio), "0000")
                        End If
                    End If
                End If
            End If
            rs2.MoveNext
        Loop
        rs2.Close
        
        per1 = Format(per_mes_1, "00") & "/" & Format(per_anio_1, "0000")
        per2 = Format(per_mes_2, "00") & "/" & Format(per_anio_2, "0000")
        per3 = Format(per_mes_3, "00") & "/" & Format(per_anio_3, "0000")
        
        
        ' dias
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_1
        StrSql = StrSql & "   and ammes  = " & per_mes_1
        StrSql = StrSql & "   and acunro = " & conf_acum_dias
        OpenRecordset StrSql, rs2
        
        Dias1 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Dias1 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close
        

        ' dias
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_2
        StrSql = StrSql & "   and ammes  = " & per_mes_2
        StrSql = StrSql & "   and acunro = " & conf_acum_dias
        OpenRecordset StrSql, rs2
        
        Dias2 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Dias2 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close

        ' dias
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_3
        StrSql = StrSql & "   and ammes  = " & per_mes_3
        StrSql = StrSql & "   and acunro = " & conf_acum_dias
        OpenRecordset StrSql, rs2
        
        Dias3 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Dias3 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close


        ' Remuneración
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_1
        StrSql = StrSql & "   and ammes  = " & per_mes_1
        StrSql = StrSql & "   and acunro = " & conf_acum_rem
        OpenRecordset StrSql, rs2
        
        Rem1 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Rem1 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close

        ' Remuneración
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_2
        StrSql = StrSql & "   and ammes  = " & per_mes_2
        StrSql = StrSql & "   and acunro = " & conf_acum_rem
        OpenRecordset StrSql, rs2
        
        Rem2 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Rem2 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close
        
        ' Remuneración
        StrSql = " SELECT ammonto "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ternro = " & rs!ternro
        StrSql = StrSql & "   and amanio = " & per_anio_3
        StrSql = StrSql & "   and ammes  = " & per_mes_3
        StrSql = StrSql & "   and acunro = " & conf_acum_rem
        OpenRecordset StrSql, rs2
        
        Rem3 = ""
        If Not rs2.EOF Then
            If Not IsNull(rs2!ammonto) Then
                Rem3 = CStr(rs2!ammonto)
            End If
        End If
        rs2.Close

        
        ' Convenio
        StrSql = " SELECT estr_cod.estrnro "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro
        StrSql = StrSql & " AND his_estructura.tenro = 19"
        StrSql = StrSql & " AND his_estructura.estrnro = " & conf_fuera_conv
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
                
        OpenRecordset StrSql, rs2
        conv = "NO"
        If Not rs2.EOF Then
            If IsNull(rs2!Estrnro) Then
                conv = "NO"
            Else
                conv = "SI"
            End If
        End If
        rs2.Close
        
        ' Empresa
        StrSql = " SELECT empnom "
        StrSql = StrSql & " From Empresa "
        StrSql = StrSql & " Where Estrnro = " & Empnro
        OpenRecordset StrSql, rs2
        razon = ""
        If Not rs2.EOF And Not IsNull(rs2!empnom) Then
            razon = rs2!empnom
        End If
        rs2.Close
        
            
        'RENGLON

        'Cuil , apynom_trabajador, fecha_de_ingreso,tarea_realizada,Trab
        'periodo1,Dias_periodo1,Rem_suj_aportes_1,
        'periodo2,Dias_periodo2,Rem_suj_aportes_2,
        'periodo3,Dias_periodo3,Rem_suj_aportes_3,
        'conv_corresp,fecha_de_cese,Cuit,razon_social,Legajo
        
        
        linea = """" & Cuil & """" & separador & """" & _
                apynom & """" & separador & """" & _
                f_ingreso & """" & separador & """" & _
                tarea & """" & separador & """" & _
                Trab & """" & separador & """" & _
                per1 & """" & separador & """" & _
                Dias1 & """" & separador & """" & _
                Rem1 & """" & separador & """" & _
                per2 & """" & separador & """" & _
                Dias2 & """" & separador & """" & _
                Rem2 & """" & separador & """" & _
                per3 & """" & separador & """" & _
                Dias3 & """" & separador & """" & _
                Rem3 & """" & separador & """" & _
                conv & """" & separador & """" & _
                f_cese & """" & separador & """" & _
                Cuit & """" & separador & """" & _
                razon & """" & separador & """" & _
                Legajo & """"


                
        ArchExp.Write linea
        ArchExp.writeline
        
        rs.MoveNext
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    Loop
    
            
    ArchExp.Close
    
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
  

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub





