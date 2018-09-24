Attribute VB_Name = "MdlExportacion"
Option Explicit

''Version no liberada aun, se hará en el proximo caso
'Global Const Version = "1.00"
'Global Const FechaVersion = "16/05/2014"    ' EAM- Se creo el reporte de Headcount
'Global Const UltimaModificacion = " "
'Global Const UltimaModificacion1 = " "

'Version no liberada aun, se hará en el proximo caso
'Global Const Version = "1.0.1"
'Global Const FechaVersion = "11/08/2014"    ' EAM- Se agrego la opcion de agrupar por lineas las mismas aperturas de un empleado
'Global Const UltimaModificacion = " "
'Global Const UltimaModificacion1 = " "


'Global Const Version = "1.0.2"
'Global Const FechaVersion = "21/07/2014"    ' FGZ- ARREGLOS - 22372
'Global Const UltimaModificacion = " "
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.0.3"
'Global Const FechaVersion = "06/08/2014"    ' LED - ARREGLOS - CAS-22372 - SGS - Nuevo Reporte Headcount [Entrega 4]
'Global Const UltimaModificacion = " "
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.0.4"
'Global Const FechaVersion = "06/05/2015"    ' LED - ARREGLOS - CAS-22372 - SGS - Nuevo Reporte Headcount [Entrega 4] control por nulos
                                            ' Miriam Ruiz - CAS-30761 - SGS - Error en reporte Headcount - se corrigió división por cero
'Global Const UltimaModificacion = " "
'Global Const UltimaModificacion1 = " "

Global Const Version = "1.0.5"
Global Const FechaVersion = "20/07/2015"    ' Carmen Quintero - CAS-30759 - SGS - CUSTOM REPORTE HEADCOUNT - Se cambio condicion en la consulta principal
Global Const UltimaModificacion = " "
Global Const UltimaModificacion1 = " "



'---------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------
Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser As String
Global Fecha As Date
Global Hora As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa As New ADODB.Recordset
Global rs_tipocod As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global strsql2 As String

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 07/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
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

    Nombre_Arch = PathFLog & "rep_Headcount" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PID = " & PID
    
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
    
    Version_Valida = ValidarV(Version, 414, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Fin
    End If
    
    On Error GoTo ME_Main

    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Archivo a Generar = " & Nombre_Arch
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 414 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
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
    GoTo Fin:
End Sub

Public Sub Generacion(ByVal bpronro As Long, ByVal Vol_Cod As String, ByVal Separador As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Detalle del Asiento Contable
' Autor      : FGZ
' Fecha      : 16/12/2004
' Ult. Mod   : 24/08/2006 - Se agrego desc del modelo
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim NroLiq As Long
Dim strLinea As String
Dim Aux_Linea As String
Dim Aux_sql As String
Dim Texto As String
Dim listCC As String
Dim empleg As Long
Dim Monto As Double

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Detalles As New ADODB.Recordset
Dim rs_Monto As New ADODB.Recordset
Dim rs_detliq As New ADODB.Recordset
Dim rs_fase As New ADODB.Recordset
Dim rs_Convenio As New ADODB.Recordset
Dim rs_DetalleAux As New ADODB.Recordset


Dim CantidadAux As Double
Dim headcountID
Dim i As Long
Dim cencosAux As String
Dim ubicaAux As String
Dim d_gerenciaAux As String
Dim porcentajeAux As Double
Dim sinUser As String

'EAM- (v1.0) - Busca el directorio de exportacion del sistema
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    directorio = Trim(rs!sis_dirsalidas)
End If

'EAM- (v1.0) - Debe levantar la configuracion del reporte .
'FGZ - 18/07/2014 --------------------------
StrSql = "SELECT confval2 FROM confrep where repnro= 438 and upper(conftipo)= 'CO'"
OpenRecordset StrSql, rs_Modelo
listCC = "0"

Do While Not rs_Modelo.EOF
    'listCC = listCC & ", " & rs_Modelo!ConcNro
    listCC = listCC & ", " & rs_Modelo!confval2
    rs_Modelo.MoveNext
Loop



StrSql = "SELECT * FROM modelo WHERE modnro = 234"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        'busco si guardo el archivo por usuario o no
        sinUser = ""
        StrSql = "SELECT confval2 FROM confrep where repnro= 438 and upper(conftipo)= 'SEG'"
        OpenRecordset StrSql, rs_Detalles
        If Not rs_Detalles.EOF Then
            If UCase(rs_Detalles!confval2) = "SI" Then
                directorio = ValidarRuta(directorio, "\PorUsr", 1)
                directorio = ValidarRuta(directorio, "\" & IdUser, 1)
            Else
                sinUser = "sin_user"
            End If
        End If
        directorio = ValidarRuta(directorio, "\" & Trim(rs_Modelo!modarchdefault), 1)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If

    
''EAM- (v1.0) - Busco los periodos de liquidacion del proceso de volcado del asiento
'StrSql = "SELECT * FROM proc_vol " & _
'        " INNER JOIN periodo ON proc_vol.pliqnro = periodo.pliqnro " & _
'        " WHERE vol_cod = " & Vol_Cod
'OpenRecordset StrSql, rs_Periodo
'If rs_Periodo.EOF Then
'    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron periodos de Liquidacion ene el Proceso de volcado"
'    Exit Sub
'Else
'     NroLiq = rs_Periodo!pliqnro
'End If


'EAM- (v1.0) - Genera el nombre del archivo Excel
Archivo = directorio & "\Headcount. Proceso de volcado - " & CStr(Vol_Cod) & "-" & NroProcesoBatch & ".csv"
'FGZ - 01/03/2013 ------------------------------------------------


Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
Else
   Flog.writeline Espacios(Tabulador * 1) & "Se creo el archivo en destino."
 End If

'desactivo el manejador de errores
On Error GoTo 0


' Comienzo la transaccion
MyBeginTrans
    
    'EAM- (v1.01) - Inserto la cabecera del reporte
    StrSql = "INSERT INTO cab_headcount (desabr,iduser) values ('Headcount. Proceso de volcado - " & CStr(Vol_Cod) & "-" & NroProcesoBatch & ".csv',"
    If sinUser = "sin_user" Then
        StrSql = StrSql & "'" & sinUser & "')"
    Else
        StrSql = StrSql & "'" & IdUser & "')"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    headcountID = getLastIdentity(objConn, "cab_headcount")
    
    'EAM- (v1.01) - Remplazo los !! por la coma
    Vol_Cod = Replace(Vol_Cod, "!!", ",")



'EAM- (v1.0) - Busca el detalle del asiento en el proceso de volcado
StrSql = "SELECT SUM(dlmonto) montoCC,detalle_asi.masinro,detalle_asi.ternro,detalle_asi.empleg,detalle_asi.terape,detalle_asi.dlcosto1,detalle_asi.dlcosto2,detalle_asi.dlcosto3, mod_asiento.masidesc, mod_asiento.masidesc,detalle_asi.masinro, detalle_asi.ternro, " & _
        " e1.estrcodext estrcodext1,e1.estrdabr estrdabr1,e2.estrcodext estrcodext2, e2.estrdabr estrdabr2,e3.estrdabr estrdabr3,proceso.tprocnro,proceso.pronro,proceso.profecfin,emp.empest,ter.tersex " & _
        " FROM  detalle_asi " & _
        " INNER JOIN mod_asiento ON mod_asiento.masinro = detalle_asi.masinro " & _
        " INNER JOIN proceso ON  detalle_asi.dlcosto4 = proceso.pronro " & _
        " INNER JOIN empleado emp ON emp.ternro = detalle_asi.ternro " & _
        " INNER JOIN tercero ter ON emp.ternro = ter.ternro " & _
        " LEFT JOIN estructura e1 on e1.estrnro = detalle_asi.dlcosto1" & _
        " LEFT JOIN estructura e2 on e2.estrnro = detalle_asi.dlcosto2 " & _
        " LEFT JOIN estructura e3 on e3.estrnro = detalle_asi.dlcosto3" & _
        " WHERE detalle_asi.vol_cod IN(" & Vol_Cod & ") AND linaD_H=-1 " & _
        " GROUP BY detalle_asi.dlcosto1,detalle_asi.dlcosto2,detalle_asi.dlcosto3,mod_asiento.masidesc,detalle_asi.masinro, detalle_asi.ternro, detalle_asi.empleg,detalle_asi.terape, e1.estrcodext,e1.estrdabr,e2.estrcodext " & _
        " ,e2.estrdabr,e3.estrdabr,  proceso.tprocnro,proceso.pronro,proceso.profecfin,emp.empest,ter.tersex " & _
        " ORDER BY detalle_asi.masinro, detalle_asi.ternro "
OpenRecordset StrSql, rs_Detalles



'EAM- (v1.0) - Seteo las variables de progreso y calculo el procentaje
Progreso = 0
CConceptosAProc = rs_Detalles.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay Detalles para ese Proceso de Volcado " & Vol_Cod
End If
IncPorc = (100 / CConceptosAProc)

'Procesamiento
If rs_Detalles.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No se encontraron datos en el asiento. Proceo de volcado: " & Vol_Cod
End If

'    Flog.writeline Espacios(Tabulador * 2) & "Registros a Exportar: " & CConceptosAProc
'    'Genero los encabezados
'    Aux_Linea = "Reporte de Headcount"
'    fExport.writeline Aux_Linea
'    fExport.writeline ""
'
'    'Aux_Linea = "Modelo" & Separador & "Descripción" & Separador & "Cuenta" & Separador & "Monto" & Separador & "Acumulado" & Separador & "Porcentaje" & Separador & "Legajo" & Separador & "Apellido" & Separador & "Proyecto" & Separador & "Nivel Costo 1" & Separador & "Nivel Costo 2" & Separador & "Nivel Costo 3" & Separador & "Nivel Costo 4" & Separador & "Tipo Origen" & Separador & "Origen"
'    Aux_Linea = "Codigo" & Separador & "Nombre" & Separador & "cencos" & Separador & "d_cencos" & Separador & "cantidad" & Separador & "porcen" & Separador & "ubica" & Separador & "d_ubica" & Separador & "d_clase" & Separador & "d_gerencia" & Separador & "tipocon" & Separador & "estado" & Separador & "fecha_ing" & Separador & "fecha_ret" & Separador & "reppag" & Separador & "cargo" & Separador & "d_cargo" & Separador & "sexo" & Separador & "Mano de obra"
'    fExport.writeline Aux_Linea

    'EAM- Seteo el legajo en 0 para usarlo como auxiliar e ir a buscar el saldo del asiento para el empleado
    empleg = 0
    Monto = 0
    Aux_sql = ""
Do While Not rs_Detalles.EOF
        porcentajeAux = 0
        cencosAux = ""
        ubicaAux = ""
        d_gerenciaAux = ""
        If empleg <> rs_Detalles!empleg Then
            empleg = rs_Detalles!empleg
            
            StrSql = "SELECT SUM(dlmonto) Monto FROM detalle_asi WHERE detalle_asi.vol_cod in(" & Vol_Cod & ") and linad_h=0 and empleg= " & empleg
            OpenRecordset StrSql, rs_Monto
            
            Monto = IIf(EsNulo(rs_Monto!Monto), 0, rs_Monto!Monto)
        End If
        'COl 0- Id de la cabecera del headcount
        Aux_sql = headcountID
        
        'COl 1- Modelo, Descripcion,cuenta,monto
        Aux_Linea = rs_Detalles!empleg
        Aux_sql = Aux_sql & "," & rs_Detalles!empleg
        
        'COl 2-Apellido, Proyecto
        Aux_Linea = Aux_Linea & Separador & rs_Detalles!terape
        Aux_sql = Aux_sql & ",'" & rs_Detalles!terape & "'"
        
        'COl 3-Codigo externo del centro de costo del empleado
        cencosAux = "'" & IIf(EsNulo(rs_Detalles!estrcodext1), "", rs_Detalles!estrcodext1) & "'"
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!estrcodext1), "", rs_Detalles!estrcodext1)
        Aux_sql = Aux_sql & ",'" & IIf(EsNulo(rs_Detalles!estrcodext1), "", rs_Detalles!estrcodext1) & "'"
        
        'COl 4-Descripcion NivelCosto1 -> Centro de costo
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!estrdabr1), "0", rs_Detalles!estrdabr1)
        Aux_sql = Aux_sql & ",'" & IIf(EsNulo(rs_Detalles!estrdabr1), "0", rs_Detalles!estrdabr1) & "'"
        
        'COl 5-Cantidad (mensuales -->0 | resto-> suma de CC)
        If (rs_Detalles!tprocnro = 3) Then
            Aux_Linea = Aux_Linea & Separador & "0"
            Aux_sql = Aux_sql & ",'" & 0 & "'"
            CantidadAux = 0
        Else
'            StrSql = " SELECT SUM(dlicant) cantidad from cabliq " & _
'                    " INNER JOIN detliq on cabliq.cliqnro = detliq.cliqnro " & _
'                    " INNER JOIN concepto ON concepto.concnro = detliq.concnro " & _
'                    " WHERE concepto.conccod IN (" & listCC & ") AND pronro= " & rs_Detalles!pronro & " and empleado= " & rs_Detalles!Ternro
            
            'StrSql = "SELECT SUM(dlcantidad) cantidad FROM detalle_asi WHERE ternro= " & rs_Detalles!Ternro & " and vol_cod = " & Vol_Cod & _
            '        " AND origen IN  (" & listCC & ") and dlcosto1= " & rs_Detalles!dlcosto1 & " and dlcosto4= " & rs_Detalles!pronro
            
            'FGZ - 18/07/2014 -------
            StrSql = "SELECT SUM(dlcantidad) cantidad FROM detalle_asi WHERE ternro= " & rs_Detalles!Ternro & " and vol_cod IN (" & Vol_Cod & ")" & _
                    " AND origen IN  (" & listCC & ") and dlcosto1= " & rs_Detalles!dlcosto1 & " and dlcosto4= " & rs_Detalles!pronro
            OpenRecordset StrSql, rs_detliq
            
            Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_detliq!Cantidad), "0", rs_detliq!Cantidad)
            Aux_sql = Aux_sql & "," & IIf(EsNulo(rs_detliq!Cantidad), "0", rs_detliq!Cantidad)
            CantidadAux = IIf(EsNulo(rs_detliq!Cantidad), "0", rs_detliq!Cantidad)
        End If
                
        'COl 6- Porcentaje (mensuales -->muestra % disribución del asiento | resto-> suma de CC /20)
        If (rs_Detalles!tprocnro = 3) Then
          If Monto = 0 Then
            Aux_Linea = Aux_Linea & Separador & 0
            Aux_sql = Aux_sql & "," & 0
            porcentajeAux = 0
            
          Else
            Aux_Linea = Aux_Linea & Separador & Round((rs_Detalles!montoCC) / Monto, 4)
            Aux_sql = Aux_sql & "," & Round((rs_Detalles!montoCC) / Monto, 4)
            porcentajeAux = Round((rs_Detalles!montoCC) / Monto, 4)
          End If
        Else
            If Not IsNull(rs_detliq!Cantidad) Then
                'Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_detliq!Cantidad), "0", IIf((rs_detliq!Cantidad / 20) > 1, "1", Round((rs_detliq!Cantidad / 20), 4)))
                Aux_Linea = Aux_Linea & Separador & IIf((rs_detliq!Cantidad / 20) > 1, "1", Round((rs_detliq!Cantidad / 20), 4))
                'Aux_sql = Aux_sql & "," & IIf(EsNulo(rs_detliq!Cantidad), "0", IIf((rs_detliq!Cantidad / 20) > 1, "1", Round((rs_detliq!Cantidad / 20), 4)))
                Aux_sql = Aux_sql & "," & IIf((rs_detliq!Cantidad / 20) > 1, "1", Round((rs_detliq!Cantidad / 20), 4))
                porcentajeAux = Round((rs_detliq!Cantidad / 20), 4)
            Else
                Aux_Linea = Aux_Linea & Separador & "0"
                Aux_sql = Aux_sql & ",0"
                porcentajeAux = 0
            End If
        End If
        
        'COl 7- Codigo externo de la estructura Lugar de Trabajo
        ubicaAux = "'" & IIf(EsNulo(rs_Detalles!estrcodext2), "", rs_Detalles!estrcodext2) & "'"
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!estrcodext2), "", rs_Detalles!estrcodext2)
        Aux_sql = Aux_sql & ",'" & IIf(EsNulo(rs_Detalles!estrcodext2), "", rs_Detalles!estrcodext2) & "'"
        
        'COl 8- Descripcion NivelCosto2 -> Lugar de Trabajo
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!estrdabr2), "", rs_Detalles!estrdabr2)
        Aux_sql = Aux_sql & ",'" & IIf(EsNulo(rs_Detalles!estrdabr2), "", rs_Detalles!estrdabr2) & "'"
        
        'COl 9- Descripcion del la estructura convenio del empleado
        Aux_Linea = Aux_Linea & Separador & EstructuraPropiedad(rs_Detalles!Ternro, 19, rs_Detalles!profecfin, 0)
        Aux_sql = Aux_sql & ",'" & EstructuraPropiedad(rs_Detalles!Ternro, 19, rs_Detalles!profecfin, 0) & "'"
        
        
        'COl 10- Descripcion del la estructura Gerencia del empleado
        d_gerenciaAux = "'" & EstructuraPropiedad(rs_Detalles!Ternro, 35, rs_Detalles!profecfin, 0) & "'"
        Aux_Linea = Aux_Linea & Separador & EstructuraPropiedad(rs_Detalles!Ternro, 35, rs_Detalles!profecfin, 0)
        Aux_sql = Aux_sql & ",'" & EstructuraPropiedad(rs_Detalles!Ternro, 35, rs_Detalles!profecfin, 0) & "'"
        
        'COl 11- Tipo de Contrato (Mensuales ->I, sino ->O)
        StrSql = "SELECT confval2 FROM his_estructura he " & _
                " INNER JOIN confrep on confrep.confval = he.estrnro and repnro= 438 and upper(conftipo)='CON' " & _
                " WHERE ternro= " & rs_Detalles!Ternro & " and tenro= 18 and htetdesde<=" & ConvFecha(rs_Detalles!profecfin) & " and (htethasta<=" & ConvFecha(rs_Detalles!profecfin) & " or htethasta is null )"
        OpenRecordset StrSql, rs_Convenio
        If Not rs_Convenio.EOF Then
            Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Convenio!confval2), "", rs_Convenio!confval2)
            Aux_sql = Aux_sql & ",'" & IIf(EsNulo(rs_Convenio!confval2), "", rs_Convenio!confval2) & "'"
        Else
            Aux_Linea = Aux_Linea & Separador & ""
            Aux_sql = Aux_sql & ",''"
        End If
        
        'COl 12- Estado del empleado
        Aux_Linea = Aux_Linea & Separador & IIf((rs_Detalles!empest = -1), "A", "S")
        Aux_sql = Aux_sql & ",'" & IIf((rs_Detalles!empest = -1), "A", "S") & "'"
        
        
        'COl 13- Fecha de alta reconocida del empelado
        'StrSql = "SELECT altfec,bajfec FROM fases WHERE empleado=" & rs_Detalles!Ternro & " AND fasrecofec= -1 ORDER BY altfec DESC "
        StrSql = "SELECT altfec,bajfec FROM fases WHERE empleado=" & rs_Detalles!Ternro & " ORDER BY altfec DESC "
        OpenRecordset StrSql, rs_fase
        If Not rs_fase.EOF Then
            Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_fase!altfec), "", rs_fase!altfec)
            Aux_sql = Aux_sql & "," & IIf(EsNulo(rs_fase!altfec), "", ConvFecha(rs_fase!altfec))
            
            'COl 14- Si el empleado esta inactivo muestra la fecha sino "01/01/3000"
            Aux_Linea = Aux_Linea & Separador & IIf((Not EsNulo(rs_fase!bajfec) And (rs_Detalles!empest = 0)), rs_fase!bajfec, "01/01/3000")
            
            If Not EsNulo(rs_fase!bajfec) And (rs_Detalles!empest = 0) Then
                Aux_sql = Aux_sql & "," & ConvFecha(rs_fase!bajfec)
            Else
                Aux_sql = Aux_sql & "," & "'01/01/3000'"
            End If
            
            'Aux_sql = Aux_sql & "," & IIf((Not EsNulo(rs_fase!bajfec) And (rs_Detalles!empest = 0)), ConvFecha(rs_fase!bajfec), "'01/01/3000'")
        Else
            Aux_Linea = Aux_Linea & Separador & "" & Separador & ""
            Aux_sql = Aux_sql & ",'',''"
        End If
        
        
        
        'COl 15- Forma de liquidación
        Aux_Linea = Aux_Linea & Separador & Mid(EstructuraPropiedad(rs_Detalles!Ternro, 22, rs_Detalles!profecfin, 0), 1, 1)
        Aux_sql = Aux_sql & ",'" & IIf(EsNulo(Mid(EstructuraPropiedad(rs_Detalles!Ternro, 22, rs_Detalles!profecfin, 0), 1, 1)), "''", Mid(EstructuraPropiedad(rs_Detalles!Ternro, 22, rs_Detalles!profecfin, 0), 1, 1)) & "'"
        
        'COl 16- Codigo externo estructura Puesto del empleado
        Aux_Linea = Aux_Linea & Separador & EstructuraPropiedad(rs_Detalles!Ternro, 4, rs_Detalles!profecfin, 1)
        Aux_sql = Aux_sql & ",'" & IIf(EsNulo(EstructuraPropiedad(rs_Detalles!Ternro, 4, rs_Detalles!profecfin, 1)), "''", EstructuraPropiedad(rs_Detalles!Ternro, 4, rs_Detalles!profecfin, 1)) & "'"
        
        'COl 17- Descripción de la estructura Puesto del empleado
        Aux_Linea = Aux_Linea & Separador & EstructuraPropiedad(rs_Detalles!Ternro, 4, rs_Detalles!profecfin, 0)
        Aux_sql = Aux_sql & "," & IIf(EsNulo(EstructuraPropiedad(rs_Detalles!Ternro, 4, rs_Detalles!profecfin, 0)), "''", "'" & EstructuraPropiedad(rs_Detalles!Ternro, 4, rs_Detalles!profecfin, 0) & "'")
        
        'COl 18- Sexo del empleado
        Aux_Linea = Aux_Linea & Separador & IIf((rs_Detalles!tersex = -1), "'M'", "'F'")
        Aux_sql = Aux_sql & "," & IIf((rs_Detalles!tersex = -1), "'M'", "'F'")
        
        
        'COl 19- Descripción de la estructura Mano de Obra del empleado
        Aux_Linea = Aux_Linea & Separador & EstructuraPropiedad(rs_Detalles!Ternro, 11, rs_Detalles!profecfin, 1)
        Aux_sql = Aux_sql & ",'" & EstructuraPropiedad(rs_Detalles!Ternro, 11, rs_Detalles!profecfin, 1) & "'"
                        
'        'Escribo en el archivo de texto
'        fExport.writeline Aux_Linea
            
        StrSql = "SELECT porcen, Cantidad FROM det_headcount where headcountid= " & headcountID & " and codigo= " & rs_Detalles!empleg & " and cencos= " & cencosAux & " and ubica= " & ubicaAux & " and d_gerencia= " & d_gerenciaAux
        OpenRecordset StrSql, rs_DetalleAux
        
        If Not rs_DetalleAux.EOF Then
            
            'StrSql = "UPDATE det_headcount SET porcen=" & (Round((rs_Detalles!montoCC) / Monto, 4) + rs_DetalleAux!porcen)
            StrSql = "UPDATE det_headcount SET porcen=" & (porcentajeAux + rs_DetalleAux!porcen) & _
                    ", cantidad= " & CantidadAux + rs_DetalleAux!Cantidad & _
                    " where headcountid= " & headcountID & " and codigo= " & rs_Detalles!empleg & " and cencos= " & cencosAux & " and ubica= " & ubicaAux & " and d_gerencia= " & d_gerenciaAux
        Else
            StrSql = "INSERT INTO det_headcount (headcountID, codigo, nombre, cencos, d_cencos,cantidad,porcen,ubica,d_ubica,d_clase,d_gerencia,tipocon,estado,fecha_ing, fecha_ret,reppag,cargo,d_cargo, sexo,Mano_de_obra) " & _
                    " VALUES ( " & Aux_sql & ")"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    'Siguiente proceso
    rs_Detalles.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans

StrSql = "SELECT * FROM det_headcount where headcountid= " & headcountID
OpenRecordset StrSql, rs_DetalleAux

'EAM - (v1.01) armo el excel
Flog.writeline Espacios(Tabulador * 2) & "Registros a Exportar: " & CConceptosAProc
'Genero los encabezados
Aux_Linea = "Reporte de Headcount"
fExport.writeline Aux_Linea
fExport.writeline ""
    
'Aux_Linea = "Modelo" & Separador & "Descripción" & Separador & "Cuenta" & Separador & "Monto" & Separador & "Acumulado" & Separador & "Porcentaje" & Separador & "Legajo" & Separador & "Apellido" & Separador & "Proyecto" & Separador & "Nivel Costo 1" & Separador & "Nivel Costo 2" & Separador & "Nivel Costo 3" & Separador & "Nivel Costo 4" & Separador & "Tipo Origen" & Separador & "Origen"
Aux_Linea = "Codigo" & Separador & "Nombre" & Separador & "cencos" & Separador & "d_cencos" & Separador & "cantidad" & Separador & "porcen" & Separador & "ubica" & Separador & "d_ubica" & Separador & "d_clase" & Separador & "d_gerencia" & Separador & "tipocon" & Separador & "estado" & Separador & "fecha_ing" & Separador & "fecha_ret" & Separador & "reppag" & Separador & "cargo" & Separador & "d_cargo" & Separador & "sexo" & Separador & "Mano de obra"
fExport.writeline Aux_Linea

Do While Not rs_DetalleAux.EOF
    Aux_Linea = "" 'rs_DetalleAux(0)
    For i = 1 To rs_DetalleAux.Fields.Count() - 1
        Aux_Linea = Aux_Linea & Separador & rs_DetalleAux(i)
    Next
    Aux_Linea = Right(Aux_Linea, Len(Aux_Linea) - 1)
    'Escribo en el archivo de texto
    
    fExport.writeline Aux_Linea
    rs_DetalleAux.MoveNext
Loop

'Cierro el archivo creado
fExport.Close



Flog.writeline Espacios(Tabulador * 2) & "Termino de Exportar "

Fin:
If rs_Detalles.State = adStateOpen Then rs_Detalles.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_detliq.State = adStateOpen Then rs_detliq.Close
If rs_fase.State = adStateOpen Then rs_fase.Close

Set rs_Detalles = Nothing
Set rs_Modelo = Nothing
Set rs_detliq = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    GoTo Fin
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : EAM
' Fecha      : 05/04/2014
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim Separador_de_Campos As String
Dim Vol_Cod As String
Dim Periodo As Long
Dim aux As String

'Orden de los parametros
'Periodo
'proceso de volcado
'Separador de campos en el archivo a generar

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
    
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Vol_Cod = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        If aux = 1 Then
            Separador_de_Campos = ","
        Else
            Separador_de_Campos = ";"
        End If
    End If
End If
Call Generacion(bpronro, Vol_Cod, Separador_de_Campos)
End Sub




Public Function EstructuraPropiedad(ByVal Ternro As Long, ByVal Tenro As Long, ByVal Fecha As Date, ByVal Propiedad As Integer) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca la descripcion de una estructura a una fecha.
' Autor      : Emanuel, Margiotta
' Fecha      : 06/04/2014
' ---------------------------------------------------------------------------------------------

Dim rsEstructura As New ADODB.Recordset

On Error GoTo CE

    StrSql = "SELECT e.estrdabr,e.estrcodext,e.Estrnro FROM his_estructura he " & _
            " INNER JOIN estructura e ON e.estrnro = he.estrnro " & _
            " WHERE he.ternro= " & Ternro & " AND he.tenro= " & Tenro & _
            " AND he.htetdesde<= " & ConvFecha(Fecha) & " AND (he.htethasta >= " & ConvFecha(Fecha) & " OR he.htethasta IS NULL)"
    OpenRecordset StrSql, rsEstructura
    
    If Not rsEstructura.EOF Then
        Select Case Propiedad
            Case 0: ' Descripcion
                EstructuraPropiedad = IIf(EsNulo(rsEstructura!estrdabr), "", rsEstructura!estrdabr)
            Case 1: ' Codigo externo
                EstructuraPropiedad = IIf(EsNulo(rsEstructura!estrcodext), "", rsEstructura!estrcodext)
            Case 2: 'Codigo de estructura
                EstructuraPropiedad = IIf(EsNulo(rsEstructura!Estrnro), "", rsEstructura!Estrnro)
            Case Else
                EstructuraPropiedad = IIf(EsNulo(rsEstructura!estrdabr), "", rsEstructura!estrdabr)
        End Select
    Else
        EstructuraPropiedad = ""
    End If
    
    Exit Function
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error al buscar la Estructura."
    Flog.writeline Espacios(Tabulador * 1) & StrSql
    Flog.writeline
End Function


Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si el proceso esta en condiciones de ejecutarse.
' Autor      : FGZ
' Fecha      : 05/08/2009
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

Select Case TipoProceso
Case 66: 'Exportacion de detalle de Asiento Contable
'    If Version >= "1.03" Then
'
'
'        Texto = "Revisar que exista tabla XXXX y su estructura sea correcta."
'
'        StrSql = "Select campo FROM tabla WHERE campo = 1"
'        OpenRecordset StrSql, rs
'
'        V = True
'    End If
    



'Cambio en busqueda de embargos bus_embargos
Case Else:
    Texto = "version correcta"
    V = True
End Select



    ValidarV = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function
