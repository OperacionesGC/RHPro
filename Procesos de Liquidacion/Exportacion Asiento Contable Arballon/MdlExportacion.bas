Attribute VB_Name = "MdlExportacion"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "08/05/2006"
'Global Const UltimaModificacion = " " 'Nro de version

'Global Const Version = "1.02"
'Global Const FechaModificacion = "09/05/2006"
'Global Const UltimaModificacion = " " 'Corta el CC a 4 digitos y el nro de cuenta a 10 digitos


'Global Const Version = "1.03"
'Global Const FechaModificacion = "20/02/2009"
'Global Const UltimaModificacion = " " 'Encriptacion de string de conexion

'Global Const Version = "1.04"
'Global Const FechaModificacion = "12/05/2009"
'Global Const UltimaModificacion = " " 'Se cambio lo que se informa en la cuenta analitica

'Global Const Version = "1.05"
'Global Const FechaModificacion = "10/09/2009"
'Global Const UltimaModificacion = " " 'Se cambio lo que se informa en la cuenta analitica

'Global Const Version = "1.06"
'Global Const FechaModificacion = "11/09/2009"
'Global Const UltimaModificacion = " " 'Se agrego distinct en la query principal

'Global Const Version = "1.07"
'Global Const FechaModificacion = "29/09/2009"
'Global Const UltimaModificacion = " " 'Se eliminan los 0 delante de la cuenta (concepto) cod_att

'Global Const Version = "1.08"
'Global Const FechaModificacion = "22/10/2009"
'Global Const UltimaModificacion = " "
'Todas las cuentas son de 14 o 8 caracteres una vez generadas con los siguientes casos y como deben
'ser informadas a Arballon:
'Nro    Cuenta                Cuenta Generada   Saldo  Arballon
'1      05210501E1E1E1E1E1E1  05210501654321    D      52105016D       'Centro Costo 654321
'1      05210501E1E1E1E1E1E1  05210501054321    D      52105010D       'Centro Costo 054321. Se puede presentar este caso? Va el 0 final o no
'12     00021301              00021301          H      21301H
'16     00114232LLLLLL        00114232102511    D      1142321D        'Legajo 102511
'16     00114232LLLLLL        00114232102511    H      1142321H        'Legajo 102511
'16     00114232LLLLLL        00114232013249    D      114232D         'Legajo 013249
'16     00114232LLLLLL        00114232013249    H      114232H         'Legajo 013249

'De lo anterior deduzco que la configuración se mantiene y que la modificación en la interfase es en
'el caso marcado en oscurita. Creo que abarcamos todos los casos que se nos pueden presentar en las
'cuentas ya generadas.

'RESPUESTA DE BIBIANA (NORTON) DIA 22/10/2009
'HAY QUE ELIMINAR LOS 0 DE LA CUENTA Y DE LOS CENTROS DE COSTOS (0 DE ADELANTE) Y SIEMPRE
'CORTAR EN EL OCTAVO CAMPO Y EN EL NOVENO VA LA H O LA D
'1      05210501E1E1E1E1E1E1  05210501054321    D      5210501D        'Centro Costo 54321. .

'Global Const Version = "1.09"
'Global Const FechaModificacion = "11/11/2009"
'Global Const UltimaModificacion = " "
'FAF - Como no se blanquea queda con el valor anterior, variable CtaAnalitica

'Global Const Version = "1.10"
'Global Const FechaModificacion = "11/11/2009"
'Global Const UltimaModificacion = " "
'FAF Se eliminan los 0 delante de la cuenta Analitica. Campo cta_anl de la tabla anlcptgen. Ej 00021305 --> 21305

'Global Const Version = "1.11"
'Global Const FechaModificacion = "12/11/2009"
'Global Const UltimaModificacion = " "
'FAF Para una cta en particular se debe generar el saldo del haber con la letra h minuscula. Cta 114233h

'Global Const Version = "1.12"
'Global Const FechaModificacion = "13/11/2009"
'Global Const UltimaModificacion = " "
'FAF Faltaba agregar 2 y 3 nivel de apertura por esctructura en CalcularCtaAnalitica.

'Global Const Version = "1.13"
'Global Const FechaModificacion = "18/11/2009"
'Global Const UltimaModificacion = " "
'FAF Faltaba que calculara la Cta Analitica cuando ya esta expresada e el Plan de cuenta. Por ej 00114233008080.
' La cuenta es 00114233 y la cta analitica 008080

'Global Const Version = "1.14"
'Global Const FechaModificacion = "09/12/2009"
'Global Const UltimaModificacion = " "
''FAF - El total debe totalizar, contemplar si suma o resta (debe o haber) en la tabla anlcptgen, columna mto_cto

'Global Const Version = "1.15"
'Global Const FechaModificacion = "18/07/2014"
'Global Const UltimaModificacion = " "
'Carmen Quintero - CAS-26415 - Bodegas Norton - Bug Asiento contable sistema Arballon  - se actualiza el signo del campo mto_cto y cod_att cuando el valor es negativo
'cuando se realiza un cambio de cuenta analitica sienpre y cuando sea de produccion empiece en 52

Global Const Version = "1.16"
Global Const FechaModificacion = "31/07/2014"
Global Const UltimaModificacion = " "
'Carmen Quintero - CAS-26415 - Bodegas Norton - Bug Asiento contable sistema Arballon [Entrega 2]  - Se modificó para que inserte 2 registros en la tabla anlcptgen
'cuando la cuenta comiencen con "052" (de producción), y si en la tabla línea_asi, existan registros para la misma cuenta, con montos en el debe y en el haber


'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------

Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser               As String
Global Fecha                As Date
Global hora                 As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa           As New ADODB.Recordset
Global rs_tipocod           As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo    As Date
Global StrSql2              As String
Global SeparadorDecimales   As String
Global totalImporte         As Double
Global total                As Double
Global TotalABS             As Double
Global UltimaLeyenda        As String
Global EsUltimoItem         As Boolean
Global EsUltimoProceso      As Boolean


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

    Nombre_Arch = PathFLog & "Exp_Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    'Abro la conexion
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
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 116 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encuentran los datos del proceso"
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
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub

Function BuscarMonto(vol_cod As String, tipo As Integer) As Double

    Dim rs_Monto As New ADODB.Recordset
        
    StrSql = " SELECT SUM (monto) monto FROM linea_asi "
    StrSql = StrSql & " WHERE vol_cod IN (" & vol_cod & ")"
    StrSql = StrSql & " AND cuenta like '%052%' "
    StrSql = StrSql & " AND dh = " & tipo

    If tipo = -1 Then
        Flog.writeline "Monto Debe: " & StrSql
    Else
        Flog.writeline "Monto Haber: " & StrSql
    End If
    
    OpenRecordset StrSql, rs_Monto
    If Not rs_Monto.EOF Then
        If Not IsNull(rs_Monto("monto")) Then
            BuscarMonto = rs_Monto("monto")
        Else
            BuscarMonto = 0
        End If
    End If
    
End Function

Public Sub Generacion(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Asinro As String, ByVal Empresa As Long, ByVal ProcVol As String) ', ByVal Minuta As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la exportacion de Asientos Contables
' Autor      : Fapitalle N.
' Fecha      : 09/11/2005
' Modificado : FGZ - 08/05/2006
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim directorio As String

Dim vol_cod_ant
Dim cuenta_ant
Dim masinro_ant
Dim linea_ant
Dim linacuenta
Dim linacuenta_aux
Dim nro_cuenta As String
Dim nro_cuenta_7 As String
Dim ccosto As String
Dim ccosto_insertado As Boolean
Dim ID As Integer
Dim total_modelo
Dim secuencia
Dim CtaAnalitica As String
Dim Seguir As Boolean

Dim Nro As Long
Dim separadorCampos
Dim minuta_insertada As Boolean

Dim objConnArballon As New ADODB.Connection

'Registros
Dim rs_ID As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 234" 'modelo de exportacion de asiento
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        directorio = directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
    SeparadorDecimales = rs_Modelo!modsepdec
    separadorCampos = rs_Modelo!modseparador
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Periodo"
    Exit Sub
End If

' manejador de errores
On Error GoTo CE

' Comienzo la transaccion
MyBeginTrans

'Busco los procesos a evaluar
'11-09-2009 - FAF - Se agrego DISTINCT
StrSql = " SELECT DISTINCT linea_asi.*, "
StrSql = StrSql & " mod_linea.linacuenta, "
StrSql = StrSql & " proc_vol.vol_fec_asiento, proc_vol.pliqnro, proc_vol.vol_desc,"
StrSql = StrSql & " arballon_mod_asiento.centro_op, arballon_mod_asiento.moneda, arballon_mod_asiento.suboper, arballon_mod_asiento.operador,"
StrSql = StrSql & " estr_cod.nrocod,"
StrSql = StrSql & " mod_asiento.masidesc,"
StrSql = StrSql & " conexion.cnstring"
StrSql = StrSql & " From linea_asi"
StrSql = StrSql & " INNER JOIN mod_linea ON mod_linea.linaorden = linea_asi.linea and linea_asi.masinro = mod_linea.masinro"
StrSql = StrSql & " INNER JOIN proc_vol ON proc_vol.vol_cod = linea_asi.vol_cod"
StrSql = StrSql & " INNER JOIN arballon_mod_asiento ON linea_asi.masinro = arballon_mod_asiento.masinro"
StrSql = StrSql & " LEFT JOIN proc_vol_pl ON proc_vol.vol_cod = proc_vol_pl.vol_cod"
StrSql = StrSql & " LEFT JOIN proceso ON proceso.pronro = proc_vol_pl.pronro"
StrSql = StrSql & " LEFT JOIN empresa ON empresa.empnro = proceso.empnro"
StrSql = StrSql & " LEFT JOIN estr_cod ON estr_cod.estrnro = empresa.estrnro and estr_cod.tcodnro = 46"
StrSql = StrSql & " INNER JOIN mod_asiento ON mod_asiento.masinro = arballon_mod_asiento.masinro"
StrSql = StrSql & " LEFT JOIN conexion ON conexion.cnnro = arballon_mod_asiento.cnnro"
StrSql = StrSql & " WHERE proc_vol.pliqnro = " & Nroliq
If ProcVol <> 0 Then 'si no son todos
    StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
End If
StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
StrSql = StrSql & " ORDER BY linea_asi.vol_cod, linea_asi.masinro, linea_asi.linea"
Flog.writeline "Main Query: " & StrSql
OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / (CConceptosAProc))

'Procesamiento
If rs_Procesos.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
End If

'------------------------------------------------------------------------
' Genero el detalle de la exportacion
'------------------------------------------------------------------------

'31/07/2014
'inicio
Dim MontoDebe
Dim MontoHaber
If ProcVol <> 0 Then
    MontoDebe = BuscarMonto(ProcVol, -1)
    MontoHaber = BuscarMonto(ProcVol, 0)
End If
'fin

EsUltimoItem = False
EsUltimoProceso = False
vol_cod_ant = -1
masinro_ant = -1
linea_ant = -1
minuta_insertada = False
ccosto_insertado = False

Do While Not rs_Procesos.EOF
    If EsUltimoRegistro(rs_Procesos) Then
        EsUltimoProceso = True
    End If
    Flog.writeline
    Cantidad_Warnings = 0
    Nro = Nro + 1 'Contador de Lineas
'-*-*-*-*-*-*------------------------------*-*-*-*-*-*-*-*-*---------------
    
    If rs_Procesos!vol_cod <> vol_cod_ant Then 'cambia el proceso de volcado
        vol_cod_ant = rs_Procesos!vol_cod
        masinro_ant = -1
        linea_ant = -1
        'hacer lo que se hace para cada proceso de volcado
        Flog.writeline " Proceso de Volcado: " & vol_cod_ant
    End If
    
    If rs_Procesos!masinro <> masinro_ant Then 'cambia el modelo
        
        If Not IsNull(rs_Procesos!cnstring) Then
            OpenConnection rs_Procesos!cnstring, objConnArballon
            Flog.writeline " Usando Conexion Del Modelo: " & rs_Procesos!cnstring
        Else
            OpenConnection strconexion, objConnArballon
            Flog.writeline " Usando Conexion Default: " & strconexion
        End If
        
'        'hacer lo que se hace para cada modelo
'        If minuta_insertada Then
'            ID = ID + 1
'        Else
'            ID = Minuta 'inicialmente pone la minuta que fue puesta manualmente
'            minuta_insertada = True
'        End If
'        Flog.writeline " Numero de Minuta: " & ID
        
        'FGZ - 08/05/2006
        StrSql = "SELECT MAX(nro_ope) nro_ope FROM anlcabpos"
        OpenRecordsetWithConn StrSql, rs_ID, objConnArballon
        If Not rs_ID.EOF Then
            If Not IsNull(rs_ID!nro_ope) Then ID = rs_ID!nro_ope + 1 Else ID = 1
        Else
            ID = 1
        End If
        
        'Esto seria lo que deberiamos hacer si el campo fuese serial pero ...aparentemente no lo es
        'ID = getLastIdentity(objConnArballon, "anlcabpos")
        Flog.writeline " Numero de Minuta a insertar: " & ID
        
        
        
        secuencia = 0
        ' INSERT en la tabla anlcabpos por CADA MODELO del proceso de volcado
        Flog.writeline "  Nuevo Modelo de Asiento: " & rs_Procesos!masinro
        StrSql = "INSERT INTO anlcabpos "
        StrSql = StrSql & "(tip_act, cod_emp, cod_ope, cen_ope, nro_ope, "
        StrSql = StrSql & " fch_ope, cod_ref, cen_ref, nro_ref, fch_ref, "
        StrSql = StrSql & " prv_ref, mda_ref, hoj_ref, cod_adu, net_gra, "
        StrSql = StrSql & " fch_pza, cod_mda, cot_mda, sub_ope, cod_opr, "
        StrSql = StrSql & " cod_fin, obs_ope, mto_tot, cod_edo )"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & "'A', "                                       ' 1 tip_act char(1) not null,
        If IsNull(rs_Procesos!nrocod) Then Flog.writeline " **** ERROR: No existe el codigo de la empresa. Configure uno. **** "
        StrSql = StrSql & rs_Procesos!nrocod & ", "                     ' 2 cod_emp smallint not null,
'       StrSql = StrSql & "558, "     ' DEBUG LINE PLEASE DELETE !!!!!!
        StrSql = StrSql & "'Mta', "                                     ' 3 cod_ope char(6) not null,
        StrSql = StrSql & rs_Procesos!centro_op & ", "                  ' 4 cen_ope smallint not null,
        StrSql = StrSql & ID & ", "                                     ' 5 nro_ope int not null,
        StrSql = StrSql & ConvFecha(rs_Procesos!vol_fec_asiento) & ", " ' 6 fch_ope datetime not null,
        StrSql = StrSql & "'', "                                        ' 7 cod_ref char(6) not null,
        StrSql = StrSql & "0, "                                         ' 8 cen_ref smallint not null,
        StrSql = StrSql & "0, "                                         ' 9 nro_ref int not null,
        StrSql = StrSql & "'', "                                        '10 fch_ref datetime not null,
        StrSql = StrSql & "'', "                                        '11 prv_ref char(4) not null,
        StrSql = StrSql & "'', "                                        '12 mda_ref char(4) not null,
        StrSql = StrSql & "0, "                                         '13 hoj_ref smallint not null,
        StrSql = StrSql & "'', "                                        '14 cod_adu char(4) not null,
        StrSql = StrSql & "0, "                                         '15 net_gra decimal(18,2) not null,
        StrSql = StrSql & "'', "                                        '16 fch_pza datetime not null,
        StrSql = StrSql & "'" & rs_Procesos!moneda & "', "              '17 cod_mda char(4) not null,
        StrSql = StrSql & "1, "                                         '18 cot_mda decimal(18,6) not null,
        StrSql = StrSql & rs_Procesos!suboper & ", "                    '19 sub_ope smallint not null,
        StrSql = StrSql & "'" & rs_Procesos!operador & "', "            '20 cod_opr char(8) not null,
        StrSql = StrSql & "'', "                                        '21 cod_fin char(3) not null,
        StrSql = StrSql & "'" & Left("RHPro " & rs_Procesos!vol_desc & " - " & rs_Procesos!masidesc, 60) & "', " '22 obs_ope char(60) not null,
        StrSql = StrSql & "0, "                                         '23 mto_tot decimal(18,2) not null,
        StrSql = StrSql & "0)"                                          '24 cod_edo smallint not null
        Flog.writeline "   SQL: " & StrSql
        objConnArballon.Execute StrSql, , adExecuteNoRecords
        '*/*/*/*/*/*/*/*/*/*/*/*
        masinro_ant = rs_Procesos!masinro
        
        ' INSERT en la tabla anlvtogen por CADA MODELO del proceso de volcado
        StrSql = "INSERT INTO anlvtogen ("
        StrSql = StrSql & " cod_emp, cod_ope, cen_ope, "
        StrSql = StrSql & " nro_ope, fch_vto, mto_vto )"
        StrSql = StrSql & " VALUES ("
        If IsNull(rs_Procesos!nrocod) Then Flog.writeline " **** ERROR: No existe el codigo de la empresa. Configure uno. **** "
        StrSql = StrSql & rs_Procesos!nrocod & ", "                     ' 2 cod_emp smallint not null,
'       StrSql = StrSql & "558, "     ' DEBUG LINE PLEASE DELETE !!!!!!
        StrSql = StrSql & "'Mta', "                                     ' 3 cod_ope char(6) not null,
        StrSql = StrSql & rs_Procesos!centro_op & ", "                  ' 4 cen_ope smallint not null,
        StrSql = StrSql & ID & ", "                                     ' 5 nro_ope int not null,
        StrSql = StrSql & ConvFecha(rs_Procesos!vol_fec_asiento) & ", " ' 6 fch_vto datetime not null,
        StrSql = StrSql & "0)"                                          ' 7 mto_vto decimal(18,2) not null
        Flog.writeline "   SQL: " & StrSql
        objConnArballon.Execute StrSql, , adExecuteNoRecords
        
        total_modelo = 0
        linea_ant = -1
    End If
    
    If rs_Procesos!Linea <> linea_ant Then 'cambia la linea
        linea_ant = rs_Procesos!Linea
        linacuenta = rs_Procesos!linacuenta
        ccosto_insertado = False
        
        '31/07/2014
        'inicio
        If Mid(nro_cuenta_7, 1, 2) = "52" Then
            Dim nuevomonto
            Dim letra
            Dim pos1
            Dim auxcuenta
            Dim nuevacuenta
            StrSql = "SELECT * FROM anlcptgen "
            StrSql = StrSql & " WHERE cod_att = '" & nro_cuenta_7 & "' "
            Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
            OpenRecordsetWithConn StrSql, rs, objConnArballon
            If Not rs.EOF Then
                If Not IsNull(rs("mto_cto")) Then
                    If CDbl(rs("mto_cto")) < 0 Then
                        If MontoDebe <> 0 And MontoHaber <> 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "Monto Debe: " & MontoDebe
                            
                            'Actualizo monto debe en la tabla anlcptgen
                            StrSql = "UPDATE anlcptgen SET mto_cto = " & MontoDebe & ""
                            StrSql = StrSql & " WHERE cod_att = '" & nro_cuenta_7 & "'"
                            Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
                            objConnArballon.Execute StrSql, , adExecuteNoRecords
                            
                            'Incremento la secuencia
                            secuencia = secuencia + 1
                            Flog.writeline Espacios(Tabulador * 1) & "Nueva secuencia: " & secuencia
                            
                            ' se arma la nueva cuenta
                            letra = "D"
                            pos1 = InStr(1, nro_cuenta_7, letra)
                            Flog.writeline Espacios(Tabulador * 1) & "Posicion de D: " & pos1
    
                            auxcuenta = Mid(nro_cuenta_7, 1, pos1 - 1)
                            Flog.writeline Espacios(Tabulador * 1) & "Cuenta Auxiliar: " & auxcuenta
    
                            nuevacuenta = auxcuenta + "H"
                            
                            'Se inserta el nuevo registro con monto haber y nombre de cuenta con H al final
                            Flog.writeline Espacios(Tabulador * 1) & "Insertar Modelo de Cuenta: " & nuevacuenta
                
                            StrSql = "INSERT INTO anlcptgen ("
                            StrSql = StrSql & " nro_nat, cod_emp, cod_ope, cen_ope, nro_ope, "
                            StrSql = StrSql & " nro_itm, cod_att, cta_anl, mto_cto )"
                            StrSql = StrSql & " VALUES ("
                            StrSql = StrSql & "0, "                                  ' 1 nro_nat smallint null,
                            StrSql = StrSql & rs!cod_emp & ", "                      ' 2 cod_emp smallint not null,
                            StrSql = StrSql & "'Mta', "                              ' 3 cod_ope char(6) not null,
                            StrSql = StrSql & rs!cen_ope & ", "                      ' 4 cen_ope smallint not null,
                            StrSql = StrSql & rs!nro_ope & ", "                      ' 5 nro_ope int not null,
                            StrSql = StrSql & secuencia & ", "                       ' 6 nro_itm smallint not null,
                            StrSql = StrSql & "'" & nuevacuenta & "', "              ' 7 cod_att char(10) not null,
                            StrSql = StrSql & "'',"                                  ' 8 cta_anl char(8) not null,
                            StrSql = StrSql & MontoHaber & ")"                       ' 9 mto_cto decimal(18, 2) not null,
                            Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
                            objConnArballon.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
            End If
        End If
        'fin
        
        'hacer lo que se hace para cada linea
        Flog.writeline Espacios(Tabulador * 1) & "Nueva Linea: " & linacuenta
        
    End If
    
    Call correspondeCentroCosto(rs_Procesos!cuenta, rs_Procesos!linacuenta, ccosto, nro_cuenta)
    If (ccosto <> "") Then
        If (CLng(ccosto) > 9999) Then
            Err.Description = Espacios(Tabulador * 1) & "Centro de Costo " & ccosto & " es mayor a 9999." & _
                                                        " Imposible insertar en Arballon."
            'FGZ - 09/05/2006 - lo corto en los 1eros 4
            'ccosto = Left(ccosto, 4)
            'FAF - 11/09/2009 - son los ultimos 4 digitos
            ccosto = Right(ccosto, 4)
            'GoTo CE
        End If
    End If
        
    'FAF - 11-11-2009 - Como no se blanquea queda con el valor anterior
    CtaAnalitica = ""
    'FAF - 11-11-2009
        
    'FGZ - 12/05/2009
    Call CalcularCtaAnalitica(rs_Procesos!cuenta, rs_Procesos!linacuenta, CtaAnalitica)
    
    '11/11/2009 - FAF Se eliminan los 0 delante de la cuenta Analitica. Ej 00021305 --> 21305
    Seguir = True
    Do While Seguir And Len(CtaAnalitica) > 0
        If Mid(CtaAnalitica, 1, 1) = "0" Then
            CtaAnalitica = Mid(CtaAnalitica, 2, Len(CtaAnalitica))
        Else
            Seguir = False
        End If
    Loop
    '11/11/2009 - FAF Se eliminan los 0 delante de la cuenta Analitica. Ej 00021305 --> 21305
    
    Flog.writeline "   Cta Analitica: " & CtaAnalitica
    
    If nro_cuenta = "" Then
        nro_cuenta = rs_Procesos!cuenta
    End If
    If rs_Procesos!dh = -1 Then
        total_modelo = total_modelo + rs_Procesos!Monto
        'nro_cuenta = nro_cuenta + "D"
        'FGZ - 09/05/2006
        'si el nro de cuenta es >= 10 hace macanas
        nro_cuenta = Left(nro_cuenta, 9) + "D"
    Else
        total_modelo = total_modelo - rs_Procesos!Monto
        'nro_cuenta = nro_cuenta + "H"
        'FGZ - 09/05/2006
        'si el nro de cuenta es >= 10 hace macanas
        nro_cuenta = Left(nro_cuenta, 9) + "H"
    End If
        
    'FAF - 22/10/2009. Si el caracter en la posicion 9 es 0, se debe eliminar
    If Mid(nro_cuenta, 9, 1) = "0" Then
        nro_cuenta = Mid(nro_cuenta, 1, 8) + Mid(nro_cuenta, 10, 1)
    End If
    'FAF - 22/10/2009
    
    
    'hacer el update para el total del monto en la tabla anlcabpos
    StrSql = "UPDATE anlcabpos SET mto_tot = " & total_modelo
    StrSql = StrSql & " WHERE nro_ope = " & ID
    Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
    objConnArballon.Execute StrSql, , adExecuteNoRecords
    'hacer el update para el total del monto en la tabla anlvtogen
    StrSql = "UPDATE anlvtogen SET mto_vto = " & total_modelo
    StrSql = StrSql & " WHERE nro_ope = " & ID
    Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
    objConnArballon.Execute StrSql, , adExecuteNoRecords
    
    '11/09/2009 - FAF ----------- Solamente se debe informar si comienza con 52 la cuenta (cta produccion)
    If Mid(rs_Procesos!cuenta, 1, 2) = "52" Or Mid(rs_Procesos!cuenta, 1, 3) = "052" Then
'    If ccosto <> "" Then 'la cuenta tiene centros de costo asociados
        If Not ccosto_insertado Then
            '11/09/2009 - FAF Si es 052... se debe eliminar el 0
            nro_cuenta_7 = Left(nro_cuenta, 10)
            If Mid(nro_cuenta_7, 1, 1) = "0" Then
                nro_cuenta_7 = Mid(nro_cuenta_7, 2, Len(nro_cuenta_7))
            End If
            '11/09/2009 - FAF Si es 052... se debe eliminar el 0
            'insertar en la tabla anlcptgen el encabezado correspondiente
            secuencia = secuencia + 1
            Flog.writeline Espacios(Tabulador * 1) & "Insertar Modelo de Cuenta: " & linacuenta
            ccosto_insertado = True
            StrSql = "INSERT INTO anlcptgen ("
            StrSql = StrSql & " nro_nat, cod_emp, cod_ope, cen_ope, nro_ope, "
            StrSql = StrSql & " nro_itm, cod_att, cta_anl, mto_cto )"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & "0, "             '                   ' 1 nro_nat smallint null,
            If IsNull(rs_Procesos!nrocod) Then Flog.writeline " **** ERROR: No existe el codigo de la empresa. Configure uno. **** "
            StrSql = StrSql & rs_Procesos!nrocod & ", "             ' 2 cod_emp smallint not null,
'           StrSql = StrSql & "558, "     ' DEBUG LINE PLEASE DELETE !!!!!!
            StrSql = StrSql & "'Mta', "                             ' 3 cod_ope char(6) not null,
            StrSql = StrSql & rs_Procesos!centro_op & ", "          ' 4 cen_ope smallint not null,
            StrSql = StrSql & ID & ", "                             ' 5 nro_ope int not null,
            StrSql = StrSql & secuencia & ", "                      ' 6 nro_itm smallint not null,
            StrSql = StrSql & "'" & nro_cuenta_7 & "', "            ' 7 cod_att char(10) not null,
            'FGZ - 12/05/2009 ---- cambió ---
            StrSql = StrSql & "'',"                                 ' 8 cta_anl char(8) not null,
            'StrSql = StrSql & CtaAnalitica & ", "                    ' 8 cta_anl char(8) not null,"
            'FGZ - 12/05/2009 ---- cambió ---
            StrSql = StrSql & " 0)"                                 ' 9 mto_cto decimal(18, 2) not null,
            Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
            objConnArballon.Execute StrSql, , adExecuteNoRecords
        End If
        'insertar en la tabla de centro de costo la fila del asiento
        Flog.writeline Espacios(Tabulador * 1) & "Insertar Centro de Costo: " & ccosto
        StrSql = "INSERT INTO cosbdcgen ("
        StrSql = StrSql & " cod_emp, cod_ope, cen_ope, nro_ope, nro_itm, "
        StrSql = StrSql & " cod_obs, cod_acv, cod_cen, cod_ord, nro_ord, "
        StrSql = StrSql & " uni_med, can, mto )"
        StrSql = StrSql & " VALUES ("
        If IsNull(rs_Procesos!nrocod) Then Flog.writeline " **** ERROR: No existe el codigo de la empresa. Configure uno. **** "
        StrSql = StrSql & rs_Procesos!nrocod & ", "     ' 2 cod_emp smallint not null,
'       StrSql = StrSql & "558, "     ' DEBUG LINE PLEASE DELETE !!!!!!
        StrSql = StrSql & "'Mta', "                     ' 3 cod_ope char(6) not null,
        StrSql = StrSql & rs_Procesos!centro_op & ", "  ' 4 cen_ope smallint not null,
        StrSql = StrSql & ID & ", "                     ' 5 nro_ope int not null,
        StrSql = StrSql & secuencia & ", "              '14 nro_itm smallint not null,
        StrSql = StrSql & "''" & ", "                   ' 6 cod_obs char(4) not null,
        StrSql = StrSql & "1" & ", "                    ' 7 cod_acv smallint not null,
        StrSql = StrSql & "'" & ccosto & "', "          ' 8 cod_cen smallint not null,
        StrSql = StrSql & "''" & ", "                   ' 9 cod_ord char(6) not null,
        StrSql = StrSql & "0" & ", "                    '10 nro_ord int not null,
        StrSql = StrSql & "''" & ", "                   '11 uni_med char(4) not null,
        StrSql = StrSql & "0" & ", "                    '12 can decimal(18,2) not null,
        StrSql = StrSql & rs_Procesos!Monto & ") "      '13 mto decimal(18,2) not null,
        Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
        objConnArballon.Execute StrSql, , adExecuteNoRecords
        
        'hacer el update del total para ese centro de costo
        '09-12-2009 - FAF - El total debe totalizar, contemplar si suma o resta
        If rs_Procesos!dh = -1 Then
            StrSql = "UPDATE anlcptgen SET mto_cto = mto_cto + " & rs_Procesos!Monto
        Else
            StrSql = "UPDATE anlcptgen SET mto_cto = mto_cto - " & rs_Procesos!Monto
        End If
        
'        StrSql = "UPDATE anlcptgen SET mto_cto = mto_cto + " & rs_Procesos!Monto
        '09-12-2009 - FAF - El total debe totalizar, contemplar si suma o resta
        StrSql = StrSql & " WHERE nro_ope = " & ID
        StrSql = StrSql & " AND nro_itm = " & secuencia
        Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
        objConnArballon.Execute StrSql, , adExecuteNoRecords
    Else
        'insertar la linea del asiento, la cuenta no tiene centro de costo asociado
        secuencia = secuencia + 1
        Flog.writeline Espacios(Tabulador * 1) & "Insertar Linea Sin CCosto Asociado: " & rs_Procesos!cuenta
            
        '28/09/2009 - FAF Se eliminan los 0 delante de la cuenta. Ej 00021305H --> 21305H
        nro_cuenta_7 = Left(nro_cuenta, 10)
        Seguir = True
        Do While Seguir And Len(nro_cuenta_7) > 0
            If Mid(nro_cuenta_7, 1, 1) = "0" Then
                nro_cuenta_7 = Mid(nro_cuenta_7, 2, Len(nro_cuenta_7))
            Else
                Seguir = False
            End If
        Loop
        '28/09/2009 - FAF Se eliminan los 0 delante de la cuenta. Ej 00021305H --> 21305H
        
        StrSql = "INSERT INTO anlcptgen ("
        StrSql = StrSql & " nro_nat, cod_emp, cod_ope, cen_ope, nro_ope, "
        StrSql = StrSql & " nro_itm, cod_att, cta_anl, mto_cto ) "
        StrSql = StrSql & " VALUES ("
        If IsNull(rs_Procesos!nrocod) Then Flog.writeline " **** ERROR: No existe el codigo de la empresa. Configure uno. **** "
        StrSql = StrSql & "0, "                                         ' 1 nro_nat smallint null,
        StrSql = StrSql & rs_Procesos!nrocod & ", "                     ' 2 cod_emp smallint not null,
'       StrSql = StrSql & "558, "     ' DEBUG LINE PLEASE DELETE !!!!!!
        StrSql = StrSql & "'Mta', "                                     ' 3 cod_ope char(6) not null,
        StrSql = StrSql & rs_Procesos!centro_op & ", "                  ' 4 cen_ope smallint not null,
        StrSql = StrSql & ID & ", "                                     ' 5 nro_ope int not null,
        StrSql = StrSql & secuencia & ", "                              ' 6 nro_itm smallint not null,
        'FAF - 12/11/2009 ---- cambió ---
        If Trim(nro_cuenta_7) = "114233H" Then
            nro_cuenta_7 = "114233h"
        End If
        'FAF - 12/11/2009 ---- cambió ---
        'FAF - 28/09/2009 ---- cambió ---
        StrSql = StrSql & "'" & nro_cuenta_7 & "', "                    ' 7 cod_att char(10) not null,
'        StrSql = StrSql & "'" & Left(nro_cuenta, 10) & "', "            ' 7 cod_att char(10) not null,
        'FAF - 28/09/2009 ---- cambió ---
        'FGZ - 12/05/2009 ---- cambió ---
        'StrSql = StrSql & "'',"                                        ' 8 cta_anl char(8) not null,
        StrSql = StrSql & "'" & CtaAnalitica & "', "                          ' 8 cta_anl char(8) not null,"
        'FGZ - 12/05/2009 ---- cambió ---
        StrSql = StrSql & rs_Procesos!Monto & ")"                       ' 9 mto_cto decimal(18, 2) not null,
        Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
        objConnArballon.Execute StrSql, , adExecuteNoRecords
    End If
        
'-*-*-*-*-*-*------------------------------*-*-*-*-*-*-*-*-*---------------
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans


If rs_ID.State = adStateOpen Then rs_ID.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
Set rs_ID = Nothing
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Modelo = Nothing

Exit Sub
CE:
    Flog.writeline " ************************************************************ "
    Flog.writeline " ***  Error: " & Err.Description
    Flog.writeline " ************************************************************ "
    HuboError = True
    MyRollbackTrans

    If rs_ID.State = adStateOpen Then rs_ID.Close
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    Set rs_ID = Nothing
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_Modelo = Nothing
End Sub

Public Sub correspondeCentroCosto(ByVal cuenta As String, ByVal modelo As String, ByRef ccosto As String, ByRef nro_cuenta As String)
' ----------------------------------------------------------------------------------------------------
' Descripcion:  evalua si el numero de cuenta corresponde a una cuenta que afecta a un centro de costo
'               segun el modelo de cuenta, si es asi devuelve el codigo del centro de costo en ccosto
'               sino en la misma variable devuelve ""
' Autor      : Fapitalle N.
' Fecha      : 10/11/2005
' Modificado :
' Ejemplo    : modelo = 6400E1E1E15, cuenta = 64002565
'              despues del bucle quedan
'              modelo = 64005, cuenta = 64005
'           Devuelve:
'              ccosto = 256
'              nro_cuenta = 64005
' ----------------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Indicador
Dim Reemplazo
Dim ccostoaux As String

Indicador = "E1"
Reemplazo = ""
ccosto = ""
nro_cuenta = ""
ccostoaux = ""
pos1 = InStr(1, modelo, Indicador) 'encuentro el primer indicador
Do While pos1 > 0
    'escribo el posible centro de costo
    ccostoaux = ccostoaux & Mid(cuenta, pos1, 1)
    'elimino el indicador actual
    modelo = Replace(modelo, Indicador, Reemplazo, 1, 1)
    'elimino de cuenta la posicion del indicador
    cuenta = Mid(cuenta, 1, pos1 - 1) & Reemplazo & Mid(cuenta, pos1 + 1, Len(cuenta))
    'encuentro el proximo indicador
    pos1 = InStr(1, modelo, Indicador)
Loop

If (cuenta = modelo) Then
    ccosto = ccostoaux
    nro_cuenta = cuenta
End If
End Sub

Public Sub CalcularCtaAnalitica(ByVal cuenta As String, ByVal modelo As String, ByRef cuenta_analitica As String)
' ----------------------------------------------------------------------------------------------------
' Descripcion:  Calcula cual es la cuenta analitica correspondiente
'               sino en la misma variable devuelve ""
' Autor      : Fapitalle N.
' Fecha      : 10/11/2005
' Modificado :
' Ejemplo    : modelo = 6400E1E1E1E1, cuenta = 64002565
'              despues del calculo la cuena analitica quedaria cuenta_analiica = 2565
' Ejemplo2   : modelo = 6400E2E2E2E2, cuenta = 64002565
'              despues del calculo la cuena analitica quedaria cuenta_analiica = ""
' Ejemplo3   : modelo = 6400LLLLLL, cuenta = 6400332565
'              despues del calculo la cuena analitica quedaria cuenta_analiica = 332565
' ----------------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Indicador
Dim nro_cuenta As String
Dim Reemplazo
Dim ccostoaux As String

Indicador = "E1"
Reemplazo = ""
nro_cuenta = ""
ccostoaux = ""
pos1 = InStr(1, modelo, Indicador) 'encuentro el primer indicador
If pos1 > 0 Then
    Do While pos1 > 0
        'escribo el posible centro de costo
        ccostoaux = ccostoaux & Mid(cuenta, pos1, 1)
        'elimino el indicador actual
        modelo = Replace(modelo, Indicador, Reemplazo, 1, 1)
        'elimino de cuenta la posicion del indicador
        cuenta = Mid(cuenta, 1, pos1 - 1) & Reemplazo & Mid(cuenta, pos1 + 1, Len(cuenta))
        'encuentro el proximo indicador
        pos1 = InStr(1, modelo, Indicador)
    Loop
    If (cuenta = modelo) Then
        cuenta_analitica = ccostoaux
        nro_cuenta = cuenta
    End If
Else
    Indicador = "E2"
    nro_cuenta = ""
    ccostoaux = ""
    pos1 = InStr(1, modelo, Indicador) 'encuentro el primer indicador
    If pos1 > 0 Then
        Do While pos1 > 0
            'escribo el posible centro de costo
            ccostoaux = ccostoaux & Mid(cuenta, pos1, 1)
            'elimino el indicador actual
            modelo = Replace(modelo, Indicador, Reemplazo, 1, 1)
            'elimino de cuenta la posicion del indicador
            cuenta = Mid(cuenta, 1, pos1 - 1) & Reemplazo & Mid(cuenta, pos1 + 1, Len(cuenta))
            'encuentro el proximo indicador
            pos1 = InStr(1, modelo, Indicador)
        Loop
        If (cuenta = modelo) Then
            cuenta_analitica = ccostoaux
            nro_cuenta = cuenta
        End If
    Else
        Indicador = "E3"
        nro_cuenta = ""
        ccostoaux = ""
        pos1 = InStr(1, modelo, Indicador) 'encuentro el primer indicador
        If pos1 > 0 Then
            Do While pos1 > 0
                'escribo el posible centro de costo
                ccostoaux = ccostoaux & Mid(cuenta, pos1, 1)
                'elimino el indicador actual
                modelo = Replace(modelo, Indicador, Reemplazo, 1, 1)
                'elimino de cuenta la posicion del indicador
                cuenta = Mid(cuenta, 1, pos1 - 1) & Reemplazo & Mid(cuenta, pos1 + 1, Len(cuenta))
                'encuentro el proximo indicador
                pos1 = InStr(1, modelo, Indicador)
            Loop
            If (cuenta = modelo) Then
                cuenta_analitica = ccostoaux
                nro_cuenta = cuenta
            End If
        Else
            Indicador = "L"
            nro_cuenta = ""
            ccostoaux = ""
            pos1 = InStr(1, modelo, Indicador) 'encuentro el primer indicador
            If pos1 > 0 Then
                Do While pos1 > 0
                    'escribo el posible centro de costo
                    ccostoaux = ccostoaux & Mid(cuenta, pos1, 1)
                    'elimino el indicador actual
                    modelo = Replace(modelo, Indicador, Reemplazo, 1, 1)
                    'elimino de cuenta la posicion del indicador
                    cuenta = Mid(cuenta, 1, pos1 - 1) & Reemplazo & Mid(cuenta, pos1 + 1, Len(cuenta))
                    'encuentro el proximo indicador
                    pos1 = InStr(1, modelo, Indicador)
                Loop
                If (cuenta = modelo) Then
                    cuenta_analitica = ccostoaux
                    nro_cuenta = cuenta
                End If
            Else
                If Len(modelo) >= 14 Then
                    cuenta_analitica = Mid(cuenta, 9, Len(cuenta))
                    nro_cuenta = Mid(cuenta, 1, 8)
                End If
            End If
        End If
    End If
End If

End Sub



Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   : 09/11/2005 - Fapitalle N. - Adecuado al proceso de exportacion formato Arballon
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim Periodo As Long
Dim Asiento As String
Dim Empresa As Long
Dim TipoArchivo As Long
Dim ProcVol As String
'Dim Minuta As String

'Orden de los parametros
'pliqnro
'Asinro, lista separada por comas
'proceso de volcado, 0=todos
'numero de minuta

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Periodo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Asiento = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        ProcVol = Mid(parametros, pos1, pos2 - pos1 + 1)
        
'        pos1 = pos2 + 2
'        pos2 = Len(parametros)
'        Minuta = CInt(Mid(parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If
'Call Generacion(bpronro, Periodo, Asiento, Empresa, ProcVol, Minuta)
'FGZ - 08/05/2006
Call Generacion(bpronro, Periodo, Asiento, Empresa, ProcVol)
End Sub


