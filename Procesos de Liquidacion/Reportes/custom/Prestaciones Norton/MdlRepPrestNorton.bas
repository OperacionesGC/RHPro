Attribute VB_Name = "MdlRepPrestNorton"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "23/07/2007"   'Diego Rosso
'Global Const UltimaModificacion = " "

'Global Const Version = "1.01"
'Global Const FechaModificacion = "24/07/2007"   'Diego Rosso
'Global Const UltimaModificacion = " " 'Se cambio cant por monto en la busqueda de dias.
'                                      'se cambio el tipo de dato para los porcentajes

Global Const Version = "1.02" ' Cesar Stankunas
Global Const FechaModificacion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

'===========================================================================================================================
Global IdUser As String
Global Fecha As Date
Global Hora As String

Public Sub Main()

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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
    
   ' carga las configuraciones basicas, formato de fecha, string de conexion,tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Nombre_Arch = PathFLog & "ReportePrestaDinearia" & "-" & NroProcesoBatch & ".log"
        
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha        = " & FechaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 175 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Call RepPlanillaAcc(NroProcesoBatch, bprcparam)
    End If
    Set rs_batch_proceso = Nothing
    
    TiempoFinalProceso = GetTickCount
    
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set objConn = Nothing
    

    Flog.writeline "----------------------------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.Close
End Sub


Public Sub RepPlanillaAcc(ByVal bpronro As Long, ByVal parametros As String)

'Para levantar parametros
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

'Variables que vienen del asp
Dim Desde As Long
Dim Hasta As Long

Dim Titulo             As String
Dim cantdias           As Integer
Dim AcuRem             As Long
Dim Acudias            As Long
Dim ExisteAcuDias      As Boolean
Dim AcumSac            As Long
Dim AcumRemTot         As Long
Dim TotalMontos        As Double
Dim TotalDias          As Long
Dim FechaAltaLaboral   As Date
Dim NroSiniestro       As Double
Dim FechaAccidente
Dim NombreApe          As String
Dim EmpleTernro        As Double
Dim NroCuil            As String
Dim Contrato           As String
Dim RazSoc             As String
Dim Cuit               As String
Dim DetImpoACT         As Double
Dim Acc_Nro            As Double
Dim EmpOSocial         As String
Dim EsAcumSacConcepto  As Boolean
Dim EsAcumRemTotConcepto As Boolean
Dim Remuneracion         As Double
Dim SAC                  As Double
Dim FechaDesdelicencia   As String
Dim FechaHastalicencia   As String

'Variables para guardar el tipo de conc/acum
Dim ConJubilacion   As Long
Dim ConINSSJP       As Long
Dim ConFNE          As Long
Dim ConSalario      As Long
Dim ConANSSAL       As Long
Dim ConOS           As Long

        
Dim tipo6  As String
Dim tipo7  As String
Dim tipo8  As String
Dim tipo9  As String
Dim tipo10 As String
Dim tipo11 As String

'Variables que guardan los diferentes porcentajes de los aportes
Dim PorcJubilacion      As Double
Dim PorcINSSJP          As Double
Dim PorcFNE             As Double
Dim PorcSalario         As Double
Dim PorcAnssal          As Double
Dim PorcOS              As Double

Dim EMPRESA As Integer ' Guarda el ternro de la empresa
Dim Orden

'Arreglos
Dim Aux_Det_Importe(12)  As Double
Dim Aux_Det_Dias(12)     As Integer
Dim Aux_Det_Anio(12)     As String
Dim Aux_Det_Mes(12)      As String

'Varibles Para detalle de las Remuneraciones
Dim I           As Integer
Dim mes         As Integer
Dim Anio        As Integer
Dim Total       As Double
Dim Cantidad    As Long
Dim Cant_Tot_Dias As Integer
Dim Cant_Dias_ART As Integer



'Registro
Dim rs_Empleado      As New ADODB.Recordset
Dim rs_Accidente     As New ADODB.Recordset
Dim rs_acumulador    As New ADODB.Recordset
Dim rs_Confrep       As New ADODB.Recordset
Dim rs_Reporte       As New ADODB.Recordset
Dim rs_Reporte1      As New ADODB.Recordset
Dim rsConsult        As New ADODB.Recordset
Dim rs_Tercero       As New ADODB.Recordset
Dim rs_Periodo       As New ADODB.Recordset
Dim rs_Acu_Liq       As New ADODB.Recordset
Dim rs_Detliq        As New ADODB.Recordset
Dim rs_CabLiq        As New ADODB.Recordset
Dim rs_Empresa       As New ADODB.Recordset
Dim rs_Estr_cod      As New ADODB.Recordset


' Levanto cada parametro por separado, el separador de parametros es "@"
Separador = "@"
Flog.writeline "Levantando los paramentros"


If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Desde = CLng(Mid(parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Hasta = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If
Flog.writeline "    Termino de Levantar los paramentros"

Titulo = "Siniestros desde: " & Desde & " Siniestros Hasta: " & Hasta & " - Al " & Fecha

On Error GoTo CE

'Busco los accidentes dentro de los numeros desde y hasta de siniestros
StrSql = "SELECT * FROM soaccidente "
StrSql = StrSql & " inner join empleado on empleado.ternro = soaccidente.empleado"
StrSql = StrSql & " WHERE accnrosiniestro >= " & Desde & " AND accnrosiniestro <= " & Hasta
OpenRecordset StrSql, rs_Reporte
If Not rs_Reporte.EOF Then
    
    'Verifico que este dado de alta el reporte
    StrSql = "Select * FROM reporte where reporte.repnro = 200"
    OpenRecordset StrSql, rs_Reporte1
    If rs_Reporte1.EOF Then
        Flog.writeline "ERROR El Reporte Numero 200 no esta dado de alta "
        Exit Sub
    End If
  
    
    
    'Configuracion del Reporte - Acum Mens
    Flog.writeline "Obteniendo el codigo del acumulador mensual "
    StrSql = "SELECT * FROM confrep WHERE repnro = 200 AND confnrocol = 1"
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If rs_Confrep.EOF Then
        Flog.writeline " FALTA configurar la columna 1 del reporte "
        Exit Sub
    Else
        AcuRem = rs_Confrep!confval
        Flog.writeline " Codigo Obtenido "
    End If
    
    
    'Configuracion del Reporte - Acumulador dias
    Flog.writeline "Obteniendo el codigo del acumulador de Dias liquidados"
    StrSql = "SELECT * FROM confrep WHERE repnro = 200 AND confnrocol = 2 "
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        Acudias = rs_Confrep!confval
        Flog.writeline "    ACumulador obtenido "
    Else
        Flog.writeline " FALTA configurar la columna 2 del reporte "
        Exit Sub
    End If
    
    'Configuracion del Reporte - SAC
    Flog.writeline "Obteniendo el codigo del acumulador de SAC. Busqueda 352 "
    StrSql = "SELECT * FROM confrep WHERE repnro = 200 AND confnrocol = 3 "
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        If rs_Confrep!conftipo = "CO" Then
            StrSql = "SELECT * FROM concepto "
            StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
            StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
            OpenRecordset StrSql, rsConsult
              If Not rsConsult.EOF Then
                    AcumSac = rsConsult!concnro
                    EsAcumSacConcepto = True
                    Flog.writeline "    Configuracion Obtenida. "
              Else
                    AcumSac = 0
                    Flog.writeline "    No se encontro el concepto. "
                    Exit Sub
              End If
        Else
             StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
             OpenRecordset StrSql, rsConsult
             If Not rsConsult.EOF Then
                    AcumSac = rsConsult!acunro
                    EsAcumSacConcepto = False
                    Flog.writeline "    Configuracion Obtenida. "
             Else
                    AcumSac = 0
                    Flog.writeline "    No se encontro el acumulador "
                    Exit Sub
             End If
        End If
    Else
        Flog.writeline " FALTA configurar la columna 3 del reporte "
        Exit Sub
    End If
    
     'Configuracion del Reporte - Remuneraciones totales
    Flog.writeline "Obteniendo el codigo del acumulador de Remuneraciones totales sujetas a cotizaciòn. Busqueda 449"
    StrSql = "SELECT * FROM confrep WHERE repnro = 200 AND confnrocol = 4 "
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        If rs_Confrep!conftipo = "CO" Then
            StrSql = "SELECT * FROM concepto "
            StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
            StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
            OpenRecordset StrSql, rsConsult
              If Not rsConsult.EOF Then
                    AcumRemTot = rsConsult!concnro
                    EsAcumRemTotConcepto = True
                    Flog.writeline "    Configuracion Obtenida. "
              Else
                    AcumRemTot = 0
                    Flog.writeline "    No se encontro el concepto. "
                    Exit Sub
              End If
        Else
             StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
             OpenRecordset StrSql, rsConsult
             If Not rsConsult.EOF Then
                    AcumRemTot = rsConsult!acunro
                    EsAcumRemTotConcepto = False
                    Flog.writeline "    Configuracion Obtenida. "
             Else
                    AcumRemTot = 0
                    Flog.writeline "    No se encontro el acumulador "
                    Exit Sub
             End If
        End If
    Else
        Flog.writeline " FALTA configurar la columna 4 del reporte "
        Exit Sub
    End If
    
    
    'seteo de las variables de progreso
    Progreso = 0
    IncPorc = (100 / rs_Reporte.RecordCount)
    
    Orden = 0
    
    Do While Not rs_Reporte.EOF
        Orden = Orden + 1
        NroSiniestro = rs_Reporte!accnrosiniestro
        FechaAccidente = rs_Reporte!accfecha
        NombreApe = rs_Reporte!terape & ", " & rs_Reporte!ternom
        EmpleTernro = rs_Reporte!ternro
        Acc_Nro = rs_Reporte!accnro
        
        Flog.writeline "Buscando datos del Empleado: " & NombreApe & "(" & EmpleTernro & ")"
        Flog.writeline "--------------------------------------------------------------------"
        Flog.writeline
        
        'Busco el cuil del empleado
        Flog.writeline "  Obteniendo el Cuil del Empleado "
        StrSql = "SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc "
        StrSql = StrSql & " Where ternro = " & EmpleTernro
        StrSql = StrSql & " AND   tidnro = 10"
        OpenRecordset StrSql, rs_Tercero
    
        If Not rs_Tercero.EOF Then
            NroCuil = rs_Tercero!nrodoc
            Flog.writeline "Se obtuvo el cuil del Empleado"
        Else
            Flog.writeline "El Empleado no tiene un Cuil asignado"
            NroCuil = ""
        End If
        
        'Busco el valor de la estructura Contrato del empleado.
        StrSql = " SELECT estrdabr "
        StrSql = StrSql & " From his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Now) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(Now) & ") And his_estructura.tenro = 18 And his_estructura.ternro = " & EmpleTernro
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           Contrato = rsConsult!estrdabr
           Flog.writeline "Se obtuvieron los datos de contrato"
        Else
           Contrato = ""
           Flog.writeline "Error al obtener los datos de contrato"
        End If
        
      '-----------------------------------------------------------------------------------------------
        'Busco el Codigo de la obra social
        Flog.writeline "Busco el Codigo de la obra social"
        StrSql = " SELECT estructura.estrnro "
        StrSql = StrSql & " From estructura"
        StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Now) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(Now) & ") And his_estructura.tenro = 17 And his_estructura.ternro = " & EmpleTernro
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
            StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rsConsult!Estrnro
            StrSql = StrSql & " AND tcodnro = 1"
            OpenRecordset StrSql, rs_Estr_cod
            If Not rs_Estr_cod.EOF Then
                EmpOSocial = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 15))
            Else
                EmpOSocial = ""
                Flog.writeline " No se encontró el codigo "
            End If
        Else
           EmpOSocial = ""
           Flog.writeline "Error al obtener los datos "
        End If
      '-----------------------------------------------------------------------------------------------
        
    
      '-----------------------------------------------------------------------------------------------
      'Busco las licencias que estan asignadas al accidente.
        Flog.writeline "Busco las licencias que estan asignadas al accidente"
        
        StrSql = " SELECT *  "
        StrSql = StrSql & " FROM lic_accid "
        StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_accid.emp_licnro  "
        StrSql = StrSql & " WHERE lic_accid.accnro = " & Acc_Nro
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            FechaDesdelicencia = rsConsult!elfechadesde
            FechaHastalicencia = rsConsult!elfechahasta
            Flog.writeline "Dato obtenido"
            
        Else
            Flog.writeline "No se encontro una licencia asociada al accidente"
            FechaDesdelicencia = ""
            FechaHastalicencia = ""
        End If
      '----------------------------------------------------------------------------------------------
    
    
        
        'Busco la empresa a la que pertenece el empleado
       
        Flog.writeline "Buscando datos de la empresa a la cual pertenece el empleado "
           
        StrSql = "select tercero.ternro,terrazsoc,nrodoc,htetdesde "
        StrSql = StrSql & " From his_estructura "
        StrSql = StrSql & " INNER JOIN empresa ON his_estructura.estrnro = empresa.estrnro and his_estructura.tenro = 10"
        StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empresa.ternro"
        StrSql = StrSql & " INNER JOIN ter_doc on ter_doc.ternro = tercero.ternro  AND tidnro = 6"
        StrSql = StrSql & " Where his_estructura.ternro=" & EmpleTernro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Now)
        StrSql = StrSql & " AND (htethasta >= " & ConvFecha(Now) & " or htethasta is null)"
        
        OpenRecordset StrSql, rs_Empresa
        If Not rs_Empresa.EOF Then
           RazSoc = rs_Empresa!terrazsoc 'Razon Social Empresa
           Cuit = rs_Empresa!nrodoc 'CUIT
           EMPRESA = rs_Empresa!ternro
           Flog.writeline "  Datos de la empresa obtenidos "
        Else
           RazSoc = ""
           Cuit = ""
           EMPRESA = 0
           Flog.writeline "El Registro de Empresa no esta  disponible"
        End If
    
       
        Call Columna(6, ConJubilacion, tipo6)
        Call Columna(7, ConINSSJP, tipo7)
        Call Columna(8, ConFNE, tipo8)
        Call Columna(9, ConSalario, tipo9)
        Call Columna(10, ConANSSAL, tipo10)
        Call Columna(11, ConOS, tipo11)

        '**************************************************************************************************
        '************************************ DETALLE DE LAS REMUNERACIONES *******************************
        Flog.writeline "OBTENGO EL DETALLE DE LAS REMUNERACIONES"
        '**************************************************************************************************

        'Inicializo las varibles
        PorcJubilacion = 0
        PorcINSSJP = 0
        PorcFNE = 0
        PorcSalario = 0
        PorcAnssal = 0
        PorcOS = 0
        DetImpoACT = 0
        Remuneracion = 0
        SAC = 0
                    
        Anio = Year(rs_Reporte!accfecha)
        mes = Month(rs_Reporte!accfecha)
     
        'Obtengo todos los datos del periodo
        StrSql = "SELECT * FROM periodo "
        StrSql = StrSql & " WHERE pliqanio = " & Anio
        StrSql = StrSql & " AND pliqmes = " & mes
        OpenRecordset StrSql, rs_Periodo
        Flog.writeline "    Obteniendo datos del periodo "
        
        If Not rs_Periodo.EOF Then
            Flog.writeline "      DATOS OBTENIDOS "
         
        'Levanto el valor del AC/CO de la busqueda 352
        If EsAcumSacConcepto = True Then
            
            StrSql = "SELECT * FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " WHERE detliq.concnro = " & AcumSac
            StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
            OpenRecordset StrSql, rsConsult
            
            If Not rsConsult.EOF Then
                SAC = rsConsult!dlimonto
            Else
                SAC = 0
                Flog.writeline "Error no se encontraron los datos del SAC"
            End If
            
        Else
        
            StrSql = "SELECT almonto FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " WHERE acu_liq.acunro = " & AcumSac
            StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
            OpenRecordset StrSql, rsConsult
                
            If Not rsConsult.EOF Then
                SAC = rsConsult!almonto
            Else
                SAC = 0
                Flog.writeline "Error no se encontraron los datos del SAC"
            End If
        
        End If
        
        
        'Levanto el valor del AC/CO de la busqueda 449
        If EsAcumRemTotConcepto = True Then
            
            StrSql = "SELECT * FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " WHERE detliq.concnro = " & AcumRemTot
            StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
            OpenRecordset StrSql, rsConsult
            
            If Not rsConsult.EOF Then
                Remuneracion = rsConsult!dlimonto
            Else
                Remuneracion = 0
                Flog.writeline "Error no se encontraron los datos de la Remuneracion total"
            End If
            
        Else
        
            StrSql = "SELECT almonto FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " WHERE acu_liq.acunro = " & AcumRemTot
            StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
            OpenRecordset StrSql, rsConsult
                
            If Not rsConsult.EOF Then
                Remuneracion = rsConsult!almonto
            Else
                Remuneracion = 0
                Flog.writeline "Error no se encontraron los datos del Remuneracion total"
            End If
        End If
        
           
           
            '----------BUSCO LOS PORCENTAJES DE APORTES-----------------------------
            Flog.writeline "Busco la Contribucion Patronal "
                 ' Busco el porcentaje de jubilacion del empleado
                                
                    If tipo6 = "CO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConJubilacion
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcJubilacion = rs_Detliq!dlimonto
                           rs_Detliq.MoveNext
                        Loop
                   End If
                   If tipo6 = "PCO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConJubilacion
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcJubilacion = rs_Detliq!dlicant
                           rs_Detliq.MoveNext
                        Loop
                    End If
                    
                                   
                 ' Busco el porcentaje de Obra Social del empleado
                    If tipo7 = "CO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConINSSJP
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcINSSJP = rs_Detliq!dlimonto
                           rs_Detliq.MoveNext
                        Loop
                    End If
                    If tipo7 = "PCO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConINSSJP
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcINSSJP = rs_Detliq!dlicant
                           rs_Detliq.MoveNext
                        Loop
                    End If
                    
                 
                 ' Busco el porcentaje de LEY del empleado
                 
                    If tipo8 = "CO" Then
                       StrSql = "SELECT * FROM proceso "
                       StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                       StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                       StrSql = StrSql & " WHERE detliq.concnro = " & ConFNE
                       StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                       StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                       OpenRecordset StrSql, rs_Detliq
                    
                       Do While Not rs_Detliq.EOF
                          PorcFNE = rs_Detliq!dlimonto
                          rs_Detliq.MoveNext
                       Loop
                    End If
                    If tipo8 = "PCO" Then
                       StrSql = "SELECT * FROM proceso "
                       StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                       StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                       StrSql = StrSql & " WHERE detliq.concnro = " & ConFNE
                       StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                       StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                       OpenRecordset StrSql, rs_Detliq
                    
                       Do While Not rs_Detliq.EOF
                          PorcFNE = rs_Detliq!dlicant
                          rs_Detliq.MoveNext
                       Loop
                    End If
                  
                 ' Busco el porcentaje de Salario Familiar
            
                    If tipo9 = "CO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConSalario
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcSalario = rs_Detliq!dlimonto
                           rs_Detliq.MoveNext
                        Loop
                    End If
                    If tipo9 = "PCO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConSalario
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcSalario = rs_Detliq!dlicant
                           rs_Detliq.MoveNext
                        Loop
                    End If
              
                  
                  ' Busco el porcentaje de ANSSAL
            
                    If tipo10 = "CO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConANSSAL
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcAnssal = rs_Detliq!dlimonto
                           rs_Detliq.MoveNext
                        Loop
                    End If
                    If tipo10 = "PCO" Then
                        StrSql = "SELECT * FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                        StrSql = StrSql & " WHERE detliq.concnro = " & ConANSSAL
                        StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                        OpenRecordset StrSql, rs_Detliq
                        
                        Do While Not rs_Detliq.EOF
                           PorcAnssal = rs_Detliq!dlicant
                           rs_Detliq.MoveNext
                        Loop
                    End If
             
                  
                 ' Busco el porcentaje de Obra Social
                 
                    If tipo11 = "CO" Then
                       StrSql = "SELECT * FROM proceso "
                       StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                       StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                       StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                       StrSql = StrSql & " WHERE detliq.concnro = " & ConOS
                       StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                       OpenRecordset StrSql, rs_Detliq
                    
                       Do While Not rs_Detliq.EOF
                          PorcOS = rs_Detliq!dlimonto
                          rs_Detliq.MoveNext
                       Loop
                    End If
    
                    If tipo11 = "PCO" Then
                       StrSql = "SELECT * FROM proceso "
                       StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                       StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                       StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                       StrSql = StrSql & " WHERE detliq.concnro = " & ConOS
                       StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                       OpenRecordset StrSql, rs_Detliq
                    
                       Do While Not rs_Detliq.EOF
                          PorcOS = rs_Detliq!dlicant
                          rs_Detliq.MoveNext
                       Loop
                     End If
                 
                 
              
        Else
             Flog.writeline "      No se encotraron liquidaciones para el empleado en el mes del accidente "
             
        End If
     
     
        Flog.writeline "Busco las remuneraciones para los 12 meses "

       
         
        'Inicializo las variables
        I = 1
        Total = 0
        Cantidad = 0
        TotalDias = 0
        Do While I <= 12
         
            'Obtengo todos los datos del periodo
            StrSql = "SELECT * FROM periodo "
            StrSql = StrSql & " WHERE pliqanio = " & Anio
            StrSql = StrSql & " AND pliqmes = " & mes
            OpenRecordset StrSql, rs_Periodo
            Flog.writeline "    Obteniendo datos del periodo:" & mes & " " & Anio
            
            If Not rs_Periodo.EOF Then
                Flog.writeline "      DATOS OBTENIDOS "
                'Busco el importe para el acumulador AcumRem
                StrSql = "SELECT almonto FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE acu_liq.acunro = " & AcuRem
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                OpenRecordset StrSql, rs_Acu_Liq
                
                Do While Not rs_Acu_Liq.EOF
                    Total = Total + rs_Acu_Liq!almonto
                    rs_Acu_Liq.MoveNext
                Loop
             
                'Busco la cantidad para el acumulador AcumDias
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE acu_liq.acunro = " & Acudias
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                StrSql = StrSql & " AND cabliq.empleado =" & EmpleTernro
                OpenRecordset StrSql, rs_Acu_Liq
            
                Do While Not rs_Acu_Liq.EOF
                  ' Cantidad = Cantidad + rs_Acu_Liq!alcant
                   Cantidad = Cantidad + rs_Acu_Liq!almonto
                   rs_Acu_Liq.MoveNext
                  
                Loop
               
            End If 'rs_periodo.eof
        
            Aux_Det_Importe(I) = Total
            Aux_Det_Anio(I) = CStr(Anio)
            Aux_Det_Mes(I) = mes
      
            Anio = IIf(mes = 1, Anio - 1, Anio)
            mes = IIf(mes = 1, 12, mes - 1)
            'TotalMontos = TotalMontos + Total 'Sumatoria de los ultimos 12 sueldos
            TotalDias = TotalDias + Cantidad 'Cantidad de dias corridos ultimos 12 meses
            Total = 0
            Cantidad = 0
            I = I + 1
            ExisteAcuDias = False
        Loop


        
        
        'INSERTO EN LA BD

        StrSql = "INSERT INTO Rep_prestdine (repnro, empresa, bpronro, Fecha, Hora, iduser, ternro, empdor_razsoc, empdor_cuit, emp_apeynom, emp_cuil, "
        StrSql = StrSql & " emp_osocial, emp_afjp, acc_nro, acc_fecha, acc_nrosiniestro, "
        StrSql = StrSql & " det_mes_1, det_mes_2, det_mes_3, det_mes_4, det_mes_5, "
        StrSql = StrSql & " det_mes_6, det_mes_7, det_mes_8, det_mes_9, det_mes_10, det_mes_11, det_mes_12, det_anio_1, det_anio_2, det_anio_3, det_anio_4, det_anio_5,"
        StrSql = StrSql & " det_anio_6, det_anio_7, det_anio_8, det_anio_9, det_anio_10, det_anio_11, det_anio_12, det_importe_1, det_importe_2, det_importe_3,"
        StrSql = StrSql & " det_importe_4, det_importe_5, det_importe_6, det_importe_7, det_importe_8, det_importe_9, det_importe_10, det_importe_11, det_importe_12,  "
      '  StrSql = StrSql & " det_dias_1, det_dias_2, det_dias_3, det_dias_4, det_dias_5, det_dias_6, det_dias_7, det_dias_8, det_dias_9, det_dias_10, det_dias_11, det_dias_12, "
        StrSql = StrSql & "  porcjub, porcinssjp, porcfondonac, porcasigfam, porcanssal, porcos, prestdine1, ampomin, ampomax, CantTotDias,acc_fecreintdesde,acc_fecreinthasta, emp_tel "
        StrSql = StrSql & " ) VALUES ("

         StrSql = StrSql & 200 & ","
         StrSql = StrSql & EMPRESA & ","
         StrSql = StrSql & bpronro & ","
         StrSql = StrSql & ConvFecha(Fecha) & ","
         StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
         StrSql = StrSql & "'" & IdUser & "',"
         StrSql = StrSql & EmpleTernro & ","
         StrSql = StrSql & "'" & RazSoc & "',"
         StrSql = StrSql & "'" & Cuit & "',"
         StrSql = StrSql & "'" & NombreApe & "',"
         StrSql = StrSql & "'" & NroCuil & "',"
         StrSql = StrSql & "'" & Contrato & "',"
         StrSql = StrSql & "'" & Titulo & "',"
         StrSql = StrSql & Acc_Nro & ","  'Numero de accidente
         StrSql = StrSql & ConvFecha(FechaAccidente) & ","  'Fecha del Accidente
         StrSql = StrSql & NroSiniestro & ","
         
         For I = 1 To 12
            If IsNull(Aux_Det_Mes(I)) Then
                  StrSql = StrSql & "'',"
            Else
                  StrSql = StrSql & "'" & NombreMes(Aux_Det_Mes(I)) & "',"
            End If
         Next I
         For I = 1 To 12
            If IsNull(Aux_Det_Anio(I)) Then
                  StrSql = StrSql & "'',"
            Else
                  StrSql = StrSql & "'" & Aux_Det_Anio(I) & "',"
            End If
         Next I
         For I = 1 To 12
            If IsNull(Aux_Det_Importe(I)) Then
                  StrSql = StrSql & "null,"
            Else
                  StrSql = StrSql & Aux_Det_Importe(I) & ","
            End If
         Next I

    '    For I = 1 To 12
    '        If IsNull(Aux_Det_Dias(I)) Then
    '            StrSql = StrSql & "null,"
    '        Else
    '            StrSql = StrSql & Aux_Det_Dias(I) & ","
    '        End If
    '    Next I

        StrSql = StrSql & PorcJubilacion & ","  'Porcentaje de jubilacion
        StrSql = StrSql & PorcINSSJP & ","  'Porcentaje INNSJP
        StrSql = StrSql & PorcFNE & ","  'Porcentaje FNE
        StrSql = StrSql & PorcSalario & "," 'Porcentaje de Salario
        StrSql = StrSql & PorcAnssal & "," 'Porcentaje de ANSSAL
        StrSql = StrSql & PorcOS & "," 'Porcentaje de Obra social
        StrSql = StrSql & Orden & "," 'Posicion en el reporte
        StrSql = StrSql & SAC & "," 'busqueda 352
        StrSql = StrSql & Remuneracion & "," 'Busqueda 449
        StrSql = StrSql & TotalDias & "," 'TotalDias
        
        If FechaDesdelicencia = "" Then
            StrSql = StrSql & "null,"
        Else
            StrSql = StrSql & ConvFecha(CDate(FechaDesdelicencia)) & ","
        End If
        
        If FechaHastalicencia = "" Then
            StrSql = StrSql & "null,"
        Else
            StrSql = StrSql & ConvFecha(CDate(FechaHastalicencia)) & ","
        End If
        StrSql = StrSql & EmpOSocial 'codigo de obra social
        StrSql = StrSql & ")"

        Flog.writeline "------------------------------------------------------"
        Flog.writeline "Sql ejecutado: " & StrSql
        Flog.writeline "------------------------------------------------------"
        Flog.writeline
        objConn.Execute StrSql, , adExecuteNoRecords


            'Fin de la transaccion
            MyCommitTrans
            
            Flog.writeline "Se inserto el registro"
            Flog.writeline

            'Actualizo el progreso del Proceso
            Progreso = Progreso + IncPorc
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                     "' WHERE bpronro = " & NroProcesoBatch
           objConn.Execute StrSql, , adExecuteNoRecords
                
                
            rs_Reporte.MoveNext
    
    Loop
    

            
            If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
            If rs_Accidente.State = adStateOpen Then rs_Accidente.Close
            If rs_acumulador.State = adStateOpen Then rs_acumulador.Close
            If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
            If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
            If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
            If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
            If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
            If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
            If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
            If rs_CabLiq.State = adStateOpen Then rs_CabLiq.Close
            
            Set rs_Empleado = Nothing
            Set rs_Accidente = Nothing
            Set rs_acumulador = Nothing
            Set rs_Confrep = Nothing
            Set rs_Reporte = Nothing
            Set rs_Empresa = Nothing
            Set rs_Tercero = Nothing
            Set rs_Periodo = Nothing
            Set rs_Acu_Liq = Nothing
            Set rs_Detliq = Nothing
            Set rs_CabLiq = Nothing
            
            
            
Exit Sub
CE:
HuboError = True
MyRollbackTrans
Flog.writeline "==================================================================================="
Flog.writeline Err.Description
Flog.writeline "Ultimo SQL Ejecutado:"
Flog.writeline StrSql
Flog.writeline "==================================================================================="
Exit Sub
            
                
                

Else
   Flog.writeline "  No se encontraron Accidentes que cumplan con las condiciones "
   
End If

End Sub


Public Sub Columna(ByVal NroCol As Integer, ByRef nroconc As Long, ByRef TipoNro As String)
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_acumulador As New ADODB.Recordset

Flog.writeline "Obteniendo la configuracion de la  Columna numero: " & NroCol
On Error GoTo CE

    StrSql = "SELECT * FROM confrep WHERE repnro = 200 AND confnrocol = " & NroCol
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        Select Case rs_Confrep!conftipo
            Case "CO":
                 StrSql = "SELECT * FROM concepto "
                 StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                 StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                OpenRecordset StrSql, rs_Concepto
                If Not rs_Concepto.EOF Then
                    nroconc = rs_Concepto!concnro
                    TipoNro = "CO"
                    Flog.writeline "    Configuracion Obtenida. Tipo columna CO"
                End If
            Case "AC":
                StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
                OpenRecordset StrSql, rs_acumulador
                If Not rs_acumulador.EOF Then
                    nroconc = rs_acumulador!acunro
                    TipoNro = "AC"
                    Flog.writeline "    Configuracion Obtenida. Tipo columna AC"
                End If
            Case "PCO":
                 StrSql = "SELECT * FROM concepto "
                 StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                 StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                OpenRecordset StrSql, rs_Concepto
                If Not rs_Concepto.EOF Then
                    nroconc = rs_Concepto!concnro
                    TipoNro = "PCO"
                    Flog.writeline "    Configuracion Obtenida. Tipo columna PCO"
                End If
         End Select
    End If
    
Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
End Sub


Public Function cantdias_mes(ByVal nromes)

Select Case nromes
    Case 1, 3, 5, 7, 8, 10, 12:
        cantdias_mes = 31
    Case 4, 6, 9, 11:
        cantdias_mes = 30
    Case 2:
        cantdias_mes = 28
End Select

End Function


Public Sub Calcular_Dias(ByVal fechadesdefiltro As Date, ByVal fechahastafiltro As Date, ByVal fdesde As Date, ByVal fhasta As Date, ByRef DiasPeriodoAnt As Long, ByRef DiasPeriodoAct As Long, ByRef DiasPeriodosig As Long, ByVal elcantdias As Integer)

DiasPeriodoAnt = 0
DiasPeriodoAct = 0
DiasPeriodosig = 0

       
        If fechadesdefiltro > fhasta Then   'Cuando la licencia cae toda antes del periodo del filtro
                  DiasPeriodoAnt = elcantdias
        Else
            If fechahastafiltro < fdesde Then 'Cuando la licencia cae toda despues del periodo del filtro
                 
                     DiasPeriodosig = elcantdias
            Else
                If (fechadesdefiltro > fdesde) Then 'Cuando la licencia empezo antes del periodo
                         
                         If (fechahastafiltro > fhasta) Then
                            DiasPeriodoAnt = DateDiff("d", fdesde, fechadesdefiltro)
                            DiasPeriodoAct = DateDiff("d", fechadesdefiltro, fhasta)
                         Else
                            DiasPeriodoAnt = DateDiff("d", fdesde, fechadesdefiltro)
                            DiasPeriodoAct = cantdias_mes(Month(fechadesdefiltro))
                            DiasPeriodosig = DateDiff("d", fechahastafiltro, fhasta)
                         End If
                        
                Else
                         If (fechahastafiltro > fhasta) Then
                             DiasPeriodoAct = elcantdias
                         Else
                            DiasPeriodoAct = DateDiff("d", fdesde, fechahastafiltro)
                            DiasPeriodosig = DateDiff("d", fechahastafiltro, fhasta)
                         End If
                End If 'Cuando la licencia empezo antes del periodo
        End If 'Cuando la licencia cae toda despues del periodo del filtro
       End If 'Cuando la licencia cae toda antes del periodo del filtro

End Sub


Public Function NombreMes(nro)
   Dim mes
   Select Case nro
     Case 1
        mes = "Enero"
     Case 2
        mes = "Febrero"
     Case 3
        mes = "Marzo"
     Case 4
        mes = "Abril"
     Case 5
        mes = "Mayo"
     Case 6
        mes = "Junio"
     Case 7
        mes = "Julio"
     Case 8
        mes = "Agosto"
     Case 9
        mes = "Septiembre"
     Case 10
        mes = "Octubre"
     Case 11
        mes = "Noviembre"
     Case 12
        mes = "Diciembre"
   End Select
   NombreMes = mes
End Function 'NombreMes(nro)

