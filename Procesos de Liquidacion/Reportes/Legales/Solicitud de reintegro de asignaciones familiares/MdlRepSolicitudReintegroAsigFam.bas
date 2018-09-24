Attribute VB_Name = "MdlRepSolicitudReintegroAsigFam"
Option Explicit

'Version: 0.01
'Primera Version realizada, todavia en etapa de desarrollo
Const Version = 1.1
Const FechaVersion = "05/05/2006"


'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Aux_Autoriz_Apenom As String
Global Aux_Autoriz_Docu As String
Global Aux_Autoriz_Prov_Emis As String

Global Aux_Certifi_Corresponde As String
Global Aux_Certifi_Doc_Tipo As String
Global Aux_Certifi_Doc_Nro As String
Global Aux_Certifi_Expedida As String



Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial del Generador de Reportes.
    ' Autor      : FGZ
    ' Fecha      : 17/02/2004
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim objconnMain As New ADODB.Connection
    Dim strCmdLine
    Dim Nombre_Arch As String
    Dim HuboError As Boolean
    Dim rs_batch_proceso As New ADODB.Recordset
    Dim bprcparam As String
    Dim PID As String
    Dim ArrParametros
    strCmdLine = Command()
   ' strCmdLine = "10209 "
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

    'Abro la conexion
    OpenConnection strconexion, objConn
    
    Nombre_Arch = PathFLog & "SolicitudReintegroAsigFam_PS32" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE (btprcnro = 131 ) AND bpronro =" & NroProcesoBatch
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
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.Close
    objConn.Close
    
End Sub

Public Sub Generar_Reporte(ByVal bpronro As Long, ByVal HFecha As Date, ByVal p_Empresa_ternro As Integer, ByVal p_Mes As Integer, ByVal p_Anio As Integer, ByVal p_CtaBancaria As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Certificado Anses de Servicios
' Autor      : FGZ
' Fecha      : 15/07/2004
' Ult. Mod   : FGZ - 10/11/2005
' Desc       : agrupa en hojas de hasta 5 movimientos de empresas por c/u.
' --------------------------------------------------------------------------------------------

Dim repnro As Integer

'Datos de la empresa
Dim mes As Integer
Dim Anio As Integer
Dim cuit As String
Dim ctaBancaria As String
Dim razonsocial As String
Dim actividad As String
Dim domcalle As String
Dim domnro As String
Dim dompiso As String
Dim domdepto As String
Dim domcodpost As String
Dim domlocalidad As String
Dim domprovinc As String
Dim telefono As String
Dim ddn As String
Dim cbu As String
Dim saldo As Long
Dim sonpesos As String
Dim banco As String
Dim primera As Boolean
Dim contadorError As Integer

'Asignacion familiar Total periodo
Dim acu1 As Integer
Dim aftotalper  As Single

'Asignacion familiar Retroactivo
Dim acu2 As Integer
Dim afretroactivo  As Single

'Asignacion familiar Maternidad
Dim acu3 As Integer
Dim afmaternidad As Single

'Imponibles - Prevision
Dim acu4 As Integer
Dim imponible4 As Single
Dim imponible4_importe As Single

'Imponibles - Asig. Familiares
Dim acu5 As Integer
Dim imponible5 As Single
Dim imponible5_importe As Single

'Imponibles - Reg. Nac. de empleo
Dim acu6 As Integer
Dim imponible6 As Long
Dim imponible6_importe As Long

'Acumulador
Dim acu7 As Integer
Dim imponible7 As Long
Dim imponible7_importe As Long

'Asignacion familiar Total
Dim aftotal As Single

'Asignaciones compensables Total
Dim actotal As Single
Dim primero As Boolean
Dim cambio As Double
Dim Valor As Double
Dim dlcant As Double
Dim calculo As Double

'Inicializo variables de auxiliares de empresa
mes = p_Mes
Anio = p_Anio
cuit = " "
ctaBancaria = p_CtaBancaria
razonsocial = " "
actividad = " "
domcalle = " "
domnro = " "
dompiso = " "
domdepto = " "
domcodpost = " "
domlocalidad = " "
domprovinc = " "
telefono = " "
ddn = " "
cbu = " "
banco = " "

actotal = 0
saldo = 0
sonpesos = " "

acu1 = 0
aftotalper = 0

acu2 = 0
afretroactivo = 0

acu3 = 0
afmaternidad = 0

acu4 = 0
imponible4 = 0
imponible4_importe = 0

acu5 = 0
imponible5 = 0
imponible5_importe = 0

acu6 = 0
imponible6 = 0
imponible6_importe = 0

acu7 = 0
imponible7 = 0
imponible7_importe = 0

saldo = 0
aftotal = 0


'Registros
Dim rs_Empresa As New ADODB.Recordset
Dim rs_ctaBancaria As New ADODB.Recordset
Dim rs_aftotalper As New ADODB.Recordset
Dim rs_afretroactivo As New ADODB.Recordset
Dim rs_afmaternidad As New ADODB.Recordset
Dim rs_afmatrimonio As New ADODB.Recordset
Dim rs_acumuladores As New ADODB.Recordset
Dim rs_repaux As New ADODB.Recordset
Dim rs_imponible As New ADODB.Recordset
Dim strsql1 As String
Dim lcambio As Boolean

On Error GoTo CE

' Comienzo la transaccion
    MyBeginTrans
        
        
        Flog.writeline "Procesando empresa id " & p_Empresa_ternro
        Flog.writeline "------------------------------"
        Flog.writeline "  "
    
        'Armo el query para pedir los datos de le empresa
        StrSql = "select  empresa.empnom, ter_doc.nrodoc, detdom.calle, detdom.nro, detdom.piso, "
        StrSql = StrSql & " detdom.oficdepto , detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, empresa.empactiv "
        StrSql = StrSql & " from empresa "
        StrSql = StrSql & " left join cabdom on empresa.ternro = cabdom.ternro "
        StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro "
        StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro "
        StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro "
        StrSql = StrSql & " left join ter_doc on ter_doc.ternro = empresa.ternro and tidnro = 6 "
        StrSql = StrSql & " where empresa.empnro = " & p_Empresa_ternro
        OpenRecordset StrSql, rs_Empresa
        If rs_Empresa.EOF Then
            Flog.writeline "---------------------  "
            Flog.writeline "Empresa no existente."
            Flog.writeline "---------------------  "
            MyRollbackTrans
            Exit Sub
        End If
        If Not rs_Empresa.EOF Then
            If Not IsNull(rs_Empresa!nrodoc) Then cuit = rs_Empresa!nrodoc
            If Not IsNull(rs_Empresa!empnom) Then razonsocial = rs_Empresa!empnom
            If Not IsNull(rs_Empresa!calle) Then domcalle = rs_Empresa!calle
            If Not IsNull(rs_Empresa!nro) Then domnro = rs_Empresa!nro
            If Not IsNull(rs_Empresa!piso) Then dompiso = rs_Empresa!piso
            If Not IsNull(rs_Empresa!oficdepto) Then domdepto = rs_Empresa!oficdepto
            If Not IsNull(rs_Empresa!codigopostal) Then domcodpost = rs_Empresa!codigopostal
            If Not IsNull(rs_Empresa!locdesc) Then domlocalidad = rs_Empresa!locdesc
            If Not IsNull(rs_Empresa!provdesc) Then domprovinc = rs_Empresa!provdesc
            If Not IsNull(rs_Empresa!telnro) Then telefono = rs_Empresa!telnro
            If Not IsNull(rs_Empresa!empactiv) Then actividad = rs_Empresa!empactiv
            Flog.writeline "Creadas las variables auxiliares de la empresa"
            Flog.writeline "  "
        End If
        If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
        
        If p_CtaBancaria = 0 Then
            HuboError = True
            Flog.writeline "-----------------------------"
            Flog.writeline "Cuenta bancaria no existente."
            Flog.writeline "-----------------------------"
            MyRollbackTrans
            Exit Sub
        End If
        
        'Traigo los datos de la cuenta bancaria
        StrSql = "select ctabancaria.ctabcbu, banco.bandesc "
        StrSql = StrSql & " from ctabancaria "
        StrSql = StrSql & " join banco on banco.ternro = ctabancaria.banco and ctabancaria.cbnro = " & p_CtaBancaria
        OpenRecordset StrSql, rs_ctaBancaria
        If Not rs_ctaBancaria.EOF Then
            If Not IsNull(rs_ctaBancaria!ctabcbu) Then cbu = rs_ctaBancaria!ctabcbu
            If Not IsNull(rs_ctaBancaria!bandesc) Then banco = rs_ctaBancaria!bandesc
            Flog.writeline "Creadas las variables auxiliares de la cuenta bancaria"
            Flog.writeline "  "
        End If
        If rs_ctaBancaria.State = adStateOpen Then rs_ctaBancaria.Close
        
        
        
        'Traigo configuracion del confrep
        StrSql = "select confval,confnrocol from confrep where repnro=162"
        OpenRecordset StrSql, rs_acumuladores
        If rs_acumuladores.EOF Then
            Flog.writeline "------------------------------------------"
            Flog.writeline "No se ha configurado el reporte. (ConfRep)"
            Flog.writeline "------------------------------------------"
            MyRollbackTrans
            Exit Sub
        End If
        Do While Not rs_acumuladores.EOF
            Select Case CInt(rs_acumuladores!confnrocol)
                Case 1
                    acu1 = rs_acumuladores!confval
                Case 2
                    acu2 = rs_acumuladores!confval
                Case 3
                    acu3 = rs_acumuladores!confval
                Case 4
                    acu4 = rs_acumuladores!confval
                Case 10
                    acu5 = rs_acumuladores!confval
                Case 16
                    acu6 = rs_acumuladores!confval
                Case 22
                    acu7 = rs_acumuladores!confval
            End Select
            rs_acumuladores.MoveNext
        Loop
        Flog.writeline "  obtenidos los acumuladores del ConfRep"
        Flog.writeline "  "
        rs_acumuladores.Close
        
        'Asignacion familiar Total periodo
        StrSql = "select sum(acu_mes.ammonto) acu "
        StrSql = StrSql & " from his_estructura "
        StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
        StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu1
        OpenRecordset StrSql, rs_aftotalper
        
        If rs_aftotalper.EOF Then
            HuboError = True
            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 1"
            MyRollbackTrans
            Exit Sub
        End If
        
        If Not rs_aftotalper.EOF Then
            If Not IsNull(rs_aftotalper!acu) Then aftotalper = rs_aftotalper!acu
            Flog.writeline " -- obtenido el acumulador de periodo total de asignacion familiar"
        End If
        rs_aftotalper.Close
        
        'Asignacion familiar Retroactivo-
        StrSql = "select sum(acu_mes.ammonto) acu "
        StrSql = StrSql & " from his_estructura "
        StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
        StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu2
        OpenRecordset StrSql, rs_afretroactivo
        If rs_afretroactivo.EOF Then
            HuboError = True
            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 2"
            MyRollbackTrans
            Exit Sub
        End If
        
        If Not rs_afretroactivo.EOF Then
            If Not IsNull(rs_afretroactivo!acu) Then afretroactivo = rs_afretroactivo!acu
            Flog.writeline " -- obtenido el acumulador de Asignacion familiar Retroactivo"
        End If
        rs_afretroactivo.Close
        
        
        'Asignacion familiar Maternidad
        StrSql = "select sum(acu_mes.ammonto) acu "
        StrSql = StrSql & " from his_estructura "
        StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
        StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu3
        OpenRecordset StrSql, rs_afmatrimonio
        If rs_afmatrimonio.EOF Then
            HuboError = True
            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 3"
            MyRollbackTrans
            Exit Sub
        End If
        If Not rs_afmatrimonio.EOF Then
            If Not IsNull(rs_afmatrimonio!acu) Then afmaternidad = rs_afmatrimonio!acu
            Flog.writeline " -- obtenido el acumulador de Asignacion familiar por Maternidad"
            Flog.writeline " "
        End If
        rs_afmatrimonio.Close
        
        'Asignacion familiar Total
        aftotal = aftotalper + afretroactivo + afmaternidad
        
        '''''''''''''''''''''''''''''''''''''
        'Acumulador de imponibles, Prevision
        '''''''''''''''''''''''''''''''''''''
        
        'StrSql = "select sum(acu_mes.ammonto) acu "
        'StrSql = StrSql & " from his_estructura "
        'StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
        'StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu4
        'OpenRecordset StrSql, rs_imponible
        'If rs_imponible.EOF Then
        '    HuboError = True
        '    Flog.writeline "No se ha encontrado la configuracion de reporte , columna 4"
        '    MyRollbackTrans
        '    Exit Sub
        'End If
        
        StrSql = "select periodo.pliqnro,detliq.cliqnro,pliqmes,pliqanio,prodesc,proceso.empnro,acumulador.acudesabr,acumulador.acunro,concepto.concabr,concepto.concnro,detliq.dlicant,acu_liq.almonto "
        StrSql = StrSql & "From periodo "
        StrSql = StrSql & "inner join proceso on proceso.pliqnro=periodo.pliqnro "
        StrSql = StrSql & "inner join cabliq on proceso.pronro=cabliq.pronro "
        StrSql = StrSql & "inner join detliq on cabliq.cliqnro=detliq.cliqnro "
        StrSql = StrSql & "inner join acu_liq on cabliq.cliqnro=acu_liq.cliqnro "
        StrSql = StrSql & "inner join acumulador on acu_liq.acunro=acumulador.acunro "
        StrSql = StrSql & "inner join concepto on detliq.concnro=concepto.concnro "
        StrSql = StrSql & "inner join con_acum on con_acum.concnro=concepto.concnro and con_acum.acunro=acumulador.acunro "
        StrSql = StrSql & "Where concepto.concnro = " & acu4 & " And acumulador.acunro = " & acu7 & " And Proceso.Empnro = " & p_Empresa_ternro
        StrSql = StrSql & " order by dlicant "
     
        OpenRecordset StrSql, rs_imponible
        If rs_imponible.EOF Then
            HuboError = True
            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 4"
            MyRollbackTrans
            Exit Sub
        End If
              
        If Not rs_imponible.EOF Then
'            If Not IsNull(rs_imponible!acu) Then imponible4 = rs_imponible!acu
            Flog.writeline " -- obtenido el acumulador de imponibles, Prevision"
            Flog.writeline " ---- calculo de asignaciones compensables de Prevision"
        End If
        
        primera = True
        
        ' repnro = getLastIdentity(objConn, "rep_ps32")
        
        StrSql = "select repnro from rep_ps32 order by repnro desc"
        OpenRecordset StrSql, rs_repaux
        If Not (rs_repaux.EOF) Then
            repnro = rs_repaux!repnro + 1
        Else
            repnro = 1
        End If
                    
        Flog.writeline " ---- creacion del detalle del reporte tipo P (Prevision)"
        Dim valortot
        
        valortot = 0
        contadorError = 0
        
        While Not (rs_imponible.EOF)
           Valor = 0
           dlcant = rs_imponible!dlicant
           calculo = 0
           lcambio = False
              
           While Not (rs_imponible.EOF) And Not (lcambio)
             If (primera Or (dlcant = rs_imponible!dlicant)) Then
                 If contadorError < 6 Then
                      Valor = Valor + rs_imponible!almonto
                      primera = False
                 End If
                 rs_imponible.MoveNext
              Else
                 lcambio = True
              End If
              
           Wend
           contadorError = contadorError + 1
           
           If Not (rs_imponible.EOF) And (contadorError <= 6) Then
             calculo = (Valor * dlcant) / 100
             valortot = valortot + calculo
             strsql1 = "INSERT INTO REP_PS32_det "
             strsql1 = strsql1 & " (repnro,tipo,imponible,porcentaje,importe)"
             strsql1 = strsql1 & " VALUES("
             strsql1 = strsql1 & "'" & repnro & "',"
             strsql1 = strsql1 & "'" & "P" & "',"
             strsql1 = strsql1 & "'" & Valor & "',"
             strsql1 = strsql1 & "'" & dlcant & "',"
             strsql1 = strsql1 & "'" & calculo & "');"
             objConn.Execute strsql1, , adExecuteNoRecords
           End If
           If contadorError = 7 Then
              Flog.writeline " ---- No se puede calcular mas de 6 montos de tipo P (Prevision)"
           End If
        
        Wend
        
        
        rs_imponible.Close
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Acumulador de imponibles, Asignaciones Familiares
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'StrSql = "select sum(acu_mes.ammonto) acu "
        'StrSql = StrSql & " from his_estructura "
        'StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
        'StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu4
        'OpenRecordset StrSql, rs_imponible
        'If rs_imponible.EOF Then
        '    HuboError = True
        '    Flog.writeline "No se ha encontrado la configuracion de reporte , columna 4"
        '    MyRollbackTrans
        '    Exit Sub
        'End If
        
        StrSql = "select periodo.pliqnro,detliq.cliqnro,pliqmes,pliqanio,prodesc,proceso.empnro,acumulador.acudesabr,acumulador.acunro,concepto.concabr,concepto.concnro,detliq.dlicant,acu_liq.almonto "
        StrSql = StrSql & "From periodo "
        StrSql = StrSql & "inner join proceso on proceso.pliqnro=periodo.pliqnro "
        StrSql = StrSql & "inner join cabliq on proceso.pronro=cabliq.pronro "
        StrSql = StrSql & "inner join detliq on cabliq.cliqnro=detliq.cliqnro "
        StrSql = StrSql & "inner join acu_liq on cabliq.cliqnro=acu_liq.cliqnro "
        StrSql = StrSql & "inner join acumulador on acu_liq.acunro=acumulador.acunro "
        StrSql = StrSql & "inner join concepto on detliq.concnro=concepto.concnro "
        StrSql = StrSql & "inner join con_acum on con_acum.concnro=concepto.concnro and con_acum.acunro=acumulador.acunro "
        StrSql = StrSql & "Where concepto.concnro = " & acu5 & " And acumulador.acunro = " & acu7 & " And Proceso.Empnro = " & p_Empresa_ternro
        StrSql = StrSql & " order by dlicant "
     
        OpenRecordset StrSql, rs_imponible
        If rs_imponible.EOF Then
            HuboError = True
            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 5"
            MyRollbackTrans
            Exit Sub
        End If
              
        If Not rs_imponible.EOF Then
'            If Not IsNull(rs_imponible!acu) Then imponible4 = rs_imponible!acu
            Flog.writeline " -- obtenido el acumulador de imponibles, Asignaciones Familiares"
            Flog.writeline " ---- calculo de asignaciones compensables de  Asignaciones Familiares"
        End If
        
        primera = True
                    
        Flog.writeline " ---- creacion del detalle del reporte tipo F (Asignaciones Familiares)"
        contadorError = 0
        
        While Not (rs_imponible.EOF)
           Valor = 0
           dlcant = rs_imponible!dlicant
           calculo = 0
           lcambio = False
           
           While Not (rs_imponible.EOF) And Not (lcambio)
             If (primera Or (dlcant = rs_imponible!dlicant)) Then
                 If contadorError < 6 Then
                      Valor = Valor + rs_imponible!almonto
                      primera = False
                 End If
                 rs_imponible.MoveNext
              Else
                 lcambio = True
              End If
           Wend
           
           contadorError = contadorError + 1
           
           If Not (rs_imponible.EOF) And (contadorError <= 6) Then
             calculo = (Valor * dlcant) / 100
             valortot = valortot + calculo
             strsql1 = "INSERT INTO REP_PS32_det "
             strsql1 = strsql1 & " (repnro,tipo,imponible,porcentaje,importe)"
             strsql1 = strsql1 & " VALUES("
             strsql1 = strsql1 & "'" & repnro & "',"
             strsql1 = strsql1 & "'" & "F" & "',"
             strsql1 = strsql1 & "'" & Valor & "',"
             strsql1 = strsql1 & "'" & dlcant & "',"
             strsql1 = strsql1 & "'" & calculo & "');"
             objConn.Execute strsql1, , adExecuteNoRecords
          End If
          If contadorError = 7 Then
            Flog.writeline " ---- No se puede calcular mas de 6 montos de tipo F (Asignaciones Familiares)"
          End If
        Wend
          
        rs_imponible.Close
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        'Acumulador de imponibles, Reg. Nac. Emp '
        ''''''''''''''''''''''''''''''''''''''''''''''
                
        StrSql = "select periodo.pliqnro,detliq.cliqnro,pliqmes,pliqanio,prodesc,proceso.empnro,acumulador.acudesabr,acumulador.acunro,concepto.concabr,concepto.concnro,detliq.dlicant,acu_liq.almonto "
        StrSql = StrSql & "From periodo "
        StrSql = StrSql & "inner join proceso on proceso.pliqnro=periodo.pliqnro "
        StrSql = StrSql & "inner join cabliq on proceso.pronro=cabliq.pronro "
        StrSql = StrSql & "inner join detliq on cabliq.cliqnro=detliq.cliqnro "
        StrSql = StrSql & "inner join acu_liq on cabliq.cliqnro=acu_liq.cliqnro "
        StrSql = StrSql & "inner join acumulador on acu_liq.acunro=acumulador.acunro "
        StrSql = StrSql & "inner join concepto on detliq.concnro=concepto.concnro "
        StrSql = StrSql & "inner join con_acum on con_acum.concnro=concepto.concnro and con_acum.acunro=acumulador.acunro "
        StrSql = StrSql & "Where concepto.concnro = " & acu6 & " And acumulador.acunro = " & acu7 & " And Proceso.Empnro = " & p_Empresa_ternro
        StrSql = StrSql & " order by dlicant "
     
        OpenRecordset StrSql, rs_imponible
        If rs_imponible.EOF Then
            HuboError = True
            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 6"
            MyRollbackTrans
            Exit Sub
        End If
              
        If Not rs_imponible.EOF Then
'            If Not IsNull(rs_imponible!acu) Then imponible4 = rs_imponible!acu
            Flog.writeline " -- obtenido el acumulador de imponibles, Reg. Nac. Emp "
            Flog.writeline " ---- calculo de asignaciones compensables de Reg. Nac. Emp "
        End If
        
        primera = True
        
        ' repnro = getLastIdentity(objConn, "rep_ps32")
        
                    
        Flog.writeline " ---- creacion del detalle del reporte tipo E (Reg. Nac. Emp)"
        
        contadorError = 0
        
        While Not (rs_imponible.EOF)
           Valor = 0
           dlcant = rs_imponible!dlicant
           calculo = 0
           lcambio = False
           
           While Not (rs_imponible.EOF) And Not (lcambio)
             If (primera Or (dlcant = rs_imponible!dlicant)) Then
                 If contadorError < 6 Then
                      Valor = Valor + rs_imponible!almonto
                      primera = False
                 End If
                 rs_imponible.MoveNext
              Else
                 lcambio = True
              End If
           Wend
           
           contadorError = contadorError + 1
           
           If Not (rs_imponible.EOF) And (contadorError <= 6) Then
             calculo = (Valor * dlcant) / 100
             valortot = valortot + calculo
             strsql1 = "INSERT INTO REP_PS32_det "
             strsql1 = strsql1 & " (repnro,tipo,imponible,porcentaje,importe)"
             strsql1 = strsql1 & " VALUES("
             strsql1 = strsql1 & "'" & repnro & "',"
             strsql1 = strsql1 & "'" & "E" & "',"
             strsql1 = strsql1 & "'" & Valor & "',"
             strsql1 = strsql1 & "'" & dlcant & "',"
             strsql1 = strsql1 & "'" & calculo & "');"
             objConn.Execute strsql1, , adExecuteNoRecords
          End If
          If contadorError = 7 Then
              Flog.writeline " ---- No se puede calcular mas de 6 montos de tipo E (Reg. Nac. Empleo)"
          End If
        Wend
                  
          
        rs_imponible.Close
        
        'StrSql = "select sum(acu_mes.ammonto) acu "
        'StrSql = StrSql & " from his_estructura "
        'StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
        'StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu5
        'OpenRecordset StrSql, rs_imponible
        
'        If rs_imponible.EOF Then
'            HuboError = True
'            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 5"
'            MyRollbackTrans
'            Exit Sub
'        End If
        
'        If Not rs_imponible.EOF Then
'            If Not IsNull(rs_imponible!acu) Then imponible5 = rs_imponible!acu
'            Flog.writeline " -- obtenido el acumulador de imponibles, Asig. Familiares"
'        End If
'        rs_imponible.Close
        
        'Acumulador de imponibles, Reg. Nac. de Empleo
'        StrSql = "select sum(acu_mes.ammonto) acu "
'        StrSql = StrSql & " from his_estructura "
'        StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and empresa.empnro = " & p_Empresa_ternro
'        StrSql = StrSql & " join acu_mes on his_estructura.ternro = acu_mes.ternro and acu_mes.amanio = " & p_Anio & " and acu_mes.ammes = " & p_Mes & " and acu_mes.acunro = " & acu6
'        OpenRecordset StrSql, rs_imponible
'        If rs_imponible.EOF Then
'            HuboError = True
'            Flog.writeline "No se ha encontrado la configuracion de reporte , columna 6"
'            MyRollbackTrans
'            Exit Sub
'        End If
'        If Not rs_imponible.EOF Then
'            If Not IsNull(rs_imponible!acu) Then imponible6 = rs_imponible!acu
'            Flog.writeline " -- obtenido el acumulador de imponibles, Reg. Nac. de Empleo"
'        End If
'        rs_imponible.Close
'        Flog.writeline " "
        
        'Calculo de Asignaciones Compensables
        'imponible4_importe = (imponible4 * 5.72) / 100
        '    Flog.writeline " ---- calculo de asignaciones compensables de Prevision"
        'imponible5_importe = (imponible5 * 0.89) / 100
        '    Flog.writeline " ---- calculo de asignaciones compensables de Asig. Familiares"
        'imponible6_importe = (imponible6 * 0.5) / 100
        '    Flog.writeline " ---- calculo de asignaciones compensables del Reg. Nac. Del Empleo"
            
         '   Flog.writeline " "

        'Asignaciones compensables Total
        'actotal = imponible4_importe + imponible5_importe + imponible6_importe
        actotal = valortot
        'Saldo
        saldo = aftotal - actotal


        'Grabo los datos del reporte
        StrSql = "INSERT INTO REP_PS32 "
        StrSql = StrSql & " (empresa,bpronro,fecha,hora,iduser,mes,anio,cuit,razonsocial,actividad,"
        StrSql = StrSql & " domcalle,domnro,dompiso,domdepto,domcodpost,domlocalidad,domprovinc,"
        StrSql = StrSql & " telefono,ddn,cbu,banco,aftotalper,afretroactivo,afmaternidad,aftotal,"
        StrSql = StrSql & " actotal,saldo,sonpesos)"
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & "'" & p_Empresa_ternro & "',"
        StrSql = StrSql & "'" & NroProcesoBatch & "',"
        StrSql = StrSql & ConvFecha(Date) & ","
        StrSql = StrSql & "'" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "',"
        StrSql = StrSql & "'" & IdUser & "',"
        StrSql = StrSql & "'" & p_Mes & "',"
        StrSql = StrSql & "'" & p_Anio & "',"
        StrSql = StrSql & "'" & cuit & "',"
        StrSql = StrSql & "'" & razonsocial & "',"
        StrSql = StrSql & "'" & actividad & "',"
        StrSql = StrSql & "'" & domcalle & "',"
        StrSql = StrSql & "'" & domnro & "',"
        StrSql = StrSql & "'" & dompiso & "',"
        StrSql = StrSql & "'" & domdepto & "',"
        StrSql = StrSql & "'" & domcodpost & "',"
        StrSql = StrSql & "'" & domlocalidad & "',"
        StrSql = StrSql & "'" & domprovinc & "',"
        StrSql = StrSql & "'" & telefono & "',"
        StrSql = StrSql & "'" & ddn & "',"
        StrSql = StrSql & "'" & cbu & "',"
        StrSql = StrSql & "'" & banco & "',"
        StrSql = StrSql & "'" & aftotalper & "',"
        StrSql = StrSql & "'" & afretroactivo & "',"
        StrSql = StrSql & "'" & afmaternidad & "',"
        StrSql = StrSql & "'" & aftotal & "',"
        StrSql = StrSql & "'" & actotal & "',"
        StrSql = StrSql & "'" & saldo & "',"
        StrSql = StrSql & "'" & sonpesos & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Ultimo reporte creado
        
        Flog.writeline " -- Creacion del registro del reporte numero " & repnro
        Flog.writeline " "
        
        'Tipos: "prevision", "asig. familiares", "reg. nac. del empleo"
        'Grabo los datos del detalle, imponible 4
        'StrSql = "INSERT INTO REP_PS32_det "
        'StrSql = StrSql & " (repnro,tipo,imponible,porcentaje,importe)"
        'StrSql = StrSql & " VALUES("
        'StrSql = StrSql & "'" & repnro & "',"
        'StrSql = StrSql & "'" & "P" & "',"
        'StrSql = StrSql & "'" & imponible4 & "',"
        'StrSql = StrSql & "'" & "5.72" & "',"
        'StrSql = StrSql & "'" & imponible4_importe & "')"
        'objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline " ---- creacion del detalle del reporte tipo P (Prevision)"
        
        'Grabo los datos del detalle, imponible 5
        'StrSql = "INSERT INTO REP_PS32_det "
        'StrSql = StrSql & " (repnro,tipo,imponible,porcentaje,importe)"
        'StrSql = StrSql & " VALUES("
        'StrSql = StrSql & "'" & repnro & "',"
        'StrSql = StrSql & "'" & "F" & "',"
        'StrSql = StrSql & "'" & imponible5 & "',"
        'StrSql = StrSql & "'" & "0.89" & "',"
        'StrSql = StrSql & "'" & imponible5_importe & "')"
        'objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline " ---- creacion del detalle del reporte tipo F (Familiar)"
        
        'Grabo los datos del detalle, imponible 4
        'StrSql = "INSERT INTO REP_PS32_det "
        'StrSql = StrSql & " (repnro,tipo,imponible,porcentaje,importe)"
        'StrSql = StrSql & " VALUES("
        'StrSql = StrSql & "'" & repnro & "',"
        'StrSql = StrSql & "'" & "E" & "',"
        'StrSql = StrSql & "'" & imponible6 & "',"
        'StrSql = StrSql & "'" & "0.5" & "',"
        'StrSql = StrSql & "'" & imponible6_importe & "')"
        'objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline " ---- creacion del detalle del reporte tipo P (Reg. Nac. Empleo)"
        'Flog.writeline " "
        
'Fin de la transaccion
MyCommitTrans

Set rs_Empresa = Nothing

Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans
End Sub







Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim aux As String

Dim HFecha As Date
Dim Aux_Separador As String

Dim p_Mes
Dim p_Anio
Dim p_CtaBancaria
Dim p_Empresa_ternro

Aux_Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es Aux_Separador

If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        p_Mes = Mid(parametros, pos1, pos2 - pos1 + 1)
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        p_Anio = Mid(parametros, pos1, pos2 - pos1 + 1)
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        p_CtaBancaria = Mid(parametros, pos1, pos2 - pos1 + 1)
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        p_Empresa_ternro = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If

HFecha = "03/06/2005"
'Certificado Ansses de Servicios
Call Generar_Reporte(bpronro, HFecha, p_Empresa_ternro, p_Mes, p_Anio, p_CtaBancaria)
End Sub




