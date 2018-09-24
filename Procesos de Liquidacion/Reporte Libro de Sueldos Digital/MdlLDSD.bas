Attribute VB_Name = "MdlLDSD"
Option Explicit
'Const version = "1.0"
'Const FechaVersion = "04/08/2015" 'Gonzalez Nicolás
                                  'CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781
'Const version = "1.1"
'Const FechaVersion = "29/09/2015" 'Gonzalez Nicolás
                                  'CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781 - Mejoras y correcciones Registro 02, 04, 05

'Const version = "1.2"
'Const FechaVersion = "07/10/2015" 'Gonzalez Nicolás - CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781
                                  ' Mejoras en mensaje de logs y previsualización en Registro 04. columnas 35 a 43
                                  
'Const version = "1.3"
'Const FechaVersion = "16/10/2015" 'Gonzalez Nicolás - CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781
                                  'Se modificaron Columnas 13 y 14 y 15 del registro 04.
                                  
'Const version = "1.4"
'Const FechaVersion = "19/10/2015" 'Gonzalez Nicolás - CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781
                                  'Registro 05: Se busca código asociado a las estructuras (Mi simplifación)
'Const version = "1.5"
'Const FechaVersion = "22/10/2015" 'Gonzalez Nicolás - CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781
                                  'Registro 05: Se buscan procesos por empresa de cada empleado para el periodo del filtro.
                                  
'Const version = "1.6"
'Const FechaVersion = "23/10/2015" 'Gonzalez Nicolás - CAS-31922 - RH Pro (Producto) - ARG - LIQ - Libro de Sueldos Digital - AFIP RG 3781
                                  'Registro 05: Se corrige error al buscar el AC de remuneracion

Const version = "1.7"
Const FechaVersion = "28/01/2016" 'Gonzalez Nicolás - CAS-35372 - RHPro (Producto) - ARG - NOM - Bug reportes legales con sit de revista
                                  'GetSitdeRevista(): Corrección de cód. a informar para cuando hay mas de 3 Situaciones de Revista

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Global rs_Estructura As New ADODB.Recordset
Global rs_Estr_cod As New ADODB.Recordset
Global Arrcod()
Global Arrdia()
Global NroCol As Integer
Global NroError As String
Global MsgErr As String
Global MsgInfo As String
Global Ternro As Long
Global NroErrCab As Long
Global rs_fases As New ADODB.Recordset
Global ProgresoAux As Double

Const topeArreglo As Integer = 302
Dim rs_Confrep As New ADODB.Recordset
'------------------------------------------
'Parametros recuperados de batch_proceso
'------------------------------------------
Private Type Parametros
    Ternro As Long
    Pliqnro As Long
    pronro As String
    Empresa As Long
    infliq As Long
    Tipodeliq As String
End Type
Global ListParam As Parametros

'------------------------------------------
'Parametros recuperados de batch_proceso
'------------------------------------------
Private Type Periodo
    Pliqnro As Long
    pliqmes As Long
    pliqanio As Long
    pliqdesde As Date
    pliqhasta As Date
End Type
Global DatosPeriodo As Periodo
'------------------------------------------
'CONFREP

Const topeArrCrep As Integer = 302
Private Type CRep
    tipo As String
    tipo2 As String
    tipo3 As String
    tipo4 As String
    tipo5 As String
    Etiqueta As String
    confval As String
    confval2 As String
    confval3 As String
    confval4 As String
    confval5 As String
End Type
Global confRep(topeArrCrep) As CRep
Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte.
' Autor      : Gonzalez Nicolás
' Fecha      : 17/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam
Dim ArrParametros
Dim Empresa
Dim nListaProc
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

    On Error GoTo ME_Main
    Nombre_Arch = PathFLog & "LibroDeSueldosDigital" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 491 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        'bprcparam = rs_batch_proceso!bprcparam
        If Not EsNulo(rs_batch_proceso!bprcparam) Then
            Flog.writeline "Levantando parametros: " & rs_batch_proceso!bprcparam
            If InStr(rs_batch_proceso!bprcparam, "@") > 0 Then
                bprcparam = Split(rs_batch_proceso!bprcparam, "@")
                ListParam.Pliqnro = bprcparam(0)
                ListParam.pronro = bprcparam(1)
                ListParam.Empresa = bprcparam(2)
                ListParam.infliq = bprcparam(3)
                ListParam.Tipodeliq = bprcparam(4)
                
                '--------------------------------------------------------
                'LLAMO FUNCION QUE ARMA DATOS DE CONFREP
                '--------------------------------------------------------
                 Call ArmoDatosConfrep
                '--------------------------------------------------------
                '--------------------------------------------------------
               
                
                '01:: Datos referenciales del envío (Liquidación de SyJ y datos para DJ F931)
                Call Registro01
                
                '====================================================================================
                Progreso = 4
                IncPorc = 0
                TiempoAcumulado = GetTickCount
                Progreso = Progreso + IncPorc
                TiempoAcumulado = GetTickCount
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                '====================================================================================
                
                
                Progreso = 5
                ProgresoAux = 95
                If CLng(ListParam.infliq) = 0 Then 'SOLO se informa cuando es SJ
                    '02:: Datos referenciales de la Liquidación de SyJ del trabajador
                    ProgresoAux = 15
                    Call Registro02
                    'Flog.writteline "Progreso02: " & Progreso
                    
                    '03::Detalle de los conceptos de sueldo liquidados al trabajador
                    ProgresoAux = 10
                    Call Registro03
                    'Flog.writteline "Progreso03: " & Progreso
                    
                    ProgresoAux = 70
                Else
                    Flog.writeline "Registro 02: Solo se genera para => Informa Liquidación (SJ) "
                    Flog.writeline "Registro 03: Solo se genera para => Informa Liquidación (SJ) "
                End If
                
                '04::Datos del trabajador para el calculo de la DJ F931
                'REGISTRO OBLIGATORIO PARA TODOS LOS CASOS
                If ProgresoAux = 70 And confRep(300).confval <> "0" Then
                    ProgresoAux = 70
                ElseIf ProgresoAux = 95 And confRep(300).confval <> "0" Then
                    ProgresoAux = 55
                End If
                
                Call Registro04
                'Flog.writteline "Progreso04: " & Progreso
                'Call Registro04(NroProcesoBatch, rs_batch_proceso!bprcparam)
                
                If confRep(300).confval <> "0" Then
                     If ProgresoAux = 55 Then
                        ProgresoAux = 40
                     Else
                        ProgresoAux = 25
                     End If
                    Call Registro05
                    'Flog.writteline "Progreso05: " & Progreso
                End If
            Else
                Flog.writeline "Error en parametros."
                HuboError = True
            End If
        Else
            Flog.writeline "Error en parametros."
            HuboError = True
        End If

        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
    Else
        Flog.writeline "No encontró el proceso."
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    'Flog.writeline "HuboError:" & HuboError
    If Not HuboError Then
        'CAMBIO EL ESTADO DE LA CABECERA A 1 (GENERADO CON ERRORES)
        StrSql = "UPDATE rep_ar_libsuedigital_cab SET repsuedigerr=" & NroErrCab & " WHERE bpronro =" & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        'ACTUALIZO EL ESTADO A ERROR
        StrSql = "UPDATE rep_ar_libsuedigital_cab SET repsuedigerr=1 WHERE bpronro =" & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

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
        'CAMBIO EL ESTADO DE LA CABECERA A 1 (GENERADO CON ERRORES)
        StrSql = "UPDATE rep_ar_libsuedigital_cab SET repsuedigerr=1 WHERE bpronro =" & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub



Public Sub Registro01()
'==================================================================================================
'01:: Datos referenciales del envío (Liquidación de SyJ y datos para DJ F931)
'==================================================================================================
Dim EtiqNoinf As String
Dim EtiqBlanco As String
Dim Auxiliar As String
Dim TipoEvenro As Long
Dim Nivel As Long
Dim ternroResL As String
Dim Arraux
Dim listaErr As String
Dim Longitud
Longitud = 50
Dim cuit As String
Dim Empnom As String
Dim Pliqdesc As String
Dim MsgErr As String
Dim Registro As String
Dim IdenEnvio As String
Dim PerMesAnio As String
Dim Tipodeliq As String
Dim NumdeLiq As String
Dim DiasBase As String
Dim Tipoliq As String
EtiqNoinf = "No informa"
EtiqBlanco = "Informa en Blanco"

listaErr = ""
Registro = "01"

Flog.writeline ""
Flog.writeline "***************************************************************************************"
Flog.writeline "REGISTRO 01 :: Datos referenciales del envío (Liquidación de SyJ y datos para DJ F931)"
Flog.writeline "***************************************************************************************"
Flog.writeline ""

Ternro = ListParam.Ternro

'-------------------------------------------------------------------------------
'BUSCO DATOS DE LA EMPRESA
StrSql = "SELECT ternro,empnom FROM estructura"
StrSql = StrSql & " INNER JOIN empresa ON  empresa.estrnro = estructura.estrnro"
StrSql = StrSql & " WHERE estructura.Estrnro = " & ListParam.Empresa
OpenRecordset StrSql, rs
If Not rs.EOF Then
    ListParam.Ternro = rs!Ternro
    Empnom = rs!Empnom
End If


'-------------------------------------------------------------------------------
'BUSCO PERIODO
StrSql = "SELECT pliqnro, Pliqdesc,pliqmes,pliqanio,pliqdesde,pliqhasta FROM periodo WHERE pliqnro =" & ListParam.Pliqnro
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Pliqdesc = rs!Pliqdesc
    'pliqmes = rs!pliqmes
    'pliqanio = rs!pliqanio
    PerMesAnio = Format_Data(rs!pliqdesde, "AAAAMM")
    DatosPeriodo.Pliqnro = rs!Pliqnro
    DatosPeriodo.pliqmes = rs!pliqmes
    DatosPeriodo.pliqanio = rs!pliqanio
    DatosPeriodo.pliqdesde = rs!pliqdesde
    DatosPeriodo.pliqhasta = rs!pliqhasta
End If

Tipoliq = "RE"
If CLng(ListParam.infliq) = 0 Then
    Tipoliq = "SJ"
End If

'-------------------------------------------------------------------------------
'GUARDO ENCABEZADO DEL REPORTE
NroErrCab = 2
StrSql = " INSERT INTO rep_ar_libsuedigital_cab (bpronro,ternro,repsuedigdesc,repsuedigerr)"
StrSql = StrSql & " VALUES ("
StrSql = StrSql & NroProcesoBatch & "," & ListParam.Ternro & ",'" & Empnom & " - " & Pliqdesc & " (" & Tipoliq & ")'," & NroErrCab
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline UCase("Cabecera de Reporte Insertada")
Flog.writeline ""
'GUARDO ENCABEZADO DEL REPORTE --------------------------------------------------


'--------------------------------------------------------------------------------
' - Identificacion del tipo de registro
NroCol = 1
NroError = ""
MsgErr = ""
Texto = "Identificacion del tipo de registro"
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & Registro
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, Registro)
'**********************************************


'--------------------------------------------------------------------------------
' - CUIT del empleador
NroCol = 2
Flog.writeline ""
Texto = "CUIT del empleador"
MsgErr = ""
If IsNumeric(confRep(10).confval) And confRep(10).confval <> 0 Then
    cuit = getTerDoc(ListParam.Ternro, CLng(confRep(10).confval), 1, 1)
Else
    'NO SE ENCONTRO CONFIGURACIÓN PARA LA COLUMNA 1. Documentos de la empresa
    NroError = "1"
    MsgErr = "No se encontró configuración válida en el confrep columna 10 valor 1. Se setea Default => 6"
    cuit = getTerDoc(ListParam.Ternro, 6, 1, 1)
End If
'CONTROLO CUIT
If (cuit = "") Then
    NroError = "2"
    MsgErr = "No se ha encontrado el CUIT."
ElseIf Len(cuit) > 11 Or Len(cuit) < 11 Then
    NroError = "3"
    MsgErr = "El CUIT debe contener 11 dígitos (13 Incluyendo los guiones medios (-)" & "(" & cuit & ")"
End If
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, cuit)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, cuit)
'**********************************************

'--------------------------------------------------------------------------------
' - Identificación del envío
NroCol = 3
MsgErr = ""
Texto = "Identificación del envío"
'Valores permitidos. 'SJ'=Informa la liquidación de SyJ y datos de la DJ F931; 'RE'=Sólo informa datos de la DJ F931 a rectificar
NroError = ""
IdenEnvio = "SJ"
If CLng(ListParam.infliq = -1) Then
    IdenEnvio = "RE"
End If
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, IdenEnvio)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, IdenEnvio)
'**********************************************



'--------------------------------------------------------------------------------
' - Período | Formato: AAAAMM
NroCol = 4
NroError = ""
MsgErr = ""
Texto = "Período"
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, PerMesAnio)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, PerMesAnio)
'**********************************************


'--------------------------------------------------------------------------------
'- Si "identificación del envío='SJ'" los valores permitidos son: 'M'=mes; 'Q'=quincena; 'S'=semanal
'- Si "identificación del envío='RE'", en blanco
NroCol = 5
NroError = ""
MsgErr = ""
Texto = "Tipo de liquidación"
Tipodeliq = ""
If IdenEnvio = "SJ" Then
    If Len(ListParam.Tipodeliq) = 0 Then
        NroError = "4"
        MsgErr = "Error en parámetros. valores permitidos [M],[Q] o [S]"
    End If
    Tipodeliq = ListParam.Tipodeliq
Else
    MsgErr = EtiqBlanco
End If
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, Tipodeliq)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, Tipodeliq)
'**********************************************

'--------------------------------------------------------------------------------
'- Si "identificación del envío='SJ'" es el número de liquidación del SyJ del empleador
'- Si "identificación del envío='RE'", en blanco
NroCol = 6
NroError = ""
MsgErr = ""
Texto = "Número de liquidación"
NumdeLiq = ""
If IdenEnvio = "SJ" Then
    NumdeLiq = "1"
Else
    MsgErr = EtiqBlanco
End If
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, NumdeLiq)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, NumdeLiq)
'**********************************************

'--------------------------------------------------------------------------------
'Valores permitidos. '30'=Base 30; 'UD'=Base dias del mes'
NroCol = 7
NroError = ""
MsgErr = ""
Texto = "Dias base"
DiasBase = "30"
If CLng(confRep(11).confval) = -1 Then
    'Calcula el último día del mes calendario
    If DatosPeriodo.pliqmes = 12 Then
        Auxiliar = "01/01/" & DatosPeriodo.pliqanio + 1
        DiasBase = Day(DateAdd("d", -1, Auxiliar))
    Else
        Auxiliar = "01/" & DatosPeriodo.pliqmes + 1 & "/" & DatosPeriodo.pliqanio
        DiasBase = Day(DateAdd("d", -1, Auxiliar))
    End If
End If
Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, DiasBase)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, DiasBase)
'**********************************************

NroCol = 8
NroError = ""
MsgErr = ""
Texto = "Cantidad de trabajadores informados en registros 04 "
'Flog.writeline Texto
'Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, 0)
'**********************************************
'----- INSERTO REGISTRO
'**********************************************
Call InsReg(NroError, Registro, Texto, "0")
'**********************************************


'********************************************************
'--------------- INSERTO REGISTRO DE CORTE --------------
NroError = ""
NroCol = 0
Ternro = 0
Call InsReg(NroError, Registro, "-1", "")
'********************************************************


End Sub
Public Sub Registro02()
'==================================================================================================
'02:: Datos referenciales de la Liquidación de SyJ del trabajador
'==================================================================================================
Dim EtiqNoinf As String
Dim EtiqBlanco As String
Dim EtiqSindatos  As String
Dim EmpleAProc As Long
Dim Auxiliar As String
Dim Formadepago As String
Dim rsEmp As New ADODB.Recordset
Dim Longitud
Longitud = 50
Dim MsgErr As String
Dim cuil As String
Dim Registro As String
Dim AcredCBU As String
Dim DepRevTrab As String
Dim FechaPago As String
Dim FechaRubrica As String
Dim CantDiasTope As String
EtiqNoinf = "No informa"
EtiqBlanco = "Informa en Blanco"
EtiqSindatos = "No se encontraron datos."

Registro = "02"

Flog.writeline ""
Flog.writeline "***************************************************************************************"
Flog.writeline "REGISTRO 02 :: Datos referenciales de la Liquidación de SyJ del trabajador"
Flog.writeline "***************************************************************************************"
Flog.writeline ""

'Ternro = ListParam.Ternro

'CICLO EMPLEADOS
StrSql = "SELECT "
StrSql = StrSql & " empleado.ternro,empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2"
StrSql = StrSql & " ,empleado.nivnro,empleado.empdiscap,empleado.empemail,empleado.empremu"
StrSql = StrSql & " ,tercero.tersex,tercero.estcivnro,tercero.terfecnac,tercero.paisnro,tercero.nacionalnro,tercero.terfecing"
StrSql = StrSql & " ,cabliq.cliqnro "
StrSql = StrSql & " FROM Empleado"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro =empleado.ternro"
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro =empleado.ternro AND tenro = 10 "

StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.ternro =" & ListParam.Ternro

StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro AND cabliq.pronro IN (" & ListParam.pronro & ")"

StrSql = StrSql & " WHERE ( his_estructura.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & " AND (his_estructura.htethasta >=" & ConvFecha(DatosPeriodo.pliqdesde) & " OR his_estructura.htethasta IS NULL))"
StrSql = StrSql & " ORDER BY empleg"
OpenRecordset StrSql, rsEmp

EmpleAProc = rsEmp.RecordCount
If EmpleAProc = 0 Then
   EmpleAProc = 1
End If


IncPorc = (ProgresoAux / EmpleAProc)

'IncPorc = (Pr - Progreso) / EmpleAProc

If Not rsEmp.EOF Then
    'IncPorc = 1
    Do While Not rsEmp.EOF
        Ternro = rsEmp!Ternro

        Flog.writeline ""
        Flog.writeline String(Longitud, "-")
        Flog.writeline "-- EMPLEADO : " & rsEmp!Empleg
        Flog.writeline String(Longitud, "-")
       
        '--------------------------------------------------------------------------------
        ' - Identificacion del tipo de registro
        NroCol = 1
        NroError = ""
        MsgErr = ""
        Texto = "Identificacion del tipo de registro"
        'Flog.Writeline Texto
        'Flog.Writeline Registro
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & Registro
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, Registro)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        ' - CUIL del trabajador
        'Flog.Writeline ""
        NroCol = 2
        MsgErr = ""
        Texto = "CUIL del trabajador"
        'Flog.Writeline Texto
        If IsNumeric(confRep(20).confval) And confRep(20).confval <> 0 Then
            cuil = getTerDoc(Ternro, CLng(confRep(20).confval), 1, 1)
        Else
            'NO SE ENCONTRO CONFIGURACIÓN PARA LA COLUMNA 1. Documentos de la empresa
            NroError = "5"
            MsgErr = "No se encontró configuración válida en el confrep columna 20 valor 1. Se setea Default => 10"
            cuil = getTerDoc(Ternro, 10, 1, 1)
        End If
        'CONTROLO CUIL
        If (cuil = "") Then
            NroError = "6"
            MsgErr = "No se ha encontrado el CUIL."
        ElseIf Len(cuil) > 11 Then
            NroError = "7"
            MsgErr = "El CUIL debe contener 11 dígitos (13 Incluyendo los guiones medios (-)" & "(" & cuil & ")"
        End If
        
        If MsgErr <> "" And cuil <> "" Then
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & MsgErr
        Else
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(cuil = "", MsgErr, cuil)
        End If
        'Flog.writeline Texto
        'Flog.Writeline cuil
        
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, cuil)
        '**********************************************
        
        '--------------------------------------------------------------------------------
        'Legajo del trabajador
        NroCol = 3
        NroError = ""
        MsgErr = ""
        Texto = "Legajo del trabajador"
        'Flog.Writeline Texto
        'Flog.Writeline rsEmp!Empleg
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & rsEmp!Empleg
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, rsEmp!Empleg)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        'Dependencia de revista del trabajador
        'Es el área donde el mismo se desempeña. También es informativo. Por ejemplo: "Gerencia de Recursos Humanos".
        NroCol = 4
        NroError = ""
        MsgErr = ""
        Texto = "Dependencia de revista del trabajador"
        'Flog.Writeline Texto
        DepRevTrab = UCase(getEmpleadoEstr(Ternro, confRep(21).confval, "ESTRDABR", DatosPeriodo.pliqdesde, DatosPeriodo.pliqhasta, ""))
        If DepRevTrab = "" Then
            NroError = "8"
            'Flog.Writeline EtiqSindatos
            MsgErr = "No se encontró estructura configurada. Configuración del reporte en columna 21 valor 1."
        Else
            If Len(DepRevTrab) > 50 Then
                NroError = "9"
                DepRevTrab = Left(DepRevTrab, 50)
                MsgErr = "El nombre de la estructura no puede exceder 50 caracteres. El texto se formatea a 50." & "(" & DepRevTrab & ")"
                'Flog.Writeline DepRevTrab
            End If
            'Flog.Writeline DepRevTrab
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, DepRevTrab)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, DepRevTrab)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        'CBU de acreditación del pago
        'Sólo se informa si "forma de pago" es igual a '3' (acreditación en cuenta)
        NroCol = 5
        NroError = ""
        MsgErr = ""
        Texto = "CBU de acreditación del pago"
        'Flog.Writeline Texto
        'RECUPERO CONFIGURACIÓN
        Auxiliar = "SIGLA"
        If CLng(confRep(22).confval) = 0 Then
            Auxiliar = "DESEXT"
        End If
        'BUSCO CÓDIGO DE LA FORMA DE PAGO.
        Formadepago = getDatosPedidoPag(Ternro, DatosPeriodo.Pliqnro, rsEmp!cliqnro, Auxiliar, "")
        
        If Len(Formadepago) > 0 Then 'CONTROLO QUE LA FORMA DE PAGO SEA = 3
            If Trim(Formadepago) = "3" Then
                AcredCBU = getDatosPedidoPag(Ternro, DatosPeriodo.Pliqnro, rsEmp!cliqnro, "CTACBU", "")
                If AcredCBU = "" Then
                    NroError = "10"
                    'Flog.Writeline EtiqSindatos
                    MsgErr = "No se encontro CBU de acreditación el pago."
                Else
                    Flog.writeline AcredCBU
                    If Len(AcredCBU) > 22 Then
                        NroError = "11"
                        MsgErr = "El N° de CBU debe contenter hasta 22 dígitos." & "(" & AcredCBU & ")"
                    End If
                End If
            Else
                MsgErr = EtiqNoinf
            End If
        Else
            MsgErr = EtiqNoinf
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, AcredCBU)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, AcredCBU)
        '**********************************************
        '--------------------------------------------------------------------------------
        'Cantidad de días para proporcionar tope
        '3 enteros | Este campo esta indefinido
        NroCol = 6
        NroError = ""
        MsgErr = ""
        Texto = "Cantidad de días para proporcionar tope"
        CantDiasTope = "0"
        Auxiliar = ""
        If (confRep(26).confval <> "0") Then
            'Obtengo valor según tipo configurado
            CantDiasTope = getValoresLiq(confRep(26).tipo, Ternro, confRep(26).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, rsEmp!cliqnro, 0)
            CantDiasTope = IIf(CLng(CantDiasTope) = 0, "000", CantDiasTope)
            If CLng(CantDiasTope) > 999 Then
                NroError = "88"
                MsgErr = "El Valor máximo permitido es de 3 enteros."
            Else
                Auxiliar = Abs(CDbl(CantDiasTope) - CLng(CantDiasTope))
                If CDbl(Auxiliar) > 0 Then
                    NroError = "89"
                    MsgErr = "El Valor contiene parte decimal, se eliminarán."
                    'FORMATEO PARA ELIMINAR PARTE DECIMAL
                    CantDiasTope = CLng(Auxiliar)
                Else
                    'FORMATEO PARA ELIMINAR PARTE DECIMAL
                    CantDiasTope = CLng(CantDiasTope)
                End If
            End If
            
        Else
            NroError = "87"
            MsgErr = "No se encontró configuración para la columna 26 valor 1."
        End If
        
        CantDiasTope = IIf(Len(CantDiasTope) < 3, Format_StrLR(CantDiasTope, 3, "L", True, "0"), CantDiasTope)
        'Flog.Writeline Texto
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, CantDiasTope)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, CantDiasTope)
        '**********************************************
        
        '--------------------------------------------------------------------------------
        'Fecha de pago
        'Formato: AAAAMMDD
        NroCol = 7
        NroError = ""
        Texto = "Fecha de pago"
        MsgErr = ""
        Auxiliar = ""
        'Flog.Writeline Texto
        FechaPago = getDatosPedidoPag(Ternro, DatosPeriodo.Pliqnro, rsEmp!cliqnro, "FECPED", "")
        If FechaPago = "" Then
            NroError = "12"
            MsgErr = "No se encontró Fecha de Pedido de Pago."
           ' Flog.Writeline EtiqSindatos
        Else
            FechaPago = Format_Data(FechaPago, "AAAAMMDD")
            'Flog.Writeline FechaPago
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, FechaPago)
        
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, FechaPago)
        '**********************************************

        '--------------------------------------------------------------------------------
        'Fecha de rúbrica
        'Formato: AAAAMMDD
        NroCol = 8
        NroError = ""
        MsgErr = ""
        Texto = "Fecha de rúbrica"
        'Flog.Writeline Texto
        FechaRubrica = IIf(confRep(23).confval = "", "", confRep(23).confval)
        If FechaRubrica = "" Then
            NroError = "13"
            'Flog.Writeline EtiqSindatos
            MsgErr = "Se debe configurar una Fecha de Rúbrica.Configuración del reporte en columna 23 valor 1."
        Else
            FechaRubrica = Format_Data(FechaRubrica, "AAAAMMDD")
            'Flog.Writeline FechaRubrica
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, FechaRubrica)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, FechaRubrica)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        'Forma de pago
        'Valores permitidos. '1'=Efectivo; '2'=Cheque; '3'=Acreditación en cuenta
        NroCol = 9
        NroError = ""
        Texto = "Forma de pago"
        MsgErr = ""
        'Flog.Writeline Texto
        If IsNumeric(Formadepago) Then
            If CInt(Formadepago) = 1 Or CInt(Formadepago) = 2 Or CInt(Formadepago) = 3 Then
                'Flog.Writeline Formadepago
            Else
                NroError = "14"
                MsgErr = "Valores permitidos. [1] Efectivo; [2] Cheque; [3] Acreditación en cuenta." & "(" & Formadepago & ")"
            End If
        Else
            NroError = "14"
            MsgErr = "Valores permitidos. '1'=Efectivo; '2'=Cheque; '3'=Acreditación en cuenta"
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, Formadepago)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, Formadepago)
        '**********************************************
        
        
        '********************************************************
        '--------------- INSERTO REGISTRO DE CORTE --------------
        NroError = ""
        NroCol = 0
        Ternro = 0
        Call InsReg(NroError, Registro, "-1", "")
        '********************************************************
        
        
        '====================================================================================
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso =" & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        '====================================================================================
        rsEmp.MoveNext
    Loop

Else
    Flog.writeline "No se encontraron datos a procesar"
End If

End Sub
Public Sub Registro03()
'==================================================================================================
'03::  Detalle de los conceptos de sueldo liquidados al trabajador
'==================================================================================================
Dim EtiqNoinf As String
Dim EtiqBlanco As String
Dim EtiqSindatos  As String
Dim EmpleAProc As Long
Dim Auxiliar As String
Dim Formadepago As String
Dim rsEmp As New ADODB.Recordset
Dim Longitud
Longitud = 50
Dim MsgErr As String
Dim cuil As String
Dim Registro As String
Dim AcredCBU As String
Dim DepRevTrab As String
Dim FechaPago As String
Dim FechaRubrica As String
Dim UnidadCO As String
Dim dlimonto As String
Dim IndicMonto As String
EtiqNoinf = "No informa"
EtiqBlanco = "Informa en Blanco"
EtiqSindatos = "No se encontraron datos."

Registro = "03"

Flog.writeline ""
Flog.writeline "***************************************************************************************"
Flog.writeline "REGISTRO 03 ::  Detalle de los conceptos de sueldo liquidados al trabajador"
Flog.writeline "***************************************************************************************"
Flog.writeline ""

'Ternro = ListParam.Ternro

'CICLO EMPLEADOS
'--SIN AGRUPAR
'StrSql = "SELECT "
'StrSql = StrSql & " empleado.empleg ,empleado.Ternro , cabliq.Cliqnro, concepto.ConcCod, concepto.tconnro"
'StrSql = StrSql & " ,detliq.dlicant, detliq.dlimonto"

'--AGRUPADO POR conccod y empleado
StrSql = "SELECT  empleado.empleg ,empleado.Ternro , concepto.ConcCod, concepto.tconnro ,detliq.dlicant, SUM(detliq.dlimonto) dlimonto"
StrSql = StrSql & " FROM empleado"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro =empleado.ternro"
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND tenro = 10 "

StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.ternro =" & ListParam.Ternro
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro AND cabliq.pronro IN (" & ListParam.pronro & ")"
StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro"
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1"

StrSql = StrSql & " WHERE ( his_estructura.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & " AND (his_estructura.htethasta >=" & ConvFecha(DatosPeriodo.pliqdesde) & " OR his_estructura.htethasta IS NULL))"

StrSql = StrSql & " GROUP BY empleado.empleg ,empleado.Ternro , concepto.ConcCod, concepto.tconnro ,detliq.dlicant, detliq.dlimonto "
StrSql = StrSql & " ORDER BY empleg"
OpenRecordset StrSql, rsEmp
'Flog.writeline StrSql
EmpleAProc = rsEmp.RecordCount
If EmpleAProc = 0 Then
   EmpleAProc = 1
End If

IncPorc = (ProgresoAux / EmpleAProc)
If Not rsEmp.EOF Then
    'IncPorc = 1
    Do While Not rsEmp.EOF
        Ternro = rsEmp!Ternro
        Flog.writeline ""
        Flog.writeline String(Longitud, "-")
        Flog.writeline "-- EMPLEADO : " & rsEmp!Empleg
        Flog.writeline String(Longitud, "-")
        
       
        '--------------------------------------------------------------------------------
        ' - Identificacion del tipo de registro
        NroCol = 1
        NroError = ""
        MsgErr = ""
        Texto = "Identificacion del tipo de registro"
        'Flog.writeline Texto
        'Flog.writeline Registro
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & Registro
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, Registro)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        ' - CUIL del trabajador
        NroCol = 2
        MsgErr = ""
        Texto = "CUIL del trabajador"
        If IsNumeric(confRep(20).confval) And confRep(20).confval <> 0 Then
            cuil = getTerDoc(Ternro, CLng(confRep(20).confval), 1, 1)
        Else
            'NO SE ENCONTRO CONFIGURACIÓN PARA LA COLUMNA 1. Documentos de la empresa
            NroError = "5"
            MsgErr = "No se encontró configuración válida en el confrep columna 20 valor 1. Se setea Default => 10"
            cuil = getTerDoc(Ternro, 10, 1, 1)
        End If
        'CONTROLO CUIL
        If (cuil = "") Then
            NroError = "6"
            MsgErr = "No se ha encontrado el CUIL."
        ElseIf Len(cuil) > 11 Then
            NroError = "7"
            MsgErr = "El CUIL debe contener 11 dígitos (13 Incluyendo los guiones medios (-)." & "(" & cuil & ")"
        End If
        
        If MsgErr <> "" And cuil <> "" Then
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & MsgErr
        Else
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(cuil = "", MsgErr, cuil)
        End If
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, cuil)
        '**********************************************
        
        '--------------------------------------------------------------------------------
        'Código de concepto liquidado por el empleador
        NroCol = 3
        NroError = ""
        MsgErr = ""
        Texto = "Código de concepto liquidado por el empleador"
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, rsEmp!ConcCod)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, rsEmp!ConcCod)
        '**********************************************
                
        
        
        
        '--------------------------------------------------------------------------------
        'Cantidad
        NroCol = 4
        NroError = ""
        MsgErr = ""
        Texto = "Cantidad"
        'Flog.writeline Texto
        'Flog.writeline rsEmp!dlicant
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, rsEmp!dlicant)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, rsEmp!dlicant)
        '**********************************************

        
        '--------------------------------------------------------------------------------
        'Unidades
        'CONTROLO EL TIPO DE UNIDAD POR CONCEPTO
        NroCol = 5
        NroError = ""
        MsgErr = ""
        Texto = "Unidades"
        UnidadCO = getUnidadConc(rsEmp!tconnro)
        If UnidadCO = "" Then
            NroError = "15"
            MsgErr = "No se encontro unidad para el Concepto. Verifique configuración de columnas 24 y 25."
        End If
        'Flog.writeline UnidadCO
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, UnidadCO)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, UnidadCO)
        '**********************************************
        '--------------------------------------------------------------------------------
        'Importe
        NroCol = 6
        NroError = ""
        MsgErr = ""
        Texto = "Importe"
        'Flog.writeline Texto
        'Flog.writeline dlimonto
        'dlimonto = rsEmp!dlimonto
        dlimonto = Replace(FormatNumber(Abs(rsEmp!dlimonto), 2), ",", "")
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, dlimonto)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, dlimonto)
        '**********************************************
        
        '--------------------------------------------------------------------------------
        'Indicador Débito / Crédito
        NroCol = 7
        NroError = ""
        MsgErr = ""
        Texto = "Indicador Débito / Crédito"
        If CLng(rsEmp!dlimonto) > 0 Then
            IndicMonto = "C"
        Else
            IndicMonto = "D"
        End If
        'Flog.writeline IndicMonto
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, IndicMonto)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, IndicMonto)
        '**********************************************
        
        '--------------------------------------------------------------------------------
        'Período de ajuste retroactivo
        NroCol = 8
        NroError = ""
        MsgErr = ""
        Texto = "Período de ajuste retroactivo"
        'Flog.writeline Texto
        'Flog.writeline EtiqBlanco
        MsgErr = EtiqBlanco
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, "")
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, "")
        '**********************************************
        
       
        '********************************************************
        '--------------- INSERTO REGISTRO DE CORTE --------------
        NroError = ""
        NroCol = 0
        Ternro = 0
        Call InsReg(NroError, Registro, "-1", "")
        '********************************************************

        
        '====================================================================================
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        '====================================================================================
        
        
        rsEmp.MoveNext
    Loop

Else

    Flog.writeline "No se encontraron datos a procesar"
End If



End Sub
Public Sub Registro04()
'==================================================================================================
'04::   Datos del trabajador para el calculo de la DJ F931
'==================================================================================================
Dim EtiqNoinf As String
Dim EtiqBlanco As String
Dim EtiqSindatos  As String
Dim EmpleAProc As Long
Dim Auxiliar As String
Dim Formadepago As String
Dim rsEmp As New ADODB.Recordset
Dim cliqnro As String
Dim CambioEmp As Boolean
Dim NroConfCol As Long
Dim NroConfCol2 As Long
Dim ax As Long
'Dim rs_Estructura As New ADODB.Recordset
'Dim rs_Estr_cod As New ADODB.Recordset

Dim Fecha_Inicio_Fase As Date
Dim Fecha_Fin_Fase As Date
Dim cuil As String
Dim Registro As String
Dim MontoCant As Long

Dim CodSitRevista As String
'Situacion de Revista -
Dim CodSitRevista1 As String
Dim Aux_diainisr1 As String
Dim CodSitRevista2 As String
Dim Aux_diainisr2  As String
Dim CodSitRevista3  As String
Dim Aux_diainisr3  As String
Dim Longitud
Longitud = 75
Dim cont As Long
Dim CodTipoOper As String
Dim MarcaCobSCVO As String
Dim MarcaCorrRed As String
Dim Conyuge As String
Dim CantHijos As String
Dim CodTipEmpl As String
Dim CodActividad As String
Dim MarcaCCT As String
Dim CodSiniest As String
Dim CodCondicion As String
Dim CodContratacion As String
Dim CodLocalidad As String
Dim CantDiasTrab As Double
Dim CantDiasTrab2 As Double
Dim CantDiasTrabAux As String
Dim CantHsTrab As Double
Dim CantHsTrab2 As Double

Dim PorcApoAdicSS  As Double
Dim PorcApoAdicSS2  As Double
Dim PorcApoAdicSSAux  As String

Dim PorcConTarDif As Double
Dim PorcConTarDif2 As Double
Dim PorcConTarDifAux As String


Dim CantHsTrabAux As String
Dim CodObraS As String
Dim CantAdhOS As Long
Dim CantAdhOS2 As Long
Dim CantAdhOSAux As String
Dim CantAporAdOS As Double
Dim CantAporAdOS2 As Double
Dim CantAporAdOSAux As String
Dim ConAdicOS As Double
Dim ConAdicOS2 As Double
Dim ConAdicOSAux As String
Dim BaseDifAporOS As Double
Dim BaseDifAporOS2 As Double
Dim BaseDifAporOSAux As String
Dim BaseDifContOS As Double
Dim BaseDifContOS2 As Double
Dim BaseDifContOSAux As String
Dim BaseDifLeyRieOS As Double
Dim BaseDifLeyRieOS2 As Double
Dim BaseDifLeyRieOSAux As String
Dim RemMaterAnses As Double
Dim RemMaterAnses2 As Double
Dim RemMaterAnsesAux As String
Dim RemBruta As Double
Dim RemBruta2 As Double
Dim RemBrutaAux As String
Dim BaseImpo As Double
Dim BaseImpo2 As Double
Dim BaseImpoAux As String
Dim a As Long
EtiqNoinf = "No informa"
EtiqBlanco = "Informa en Blanco"
EtiqSindatos = "No se encontraron datos."

Registro = "04"

Flog.writeline ""
Flog.writeline "***************************************************************************************"
Flog.writeline "REGISTRO 04 ::   Datos del trabajador para el calculo de la DJ F931"
Flog.writeline "***************************************************************************************"
Flog.writeline ""

'Ternro = ListParam.Ternro

'ARMO LISTA DE CLIQNRO
StrSql = "SELECT DISTINCT "
StrSql = StrSql & " cabliq.cliqnro"
StrSql = StrSql & " FROM Empleado"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro =empleado.ternro"
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro =empleado.ternro AND tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.ternro =" & ListParam.Ternro
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro AND cabliq.pronro IN (" & ListParam.pronro & ")"
StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro"
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
StrSql = StrSql & " WHERE (his_estructura.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & " AND (his_estructura.htethasta >=" & ConvFecha(DatosPeriodo.pliqdesde) & " OR his_estructura.htethasta IS NULL))"
OpenRecordset StrSql, rsEmp
If Not rsEmp.EOF Then
    cliqnro = ""
    Do While Not rsEmp.EOF
        If cliqnro = "" Then
            cliqnro = rsEmp!cliqnro
        Else
            cliqnro = cliqnro & ", " & rsEmp!cliqnro
        End If
        rsEmp.MoveNext
    Loop
End If
'------------------------------------------------------------------------------------------------------------------------

'CICLO EMPLEADOS
StrSql = "SELECT DISTINCT "
StrSql = StrSql & " Empleado.empleg,Empleado.ternro "
StrSql = StrSql & " FROM Empleado"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro =empleado.ternro"
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro =empleado.ternro AND tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.ternro =" & ListParam.Ternro
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro AND cabliq.pronro IN (" & ListParam.pronro & ")"
StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro"
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
'StrSql = StrSql & " AND concepto.concimp = -1"
StrSql = StrSql & " WHERE (his_estructura.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & " AND (his_estructura.htethasta >=" & ConvFecha(DatosPeriodo.pliqdesde) & " OR his_estructura.htethasta IS NULL))"
StrSql = StrSql & " ORDER BY empleg"
OpenRecordset StrSql, rsEmp
'Flog.writeline StrSql
EmpleAProc = rsEmp.RecordCount
If EmpleAProc = 0 Then
   EmpleAProc = 1
End If

IncPorc = (ProgresoAux / EmpleAProc)
'IncPorc = 1
Ternro = 0
cont = 0

If Not rsEmp.EOF Then
    Do While Not rsEmp.EOF
        If Ternro <> CLng(rsEmp!Ternro) Then
            cont = cont + 1
            CambioEmp = True
        Else
            CambioEmp = False
        End If
        
        Ternro = rsEmp!Ternro
        'Cliqnro = rsEmp!Cliqnro
        Flog.writeline ""
        Flog.writeline String(Longitud, "-")
        Flog.writeline "-- EMPLEADO : " & rsEmp!Empleg
        Flog.writeline "-- CLIQNRO : " & cliqnro
        Flog.writeline String(Longitud, "-")
        
        '------------------------------------------------------------------------------------------------
        'CONTROLO LAS FASES DEL EMPLEADO
        Call ControlFases(rsEmp!Empleg, Fecha_Inicio_Fase, Fecha_Fin_Fase)
        
        
        '--------------------------------------------------------------------------------
        ' - Identificacion del tipo de registro
        NroCol = 1
        NroError = ""
        Texto = "Identificacion del tipo de registro"
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & Registro
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, Registro)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        ' - CUIL del trabajador
        NroCol = 2
        Texto = "CUIL del trabajador"
        MsgErr = ""
        'Flog.Writeline Texto
        If IsNumeric(confRep(20).confval) And confRep(20).confval <> 0 Then
            cuil = getTerDoc(Ternro, CLng(confRep(20).confval), 1, 1)
        Else
            'NO SE ENCONTRO CONFIGURACIÓN PARA LA COLUMNA 1. Documentos de la empresa
            NroError = "5"
            MsgErr = "No se encontró configuración válida en el confrep columna 20 valor 1. Se setea Default => 10"
            cuil = getTerDoc(Ternro, 10, 1, 1)
        End If
        'CONTROLO CUIL
        If (cuil = "") Then
            NroError = "6"
            MsgErr = "No se ha encontrado el CUIL."
        ElseIf Len(cuil) > 11 Then
            NroError = "7"
            MsgErr = "El CUIL debe contener 11 dígitos (13 Incluyendo los guiones medios (-)"
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(cuil = "", MsgErr, cuil)
        If MsgErr <> "" And cuil <> "" Then
            Flog.writeline String(Longitud, " ") & ": " & MsgErr
        End If
        
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, cuil)
        '**********************************************
        
        '--------------------------------------------------------------------------------
        'Marca de cónyuge
        'Valores permitidos: [0] (No) ; [1] (Si)
        NroCol = 3
        NroError = ""
        MsgErr = ""
        Texto = "Marca de cónyuge"
        Conyuge = getFamiliar(Ternro, confRep(40).confval, 0, "0")
        'Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & Conyuge
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", Conyuge, MsgErr)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, Conyuge)
        '**********************************************
        
        
        '--------------------------------------------------------------------------------
        'Cantidad
        NroCol = 4
        NroError = ""
        MsgErr = ""
        Texto = "Cantidad de hijos"
        CantHijos = getFamiliar(Ternro, confRep(40).confval2, 1, "0")
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CantHijos, MsgErr)
        '**********************************************
        '----- INSERTO REGISTRO
        '**********************************************
        Call InsReg(NroError, Registro, Texto, CantHijos)
        '**********************************************

        '--------------------------------------------------------------------------------
        'Marca de trabajador en CCT
        NroCol = 5
        NroError = ""
        MsgErr = ""
        Texto = "Marca de trabajador en CCT"
        MarcaCCT = ""
        
        If confRep(38).confval <> "0" Then
            'MarcaCCT = getEstructura(IIf(EsNulo(confRep(38).confval), "0", confRep(38).confval), Fecha_Fin_Fase, "CODEXT")
            MarcaCCT = getEstructura(confRep(38).confval2, confRep(38).confval, Fecha_Fin_Fase, "ESTRCODEXT", "")
        Else
            MarcaCCT = confRep(38).confval3
        End If
        
        If MarcaCCT = "" Then
            NroError = "75"
            MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 38 valor 1, 2 y 3."
        Else
            If Not IsNumeric(MarcaCCT) Then
                NroError = "75"
                MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 38 valor 1, 2 y 3."
            Else
                If CLng(MarcaCCT) < 0 Or CLng(MarcaCCT) > 1 Then
                    NroError = "75"
                    MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 38 valor 1, 2 y 3."
                End If
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", MarcaCCT, MsgErr)
        '********************************************************************************************
        '----- INSERTO REGISTRO
        '********************************************************************************************
        Call InsReg(NroError, Registro, Texto, MarcaCCT)
        '********************************************************************************************
        
        
        '--------------------------------------------------------------------------------
        'Marca de cobertura de SCVO
        NroCol = 6
        NroError = ""
        MsgErr = ""
        MarcaCobSCVO = ""
        Texto = "Marca de cobertura de SCVO"
        If confRep(39).confval <> "0" Then
            MarcaCobSCVO = getEstructura(confRep(39).confval2, confRep(39).confval, Fecha_Fin_Fase, "ESTRCODEXT", "")
        Else
            MarcaCobSCVO = confRep(39).confval3
        End If
        
        If MarcaCobSCVO = "" Then
            NroError = "76"
            MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 39 valor 1, 2 y 3."
        Else
            If Not IsNumeric(MarcaCobSCVO) Then
                NroError = "76"
                MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 39 valor 1, 2 y 3."
            Else
                If CLng(MarcaCobSCVO) < 0 Or CLng(MarcaCobSCVO) > 1 Then
                    NroError = "76"
                    MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 39 valor 1, 2 y 3."
                End If
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", MarcaCobSCVO, MsgErr)
        '********************************************************************************************
        '----- INSERTO REGISTRO
        '********************************************************************************************
        Call InsReg(NroError, Registro, Texto, MarcaCobSCVO)
        '********************************************************************************************
        
        '--------------------------------------------------------------------------------
        'Marca de corresponde reducción
        'VALOR FIJO [0]
        NroCol = 7
        NroError = ""
        MsgErr = ""
        Texto = "Marca de corresponde reducción"
        MarcaCorrRed = ""
        If confRep(36).confval <> "0" Then
            MarcaCorrRed = getEstructura(confRep(36).confval2, confRep(36).confval, Fecha_Fin_Fase, "ESTRCODEXT", "")
        Else
            MarcaCorrRed = confRep(36).confval3
        End If

        If MarcaCorrRed = "" Then
            NroError = "79"
            MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 36 valor 1, 2 y 3."
        Else
            If Not IsNumeric(MarcaCorrRed) Then
                NroError = "79"
                MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 36 valor 1, 2 y 3."
            Else
                If CLng(MarcaCorrRed) < 0 Or CLng(MarcaCorrRed) > 1 Then
                    NroError = "79"
                    MsgErr = "Valores permitidos 0 ó 1. Ver configuración columna 36 valor 1, 2 y 3."
                End If
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", MarcaCorrRed, MsgErr)
        '********************************************************************************************
        '----- INSERTO REGISTRO
        '********************************************************************************************
        Call InsReg(NroError, Registro, Texto, MarcaCorrRed)
        '********************************************************************************************
        
        '--------------------------------------------------------------------------------
        'Código de tipo de empleador asociado al trabajador
        NroCol = 8
        NroError = ""
        MsgErr = ""
        Texto = "Código de tipo de empleador asociado al trabajador"
        'Flog.Writeline Texto
        CodTipEmpl = getEmpresaTipoEmp()
        If CodTipEmpl = "" Then
            NroError = "16"
            MsgErr = "No se encontró el código DGI del Tipo de Empleador Asociado a la Estructura Empresa."
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodTipEmpl, MsgErr)
        
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodTipEmpl)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Código de tipo de operación
        'VALOR FIJO [0]
        NroCol = 9
        NroError = ""
        MsgErr = ""
        Texto = "Código de tipo de operación"
        CodTipoOper = ""
        If confRep(37).confval <> "0" Then
            CodTipoOper = getEstructura(confRep(37).confval2, confRep(37).confval, Fecha_Fin_Fase, "ESTRCODEXT", "")
        Else
            CodTipoOper = confRep(37).confval3
        End If
        
        If CodTipoOper = "" Then
            NroError = "77"
            MsgErr = "No se encontró valor. Se setea default [0]. Ver configuración columna 37 valor 1, 2 y 3."
        Else
            If Len(CodTipoOper) > 1 Then
                NroError = "78"
                MsgErr = "El valor no debe exceder 1 caracter."
            End If
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodTipoOper, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodTipoOper)
        '********************************************************
        
        '--------------------------------------------------------------------------------
        'Código de situación de revista
        NroCol = 10
        NroError = ""
        MsgErr = ""
        MsgInfo = ""
        Texto = "Código de situación de revista"
        'Obtengo datos de la situación de revista (Para Columnas 10, 16, 17, 18, 19, 20 y 21)
        Call GetSitdeRevista(CodSitRevista, CodSitRevista1, Aux_diainisr1, CodSitRevista2, Aux_diainisr2, CodSitRevista3, Aux_diainisr3, NroError, MsgErr, MsgInfo)
        If Len(CodSitRevista) > 2 Then
            NroError = "17"
            MsgErr = "El código de situación de revista solo permite hasta 2 dígitos."
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodSitRevista, MsgErr)
        'MENSAJES INFORMATIVOS
        If MsgInfo <> "" Then
            Flog.writeline String(Longitud, " ") & ": " & MsgInfo
        End If
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodSitRevista)
        '********************************************************
        
        '--------------------------------------------------------------------------------
        'Código de condición
        NroCol = 11
        NroError = ""
        MsgErr = ""
        Texto = "Código de condición"
        CodCondicion = getCondicionSIJP(IIf(EsNulo(confRep(42).confval3), "0", confRep(42).confval3), Fecha_Fin_Fase)
        If CStr(CodCondicion) = "-1" Then
            CodCondicion = "0"
            NroError = "18"
            MsgErr = "No se encontró el codigo interno para la Condicion de SIJP. Se informa Default en 0 (cero)."
        ElseIf CStr(CodCondicion) = "0" Then
            NroError = "19"
            MsgErr = "No se encontro la Condicion de SIJP. Se informa Default en 0 (cero)."
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodCondicion, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodCondicion)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Código de actividad
        NroCol = 12
        NroError = ""
        MsgErr = ""
        Texto = "Código de actividad"
        CodActividad = getActividad(IIf(EsNulo(confRep(42).confval3), "0", confRep(42).confval3), Fecha_Fin_Fase)
        If CodActividad = "0" Then
            NroError = "25"
            MsgErr = "No se encontró la Actividad.Se informa Default en 0 (cero)."
            CodActividad = "0"
        ElseIf CodActividad = "-1" Then
            NroError = "26"
            MsgErr = "No se encontró el código interno para la Actividad del empleado. Se informa Default en 0 (cero)."
            CodActividad = "0"
        End If
        'Flog.Writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & CodActividad
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodActividad, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodActividad)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Código de modalidad de contratación
        NroCol = 13
        NroError = ""
        MsgErr = ""
        Texto = "Código de modalidad de contratación"
        'Flog.Writeline Texto
        CodContratacion = getContratoActual(IIf(EsNulo(confRep(42).confval3), "0", confRep(42).confval3), Fecha_Fin_Fase)
        If CodContratacion = "0" Then
            NroError = "23"
            MsgErr = "No se encontró el Tipo de Contrato.Se informa Default en 0 (cero)."
        ElseIf CodContratacion = "-1" Then
            NroError = "24"
            MsgErr = "No se encontró el codigo interno para el Tipo de Contrato. Se informa Default en 0 (cero)."
            CodContratacion = "0"
        End If
        'Flog.Writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & CodContratacion
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodContratacion, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodContratacion)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Código de siniestrado
        NroCol = 14
        NroError = ""
        MsgErr = ""
        Texto = "Código de siniestrado"
        CodSiniest = ""
        CodSiniest = getCodSiniestrado(IIf(EsNulo(confRep(42).confval3), "0", confRep(42).confval3), confRep(43).confval, Fecha_Fin_Fase)
        If CodContratacion = "00" Then
            NroError = "28"
            MsgErr = "No se encontró la estructura para el Código de Siniestrado. Se informa valor default [00] (Doble Cero)"
        ElseIf CodSiniest = "-1" Then
            NroError = "27"
            MsgErr = "No se encontraron Códigos de Siniestrado. Se informa valor default [00] (Doble Cero)"
            CodSiniest = "00"
        End If

        
        '---
'        If confRep(43).confval <> "" And confRep(43).confval <> "0" Then
'            StrSql = " SELECT emp_licnro FROM emp_lic"
'            StrSql = StrSql & " WHERE emp_lic.tdnro in (" & confRep(43).confval & ")"
'            StrSql = StrSql & " AND ("
'            StrSql = StrSql & " ((emp_lic.elfechadesde <= " & ConvFecha(DatosPeriodo.pliqdesde) & ") AND"
'            StrSql = StrSql & " (emp_lic.elfechahasta >= " & ConvFecha(DatosPeriodo.pliqdesde) & "))"
'            StrSql = StrSql & " OR"
'            StrSql = StrSql & " ((emp_lic.elfechadesde >= " & ConvFecha(DatosPeriodo.pliqdesde) & ") AND"
'            StrSql = StrSql & " (emp_lic.elfechadesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & "))"
'            StrSql = StrSql & " )"
'            StrSql = StrSql & " AND emp_lic.licestnro = 2"
'            StrSql = StrSql & " AND emp_lic.empleado = " & Ternro
'            OpenRecordset StrSql, rs_Estructura
'            If Not rs_Estructura.EOF Then
'                CodSiniest = "01"
'            Else
'                NroError = "27"
'                MsgErr = "No se encontraron Códigos de Siniestrado. Se informa valor default 00 (Doble Cero)"
'            End If
'        Else
'            NroError = "28"
'            MsgErr = "No se encontraron Licencias configuadas para el Código de Siniestrado. Se informa valor default 00 (Doble Cero)."
'            'MsgInfo = "No se encontraron licencias configuradas para el Cód. de Siniestrado. Configuración en la columna 43 valor 1 del confrep."
'           'No se encontró configuración válida en el confrep columna 20 valor 1.
'        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodSiniest, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodSiniest)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Código de localidad
        NroCol = 15
        NroError = ""
        MsgErr = ""
        Texto = "Código de localidad"
        CodLocalidad = "00"
'        Flog.writeline confRep(41).confval
        If CInt(confRep(41).confval) > 0 And CInt(confRep(41).confval) < 5 Then
            CodLocalidad = getZona(confRep(41).confval, DatosPeriodo.pliqhasta)
            If CodLocalidad = "00" Then
                NroError = "30"
                MsgErr = "No se encontró el Código de Localidad. Se informa valor default [00] (Doble Cero)."
            End If
        Else
            NroError = "29"
            MsgErr = "Valores Válidos de configuración : 1, 2 , 3 o 4. Configuración en columna 41 Valor 1."
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodLocalidad, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodLocalidad)
        '********************************************************
        
        '--------------------------------------------------------------------------------
        'Situación de revista 1
        NroCol = 16
        NroError = ""
        MsgErr = ""
        Texto = "Situación de revista 1"
        If CodSitRevista1 = "" Then
            MsgErr = EtiqSindatos
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodSitRevista1, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodSitRevista1)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Día de inicio situación de revista 1
        NroCol = 17
        NroError = ""
        MsgErr = ""
        Texto = "Día de inicio situación de revista 1"
        If Aux_diainisr1 = "" Then
            MsgErr = EtiqSindatos
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", Aux_diainisr1, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, Aux_diainisr1)
        '********************************************************
        
        '--------------------------------------------------------------------------------
        'Situación de revista
        NroCol = 18
        NroError = ""
        MsgErr = ""
        Texto = "Situación de revista 2"
        If CodSitRevista2 = "" Then
            MsgErr = EtiqSindatos
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodSitRevista2, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodSitRevista2)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Día de inicio situación de revista 1
        NroCol = 19
        NroError = ""
        MsgErr = ""
        Texto = "Día de inicio situación de revista 2"
        If Aux_diainisr2 = "" Then
            MsgErr = EtiqSindatos
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", Aux_diainisr2, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, Aux_diainisr2)
        '********************************************************

        '--------------------------------------------------------------------------------
        'Situación de revista 3
        NroCol = 20
        NroError = ""
        MsgErr = ""
        Texto = "Situación de revista 3"
        If CodSitRevista3 = "" Then
            MsgErr = EtiqSindatos
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodSitRevista3, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CodSitRevista3)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Día de inicio situación de revista 3
        NroCol = 21
        NroError = ""
        MsgErr = ""
        Texto = "Día de inicio situación de revista 3"
        If Aux_diainisr3 = "" Then
            MsgErr = EtiqSindatos
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", Aux_diainisr3, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, Aux_diainisr3)
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Cantidad de días trabajados
        NroCol = 22
        NroError = ""
        MsgErr = ""
        NroConfCol = 44
        NroConfCol2 = 244
        Texto = "Cantidad de días trabajados"
        CantDiasTrab = 0
        If (confRep(NroConfCol).confval <> "0") Then
            'Obtengo valor según tipo configurado
            CantDiasTrab = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "34"
            MsgErr = "No se encontró configuración para la columna 44 valor 1."
        End If

        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        CantDiasTrab2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                CantDiasTrab2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                CantDiasTrab = CantDiasTrab + CantDiasTrab2
            End If
        Next
        
       
        'Controlo el Valor
        If CLng(CantDiasTrab) < 0 Then
            NroError = "31"
            MsgErr = "Los días trabajados son negativos. Verifique configuración en columnas 44 y 244 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no contenga decimales
            If (CDbl(CantDiasTrab) - CLng(CantDiasTrab)) > 0 Then
                NroError = "32"
                MsgErr = "El valor contiene decimales. Verifique configuración en columnas 44 y 244 valor 1, 2, 3, 4 y 5."
            Else
                'Controlo que el valor no exceda 2 dígitos
                CantDiasTrabAux = CLng(CantDiasTrab)
                If Len(CantDiasTrabAux) > 2 Then
                    NroError = "33"
                    MsgErr = "El valor contiene más de 2 enteros. Verifique configuración en columnas 44 y 244 valor 1, 2, 3, 4 y 5."
                End If
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CantDiasTrab, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, CLng(CantDiasTrab))
        '****************************************************************************************************************
        
        
        '--------------------------------------------------------------------------------
        'Cantidad de horas trabajadas
        NroCol = 23
        NroError = ""
        MsgErr = ""
        Texto = "Cantidad de horas trabajadas"
        CantHsTrab = 0
        NroConfCol = 45
        NroConfCol2 = 245
        If (confRep(NroConfCol).confval <> "0") Then
            'Obtengo valor según tipo configurado
            CantHsTrab = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "38"
            MsgErr = "No se encontró configuración para la columna 45 valor 1."
        End If
        
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        CantHsTrab2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                CantHsTrab2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                CantHsTrab = CantHsTrab + CantHsTrab2
            End If
        Next
       
        If CLng(CantHsTrab) < 0 Then
            NroError = "35"
            MsgErr = "Las horas trabajadas son negativas. Verifique configuración en columnas 45 y 245 valor 1,2,3,4 y 5."
        Else
            'Controlo que no contenga decimales
            If (CDbl(CantHsTrab) - CLng(CantHsTrab)) > 0 Then
                NroError = "36"
                MsgErr = "El valor contiene decimales. Verifique configuración en columnas 45 y 245 valor 1,2,3,4 y 5."
            Else
                'Controlo que el valor no exceda 2 dígitos
                CantHsTrabAux = CLng(CantHsTrab)
                'Flog.writeline "CantDiasTrabAux-" & Len(CantDiasTrabAux)
                If Len(CantHsTrabAux) > 3 Then
                    NroError = "37"
                    MsgErr = "El valor contiene más de 3 enteros. Verifique configuración en columnas 45 y 245 valor 1,2,3,4 y 5."
                End If
            End If
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CantHsTrab, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, CLng(CantHsTrab))
        '********************************************************
                
       
        '--------------------------------------------------------------------------------
        'Porcentaje de aporte adicional de seguridad social
        NroCol = 24
        NroError = ""
        MsgErr = ""
        Texto = "Porcentaje de aporte adicional de seguridad social"
        NroConfCol = 63
        NroConfCol2 = 263
        PorcApoAdicSS = 0
        If (confRep(NroConfCol).confval <> "0") Then
            'Obtengo valor según tipo configurado
            PorcApoAdicSS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "69"
            MsgErr = "No se encontró configuración para la columna 63 valor 1."
        End If
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        PorcApoAdicSS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                PorcApoAdicSS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                PorcApoAdicSS = PorcApoAdicSS + PorcApoAdicSS2
            End If
        Next
        
        If CLng(PorcApoAdicSS) < 0 Then
            NroError = "70"
            MsgErr = "El Porcentaje de aporte adicional de seguridad social es negativo. Verifique configuración en columnas 63 y 263  valor 1,2,3,4 y 5."
        Else
            'Controlo que el valor no exceda 2 dígitos
            PorcApoAdicSSAux = CLng(PorcApoAdicSS)
            If Len(PorcApoAdicSSAux) > 5 Then
                NroError = "71"
                MsgErr = "El valor contiene más de 5 enteros. Verifique configuración en columnas 63 y 263 valor 1,2,3,4 y 5."
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", PorcApoAdicSS, MsgErr)
        '********************************************************
        '----- INSERTO REGISTRO
        '********************************************************
        Call InsReg(NroError, Registro, Texto, Format(PorcApoAdicSS, "#####0.00"))
        '********************************************************
        
        
        '--------------------------------------------------------------------------------
        'Porcentaje de contribución por tarea diferencial
        NroCol = 25
        NroError = ""
        MsgErr = ""
        NroConfCol = 63
        NroConfCol2 = 263
        Texto = "Porcentaje de contribución por tarea diferencial"
        PorcConTarDif = 0
        If (confRep(NroConfCol).confval <> "0") Then
            'Obtengo valor según tipo configurado
            PorcConTarDif = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "72"
            MsgErr = "No se encontró configuración para la columna 64 valor 1."
        End If
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        PorcConTarDif2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                PorcConTarDif2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                PorcConTarDif = PorcConTarDif + PorcConTarDif2
            End If
        Next
        
        If CLng(PorcConTarDif2) < 0 Then
            NroError = "73"
            MsgErr = "El Porcentaje de contribución por tarea diferencial es negativo. Verifique configuración en columnas 64 y 264  valor 1,2,3,4 y 5."
        Else
            'Controlo que el valor no exceda 2 dígitos
            PorcConTarDifAux = CLng(PorcConTarDif)
            If Len(PorcConTarDifAux) > 5 Then
                NroError = "74"
                MsgErr = "El valor contiene más de 5 enteros. Verifique configuración en columnas 64 y 264 valor 1,2,3,4 y 5."
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", PorcConTarDif, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(PorcConTarDif, "#####0.00"))
        '****************************************************************************************************************
        
        
        '--------------------------------------------------------------------------------
        'Código de obra social del trabajador
        NroCol = 26
        NroError = ""
        MsgErr = ""
        Texto = "Código de obra social del trabajador"
        CodObraS = getCodOS(IIf(EsNulo(confRep(42).confval3), "0", confRep(42).confval3), Fecha_Fin_Fase)
        If CodObraS = "-1" Then
            NroError = "39"
            MsgErr = "No se encontró el codigo interno para la Obra Social. Se informa Default en [000000]. Configuración columna 42 valor 3."
            CodObraS = "0"
        ElseIf CodObraS = "0" Then
            NroError = "40"
            MsgErr = "No se encontró la Obra Social. Se informa Default en [000000]. Configuración columna 42 valor 3."
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CodObraS, MsgErr)
        
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(CodObraS, "#######000000"))
        '****************************************************************************************************************
        
        
        '--------------------------------------------------------------------------------
        'Cantidad de adherentes de obra social
        NroCol = 27
        NroError = ""
        MsgErr = ""
        NroConfCol = 46
        NroConfCol2 = 246
        Texto = "Cantidad de adherentes de obra social"
        CantAdhOS = 0
        If confRep(NroConfCol).confval <> "0" Then
            CantAdhOS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "44"
            MsgErr = "No se encontró configuración para la columna 46 valor 1."
        End If
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        CantAdhOS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                CantAdhOS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                CantAdhOS = CantAdhOS + CantAdhOS2
            End If
        Next
        If CLng(CantAdhOS) < 0 Then
            NroError = "41"
            MsgErr = "El valor de Cantidad de adherentes es negativo. Verifique configuración en columnas 46 y 246 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no contenga decimales
            If (CDbl(CantAdhOS) - CLng(CantAdhOS)) > 0 Then
                NroError = "42"
                MsgErr = "El valor contiene decimales. Verifique configuración en columnas 46 y 246 valor 1, 2, 3, 4 y 5."
            Else
                'Controlo que el valor no exceda 2 dígitos
                CantAdhOSAux = CLng(CantAdhOS)
                If Len(CantAdhOSAux) > 2 Then
                    NroError = "43"
                    MsgErr = "El valor contiene más de 2 enteros. Verifique columnas 46 y 246 valor 1, 2, 3, 4 y 5."
                End If
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CantAdhOS, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(CantAdhOS, "##00"))
        '****************************************************************************************************************


        '--------------------------------------------------------------------------------
        'Aporte adicional de obra social
        NroCol = 28
        NroError = ""
        MsgErr = ""
        CantAporAdOS = 0
        NroConfCol = 47
        NroConfCol2 = 247
        Texto = "Aporte adicional de obra social"
        If confRep(NroConfCol).confval <> "0" Then
            CantAporAdOS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "47"
            MsgErr = "No se encontró configuración para la columna 47 valor 1."
        End If
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        CantAporAdOS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                CantAporAdOS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                CantAporAdOS = CantAporAdOS + CantAporAdOS2
            End If
        Next
        
        If CantAporAdOS < 0 Then
            NroError = "45"
            MsgErr = "El valor de Aporte adicional es negativo. Verifique configuración en columnas 47 y 247 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no exceda de 13 Enteros
            CantAporAdOSAux = Fix(CantAporAdOS)
            If Len(CantAporAdOSAux) > 13 Then
                NroError = "46"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columnas 47 y 247 valor 1, 2, 3, 4 y 5."
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", CantAporAdOS, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(CantAporAdOS, "#############0.00"))
        '****************************************************************************************************************



        '--------------------------------------------------------------------------------
        'Contribución adicional de obra social
        NroCol = 29
        NroError = ""
        MsgErr = ""
        NroConfCol = 48
        NroConfCol2 = 248
        Texto = "Contribución adicional de obra social"
        ConAdicOS = 0
        If confRep(NroConfCol).confval <> "0" Then
            ConAdicOS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "50"
            MsgErr = "No se encontró configuración para la columna 48 valor 1."
        End If
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        ConAdicOS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                ConAdicOS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                ConAdicOS = ConAdicOS + ConAdicOS2
            End If
        Next
        
        If CLng(ConAdicOS) < 0 Then
            NroError = "48"
            MsgErr = "El valor de Contribución adicional es negativo. Verifique configuración en columna 48 valor 1."
        Else
            'Controlo que no exceda de 13 Enteros
            ConAdicOSAux = Fix(ConAdicOS)
            If Len(ConAdicOSAux) > 13 Then
                NroError = "49"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columna 48 valor 1."
            End If
        End If
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", ConAdicOS, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(ConAdicOS, "#############0.00"))
       '****************************************************************************************************************
        
        '--------------------------------------------------------------------------------
        'Base para el cálculo diferencial de aporte de obra social y FSR (1)
        NroCol = 30
        NroError = ""
        MsgErr = ""
        NroConfCol = 49
        NroConfCol2 = 249
        Texto = "Base para el cálculo diferencial de aporte de obra social y FSR (1)"
        BaseDifAporOS = 0
        If confRep(NroConfCol).confval <> "0" Then
            BaseDifAporOS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "53"
            MsgErr = "No se encontró configuración para la columna 49 valor 1."
        End If
        
        'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        BaseDifAporOS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                BaseDifAporOS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                BaseDifAporOS = BaseDifAporOS + BaseDifAporOS2
            End If
        Next
        
        If CLng(BaseDifAporOS) < 0 Then
            NroError = "51"
            MsgErr = "El valor de Base para el cálculo diferencial de aporte es negativo. Verifique configuración en columnas 49 y 249 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no exceda de 13 Enteros
            BaseDifAporOSAux = Fix(BaseDifAporOS)
            If Len(BaseDifAporOSAux) > 13 Then
                NroError = "52"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columnas 49 y 249 valor 1, 2, 3, 4 y 5."
            End If
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", BaseDifAporOS, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(BaseDifAporOS, "#############0.00"))
        '****************************************************************************************************************

        '--------------------------------------------------------------------------------
        'Base para el cálculo diferencial de contribuciones de obra social y FSR (1)
        NroCol = 31
        NroError = ""
        MsgErr = ""
        NroConfCol = 50
        NroConfCol2 = 250
        Texto = "Base para el cálculo diferencial de contribuciones de obra social y FSR (1)"
        BaseDifContOS = 0
        If confRep(NroConfCol).confval <> "0" Then
            BaseDifContOS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "56"
            MsgErr = "No se encontró configuración para la columna 50 valor 1."
        End If
        
         'CONTROLO IM|I1|I2|I3 Y SUMO AL OBTENIDO DEL VALOR 1
        BaseDifContOS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                BaseDifContOS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                BaseDifContOS = BaseDifContOS + BaseDifContOS2
            End If
        Next
        
        If CLng(BaseDifContOS) < 0 Then
            NroError = "54"
            MsgErr = "El valor de Base para el cálculo diferencial de aporte es negativo. Verifique configuración en columnas 50 y 250 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no exceda de 13 Enteros
            BaseDifContOSAux = Fix(BaseDifContOS)
            If Len(BaseDifAporOSAux) > 13 Then
                NroError = "55"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columnas 50 y 250 valor 1, 2, 3, 4 y 5."
            End If
        End If
       
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", BaseDifContOS, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(BaseDifContOS, "#############0.00"))
        '****************************************************************************************************************

        '--------------------------------------------------------------------------------
        'Base para el cálculo diferencial Ley de Riesgos del Trabajo (1)
        NroCol = 32
        NroError = ""
        MsgErr = ""
        NroConfCol = 51
        NroConfCol2 = 251
        Texto = "Base para el cálculo diferencial Ley de Riesgos del Trabajo (1)"
        BaseDifLeyRieOS = 0
        If confRep(NroConfCol).confval <> "0" Then
            BaseDifLeyRieOS = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "59"
            MsgErr = "No se encontró configuración para la columna 51 valor 1."
        End If
        
        BaseDifLeyRieOS2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                BaseDifLeyRieOS2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                BaseDifLeyRieOS = BaseDifLeyRieOS + BaseDifLeyRieOS2
            End If
        Next
        
        If CLng(BaseDifLeyRieOS) < 0 Then
            NroError = "57"
            MsgErr = "El valor de Base para el cálculo diferencial Ley de Riesgos del trabajo es negativo. Verifique configuración en columnas 51 y 251 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no exceda de 13 Enteros
            BaseDifLeyRieOSAux = Fix(BaseDifLeyRieOS)
            If Len(BaseDifLeyRieOSAux) > 13 Then
                NroError = "58"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columnas 51 y 251 valor 1, 2, 3, 4 y 5."
            End If
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", BaseDifLeyRieOS, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(BaseDifLeyRieOS, "#############0.00"))
        '****************************************************************************************************************

        '--------------------------------------------------------------------------------
        'Remuneración maternidad para ANSeS
        NroCol = 33
        NroError = ""
        MsgErr = ""
        NroConfCol = 52
        NroConfCol2 = 252
        Texto = "Remuneración maternidad para ANSeS"
        RemMaterAnses = 0
        If confRep(NroConfCol).confval <> "0" Then
            RemMaterAnses = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "62"
            MsgErr = "No se encontró configuración para la columna 52 valor 1."
        End If
        
        RemMaterAnses2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                RemMaterAnses2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                RemMaterAnses = RemMaterAnses + RemMaterAnses2
            End If
        Next
                        
        If CLng(RemMaterAnses) < 0 Then
            NroError = "60"
            MsgErr = "El valor de Remuneración maternidad para ANSeS es negativo. Verifique configuración en columnas 52 y 252 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no exceda de 13 Enteros
            RemMaterAnsesAux = Fix(RemMaterAnses)
            If Len(RemMaterAnsesAux) > 13 Then
                NroError = "61"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columnas 52 y 252 valor 1, 2, 3, 4 y 5."
            End If
        End If
        
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", RemMaterAnses, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(RemMaterAnses, "#############0.00"))
        '****************************************************************************************************************

        '--------------------------------------------------------------------------------
        'Remuneración bruta
        NroCol = 34
        NroError = ""
        MsgErr = ""
        NroConfCol = 53
        NroConfCol2 = 253
        Texto = "Remuneración bruta"
        RemBruta = 0
        If confRep(NroConfCol).confval <> "0" Then
            RemBruta = getValoresLiq(confRep(NroConfCol).tipo, Ternro, confRep(NroConfCol).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
        Else
            NroError = "65"
            MsgErr = "No se encontró configuración para la columna 53 valor 1."
        End If
        
        RemBruta2 = 0
        For ax = 2 To 5
            MontoCant = confRep(NroConfCol2).confval & ax
            If ((confRep(NroConfCol).confval & ax) <> "0") Then
                RemBruta2 = getValoresLiq(confRep(NroConfCol).tipo2, Ternro, confRep(NroConfCol).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                RemBruta = RemBruta + RemBruta2
            End If
        Next
        
        If CLng(RemBruta) < 0 Then
            NroError = "63"
            MsgErr = "El valor de Remuneración bruta es negativo. Verifique configuración en columnas 53 y 253 valor 1, 2, 3, 4 y 5."
        Else
            'Controlo que no exceda de 13 Enteros
            RemBrutaAux = Fix(RemBruta)
            If Len(RemBrutaAux) > 13 Then
                NroError = "64"
                MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columnas 53 y 253 valor 1, 2, 3, 4 y 5."
            End If
        End If
        
        Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", RemBruta, MsgErr)
        '****************************************************************************************************************
        '----- INSERTO REGISTRO
        '****************************************************************************************************************
        Call InsReg(NroError, Registro, Texto, Format(RemBruta, "#############0.00"))
        '****************************************************************************************************************
        
        
        
        '--------------------------------------------------------------------------------
        'Base imponible 1,2,3,4,5,6,7,8,9 (Col 35  a 43)
        For a = 1 To 9
            NroCol = (34 + a)
            NroError = ""
            MsgErr = ""
            Texto = "Base imponible " & a
            BaseImpo = 0
            'NroConfCol = 53
            NroConfCol2 = 253
            If confRep((53 + a)).confval <> "0" Then
                BaseImpo = getValoresLiq(confRep((53 + a)).tipo, Ternro, confRep((53 + a)).confval, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
            Else
                If confRep((53 + a)).confval = "" Then
                    NroError = "68"
                    MsgErr = "No se encontró configuración para la columna " & (53 + a) & " valor 1."
                End If
            End If
            
            BaseImpo2 = 0
            NroConfCol2 = NroConfCol2 + 1
            For ax = 2 To 5
                MontoCant = confRep(NroConfCol2).confval & ax
                If ((confRep(NroConfCol).confval & ax) <> "0") Then
                    BaseImpo2 = getValoresLiq(confRep((53 + a)).tipo2, Ternro, confRep((53 + a)).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, MontoCant)
                    BaseImpo = BaseImpo + BaseImpo2
                End If
            Next
            
            
            If CLng(BaseImpo) < 0 Then
                NroError = "66"
                MsgErr = "El valor de Base imponible " & a & " es negativo. Verifique configuración en columnas " & (53 + a) & " y " & (253 + a) & " valor 1, 2, 3, 4 y 5."
            Else
                'Controlo que no exceda de 13 Enteros
                BaseImpoAux = Fix(BaseImpo)
                If Len(BaseImpoAux) > 13 Then
                    NroError = "67"
                    MsgErr = "El valor contiene más de 13 enteros. Verifique configuración en columna " & (53 + a) & " y " & (253 + a) & " valor 1, 2, 3, 4 y 5."
                End If
            End If
            
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", BaseImpo, MsgErr)
            '********************************************************
            '----- INSERTO REGISTRO
            '********************************************************
            'Call InsReg(NroError, Registro, Texto, Replace(FormatNumber(BaseImpo, 2), ",", ""))
            Call InsReg(NroError, Registro, Texto, Format(BaseImpo, "#############0.00"))
           
            '********************************************************
        Next
        

        '********************************************************
        '--------------- INSERTO REGISTRO DE CORTE --------------
        NroError = ""
        NroCol = 0
        Ternro = 0
        Call InsReg(NroError, Registro, "-1", "")
        '********************************************************
 
        
        '====================================================================================
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        '====================================================================================
        
        
        rsEmp.MoveNext
    Loop

    
    'ACTUALIZO CANTIDAD DE EMPLEADO DEL REGISTRO 01
    StrSql = "UPDATE rep_ar_libsuedigital_det SET replibsuedigval = " & cont & " WHERE replibsuedigreg = '01' AND bpronro =" & NroProcesoBatch & " AND replibsuedigcol =8"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    '---------------------------------------------------------------------------------------------------------------------------
Else

    Flog.writeline "No se encontraron datos a procesar"

End If



End Sub
Public Sub Registro05()
'==================================================================================================
'05:: Datos del trabajador de la empresa de servicios eventuales
'==================================================================================================
Dim EtiqNoinf As String
Dim EtiqBlanco As String
Dim EtiqSindatos  As String
Dim EmpleAProc As Long
Dim Auxiliar As String
Dim Formadepago As String
Dim rsEmp As New ADODB.Recordset
Dim rsFase As New ADODB.Recordset
Dim Longitud
Longitud = 50
Dim MsgErr As String
Dim cuil As String
Dim Registro As String
Dim CatProf As String
Dim PuestoDes As String
Dim FechaIng As String
Dim FechaEgr As String
Dim Remun As String
Dim cuit As String
Dim StrSqlAux As String
Dim UsaLiq As Boolean
Dim Pertenece As Boolean
Dim cliqnro As String
EtiqNoinf = "No informa"
EtiqBlanco = "Informa en Blanco"
EtiqSindatos = "No se encontraron datos."

Dim IncPorcAux As Currency


Registro = "05"

Flog.writeline ""
Flog.writeline "***************************************************************************************"
Flog.writeline "REGISTRO 05 :: Datos del trabajador de la empresa de servicios eventuales"
Flog.writeline "***************************************************************************************"
Flog.writeline ""

'CONTROLO SI BUSCA DATOS DE LIQ O DEL EMPLEADO
UsaLiq = False
If CLng(confRep(301).confval) = -1 Then
    UsaLiq = True
End If

'CICLO EMPLEADOS
StrSql = "SELECT "
StrSql = StrSql & " empleado.ternro,empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2"
StrSql = StrSql & ",empleado.empremu"
StrSql = StrSql & ",tercero.tersex,tercero.estcivnro,tercero.terfecnac,tercero.paisnro,tercero.nacionalnro,tercero.terfecing"
StrSql = StrSql & ",his_estructura.htetdesde , his_estructura.htethasta"
StrSql = StrSql & ",e2.empnro "

'If UsaLiq = True Then
'   StrSql = StrSql & " ,cabliq.cliqnro "
'End If

If confRep(300).confval = "1" Then 'SUCURSAL
    StrSql = StrSql & ",sucursal.ternro EveTernro "
    StrSqlAux = " INNER JOIN sucursal ON sucursal.estrnro = his_estructura.estrnro "
ElseIf confRep(300).confval = "10" Then 'EMPRESA
    StrSql = StrSql & ",empresa.ternro EveTernro "
    StrSqlAux = " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro "
ElseIf confRep(300).confval = "28" Then 'AGENCIA
    StrSql = StrSql & ",agencia.ternro EveTernro "
    StrSqlAux = " INNER JOIN agencia ON agencia.estrnro = his_estructura.estrnro "
End If


StrSql = StrSql & " FROM Empleado"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro =empleado.ternro"
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro =empleado.ternro AND tenro = " & confRep(300).confval

StrSql = StrSql & StrSqlAux

StrSql = StrSql & " INNER JOIN his_estructura he2 ON he2.ternro =empleado.ternro AND he2.tenro = 10"
StrSql = StrSql & " INNER JOIN empresa e2 ON e2.estrnro = he2.estrnro "


'If UsaLiq = True Then
'   StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro AND cabliq.pronro IN (" & ListParam.pronro & ")"
'End If
StrSql = StrSql & " WHERE "
StrSql = StrSql & " ( his_estructura.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & " AND (his_estructura.htethasta >=" & ConvFecha(DatosPeriodo.pliqdesde) & " OR his_estructura.htethasta IS NULL))"
StrSql = StrSql & " AND ( he2.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & " AND (he2.htethasta >=" & ConvFecha(DatosPeriodo.pliqdesde) & " OR he2.htethasta IS NULL))"



StrSql = StrSql & " ORDER BY empleg"
OpenRecordset StrSql, rsEmp

EmpleAProc = rsEmp.RecordCount
If EmpleAProc = 0 Then
   EmpleAProc = 1
End If
IncPorc = ProgresoAux / EmpleAProc
'IncPorcAux = ProgresoAux / EmpleAProc

'Flog.writeline "IncPorc:" & IncPorcAux
'Flog.writeline "ProgresoAux:" & ProgresoAux
'Flog.writeline "EmpleAProc:" & EmpleAProc

If Not rsEmp.EOF Then

    'IncPorc = 1
    Do While Not rsEmp.EOF
        
        Ternro = rsEmp!Ternro
   
        Flog.writeline ""
        Flog.writeline String(Longitud, "-")
        Flog.writeline "-- EMPLEADO : " & rsEmp!Empleg
        'Flog.Writeline String(Longitud, "-")
        
        Pertenece = True
        If UsaLiq = True Then 'Solo entro si la opcón de usar Liquidación esta en SI
            'Limpio cliqnro
            cliqnro = "0"
            '---------------------------------------------------------------------------------------------
            'BUSCO LOS PROCESOS ASOCIADOS A LA EMPRESA DEL EMPLEADO Y PERIODO SELECCIONADO EN EL FILTRO.
            'StrSql = "SELECT pronro FROM proceso WHERE pliqnro = " & ListParam.Pliqnro & " AND empnro =" & rsEmp!Empnro
            StrSql = "SELECT cliqnro FROM cabliq WHERE cabliq.pronro IN ("
            StrSql = StrSql & "SELECT pronro FROM proceso WHERE pliqnro = " & ListParam.Pliqnro & " AND empnro =" & rsEmp!Empnro
            StrSql = StrSql & ")"
            StrSql = StrSql & " AND empleado = " & Ternro
            OpenRecordset StrSql, rsFase
            If Not rsFase.EOF Then
                Do While Not rsFase.EOF
                    cliqnro = cliqnro & "," & rsFase!cliqnro
                    rsFase.MoveNext
                Loop
            Else
                Pertenece = False
                Flog.writeline " -- NO SE ENCONTRÓ LIQUIDACION."
            End If
            rsFase.Close
        End If
        Flog.writeline String(Longitud, "-")
        '---------------------------------------------------------------------------------------------
        'Si el empleado esta liquidado lo proceso, sino continuo
        If Pertenece = True Then
            '--------------------------------------------------------------------------------
            ' - Identificacion del tipo de registro
            NroCol = 1
            NroError = ""
            MsgErr = ""
            Texto = "Identificacion del tipo de registro"
            'Flog.Writeline Texto
            'Flog.Writeline Registro
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & Registro
            '**********************************************
            '----- INSERTO REGISTRO
            '**********************************************
            Call InsReg(NroError, Registro, Texto, Registro)
            '**********************************************
            
            
            '--------------------------------------------------------------------------------
            ' - CUIL del trabajador
            NroCol = 2
            MsgErr = ""
            Texto = "CUIL del trabajador"
            'Flog.Writeline Texto
            If IsNumeric(confRep(20).confval) And confRep(20).confval <> 0 Then
                cuil = getTerDoc(Ternro, CLng(confRep(20).confval), 1, 1)
            Else
                'NO SE ENCONTRO CONFIGURACIÓN PARA LA COLUMNA 1. Documentos de la empresa
                NroError = "5"
                MsgErr = "No se encontró configuración válida en el confrep columna 20 valor 1. Se setea Default => 10"
                cuil = getTerDoc(Ternro, 10, 1, 1)
            End If
            'CONTROLO CUIL
            If (cuil = "") Then
                NroError = "6"
                MsgErr = "No se ha encontrado el CUIL."
            ElseIf Len(cuil) > 11 Then
                NroError = "7"
                MsgErr = "El CUIL debe contener 11 dígitos (13 Incluyendo los guiones medios (-)" & "(" & cuil & ")"
            End If
            
            If MsgErr <> "" And cuil <> "" Then
                Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & MsgErr
            Else
                Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(cuil = "", MsgErr, cuil)
            End If
            'Flog.writeline Texto
            'Flog.Writeline cuil
            
            '**********************************************
            '----- INSERTO REGISTRO
            '**********************************************
            Call InsReg(NroError, Registro, Texto, cuil)
            '**********************************************
            
            
            '--------------------------------------------------------------------------------
            'Categoría profesional
            NroCol = 3
            NroError = ""
            MsgErr = ""
            Texto = "Categoría profesional"
           'CatProf = UCase(getEmpleadoEstr(Ternro, confRep(300).confval2, "ESTRCODEXT", DatosPeriodo.pliqdesde, DatosPeriodo.pliqhasta, ""))
            CatProf = getCodMiSimpl(Ternro, IIf(EsNulo(confRep(302).confval), "0", confRep(302).confval), confRep(300).confval2, DatosPeriodo.pliqdesde, DatosPeriodo.pliqhasta)
            '-1 -> No se encontró el codigo interno | '0 --> Sin estructura asociada
            If CatProf = "-1" Then
                NroError = "80"
                'MsgErr = "No se encontró Código externo para la estructura configurada. Configuración del reporte en columna 300 valor 2."
                MsgErr = "No se encontró tipo de código Mi Simplificación. Por defecto se Informará [000000]."
                CatProf = "000000"
            ElseIf CatProf = "0" Then
                NroError = "81"
                CatProf = Left(CatProf, 6)
                'MsgErr = "El Código Externo de la estructura no puede exceder 6 caracteres. El texto se formatea a 6." & "(" & CatProf & ")"
                MsgErr = "No se encontró el Tipo de Estructura asociado al empleado. Por defecto se Informará [000000]."
                CatProf = "000000"
            Else
                If Len(CatProf) > 6 Then
                    NroError = "91"
                    CatProf = Left(CatProf, 6)
                    'MsgErr = "El Código Externo de la estructura no puede exceder 6 caracteres. El texto se formatea a 6." & "(" & CatProf & ")"
                    MsgErr = "El Tipo de Código no puede exceder 6 caracteres. El texto se formatea a 6." & "(" & CatProf & ")"
                End If
            End If
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, CatProf)
            '********************************************************************************************
            '----- INSERTO REGISTRO
            '********************************************************************************************
            Call InsReg(NroError, Registro, Texto, CatProf)
            '********************************************************************************************
            
            
            '--------------------------------------------------------------------------------
            'Puesto desempeñado
            NroCol = 4
            NroError = ""
            MsgErr = ""
            Texto = "Puesto desempeñado"
            'PuestoDes = UCase(getEmpleadoEstr(Ternro, confRep(300).confval3, "ESTRCODEXT", DatosPeriodo.pliqdesde, DatosPeriodo.pliqhasta, ""))
            PuestoDes = getCodMiSimpl(Ternro, IIf(EsNulo(confRep(302).confval), "0", confRep(302).confval), confRep(300).confval3, DatosPeriodo.pliqdesde, DatosPeriodo.pliqhasta)
            '-1 -> No se encontró el codigo interno | '0 --> Sin estructura asociada
            If PuestoDes = "-1" Then
                NroError = "82"
                'MsgErr = "No se encontró Código externo para la estructura configurada. Configuración del reporte en columna 300 valor 3."
                MsgErr = "No se encontró tipo de código Mi Simplificación. Por defecto se Informará [0000]."
                PuestoDes = "0000"
            ElseIf PuestoDes = "0" Then
                NroError = "83"
                MsgErr = "No se encontró el Tipo de Estructura asociado al empleado. Por defecto se Informará [0000]."
                PuestoDes = "0000"
            Else
                If Len(PuestoDes) > 4 Then
                    NroError = "92"
                    PuestoDes = Left(PuestoDes, 4)
                    MsgErr = "El Tipo de Código no puede exceder 4 caracteres. El texto se formatea a 4." & "(" & PuestoDes & ")"
                End If
            End If
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, PuestoDes)
            '********************************************************************************************
            '----- INSERTO REGISTRO
            '********************************************************************************************
            Call InsReg(NroError, Registro, Texto, PuestoDes)
            '********************************************************************************************
            
            'BUSCO LA FASES REAL DEL EMPLEADO
            StrSql = " SELECT altfec,bajfec FROM fases WHERE real = -1 AND empleado = " & Ternro
            StrSql = StrSql & " AND (altfec <= " & ConvFecha(DatosPeriodo.pliqhasta) & ") "
            StrSql = StrSql & " AND ((bajfec >= " & ConvFecha(DatosPeriodo.pliqdesde) & ")"
            StrSql = StrSql & " OR (bajfec is null))"
            StrSql = StrSql & " ORDER BY altfec"
            OpenRecordset StrSql, rsFase
            FechaIng = ""
            FechaEgr = ""
            If Not rsFase.EOF Then
                FechaIng = IIf(EsNulo(rsFase!altfec), "", rsFase!altfec)
                FechaEgr = IIf(EsNulo(rsFase!bajfec), "", rsFase!bajfec)
            End If
            
            '--------------------------------------------------------------------------------
            'Fecha de ingreso
            'AAAAMMDD - Es la fecha en la que inicia la prestación de servicios en la Empresa Usuaria
            NroCol = 5
            NroError = ""
            MsgErr = ""
            Texto = "Fecha de ingreso"
            'FechaIng = ""
            'If Not EsNulo(rsEmp!htetdesde) Then
            If Not EsNulo(FechaIng) Then
                'FechaIng = Format_Data(rsEmp!htetdesde, "AAAAMMDD")
                FechaIng = Format_Data(FechaIng, "AAAAMMDD")
            Else
                NroError = "84"
                'MsgErr = "No se encontró Fecha de inicio para la estructura configurada."
                MsgErr = "No se encontró Fecha de inicio. (Se toma la configurada como REAL)"
            End If
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", FechaIng, MsgErr)
            '********************************************************************************************
            '----- INSERTO REGISTRO
            '********************************************************************************************
            Call InsReg(NroError, Registro, Texto, FechaIng)
            '********************************************************************************************
            
            '--------------------------------------------------------------------------------
            'Fecha de Egreso
            'Formato: AAAAMMDD. Es la fecha en la que finaliza la prestación de servicios en la Empresa Usuaria
            NroCol = 6
            NroError = ""
            MsgErr = ""
            Texto = "Fecha de Egreso"
            'FechaEgr = ""
            'If Not EsNulo(rsEmp!htethasta) Then
            If Not EsNulo(FechaEgr) Then
                'FechaEgr = Format_Data(rsEmp!htethasta, "AAAAMMDD")
                FechaEgr = Format_Data(FechaEgr, "AAAAMMDD")
            Else
                NroError = "85"
                'MsgErr = "No se encontró Fecha de Egreso para la estructura configurada."
                MsgErr = "2"
                'No se encontró Fecha de inicio.
            End If
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr = "", FechaEgr, MsgErr)
            '********************************************************************************************
            '----- INSERTO REGISTRO
            '********************************************************************************************
            Call InsReg(NroError, Registro, Texto, FechaEgr)
            '********************************************************************************************
                
            '--------------------------------------------------------------------------------
            'Remuneración
            NroCol = 7
            NroError = ""
            MsgErr = ""
            Texto = "Remuneración"
            Remun = "0"
            If UsaLiq = False Then
                If Not EsNulo(rsEmp!empremu) Then
                    Remun = rsEmp!empremu
                Else
                    NroError = "86"
                    MsgErr = "No se encontró remuneración del trabajador. Se informa default [0]."
                End If
            Else
                'BUSCO AC/ACL
                If confRep(301).confval2 = "0" Then
                    NroError = "90"
                    MsgErr = "No se encontró valor. Se setea default [0]. Ver configuración columna 301 valor 2."
                Else
                   ' Remun = getValoresLiq(confRep(301).tipo2, Ternro, confRep(301).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, rsEmp!Cliqnro, 0)
                    Remun = getValoresLiq(confRep(301).tipo2, Ternro, confRep(301).confval2, DatosPeriodo.pliqmes, DatosPeriodo.pliqanio, cliqnro, 0)
                    'pero la funcion le retorna
                    If CLng(Remun) = 0 Then
                        NroError = "86"
                        MsgErr = "No se encontró remuneración del trabajador. Se informa default [0]."
                    End If
                End If
            End If
            
            
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, "")
            '********************************************************************************************
            '----- INSERTO REGISTRO
            '********************************************************************************************
            Call InsReg(NroError, Registro, Texto, Format(Remun, "#####0.00"))
            '********************************************************************************************
    
            '--------------------------------------------------------------------------------
            ' - CUIT del empleador
            NroCol = 8
            Texto = "CUIT del empleador"
            NroError = ""
            MsgErr = ""
            cuit = ""
            If IsNumeric(confRep(10).confval) And confRep(10).confval <> 0 Then
                cuit = getTerDoc(rsEmp!EveTernro, CLng(confRep(10).confval), 1, 1)
            Else
                'NO SE ENCONTRO CONFIGURACIÓN PARA LA COLUMNA 1. Documentos de la empresa
                NroError = "1"
                MsgErr = "No se encontró configuración válida en el confrep columna 10 valor 1. Se setea Default => 6"
                cuit = getTerDoc(rsEmp!EveTernro, 6, 1, 1)
            End If
            'CONTROLO CUIT
            If (cuit = "") Then
                NroError = "2"
                MsgErr = "No se ha encontrado el CUIT."
            ElseIf Len(cuit) > 11 Or Len(cuit) < 11 Then
                NroError = "3"
                MsgErr = "El CUIT debe contener 11 dígitos (13 Incluyendo los guiones medios (-)" & "(" & cuit & ")"
            End If
            Flog.writeline Format_StrLR(Texto, Longitud, "R", True, " ") & ": " & IIf(MsgErr <> "", MsgErr, cuit)
            '**********************************************
            '----- INSERTO REGISTRO
            '**********************************************
            Call InsReg(NroError, Registro, Texto, cuit)
            '**********************************************
            
            
            
            '********************************************************
            '--------------- INSERTO REGISTRO DE CORTE --------------
            NroError = ""
            NroCol = 0
            Ternro = 0
            Call InsReg(NroError, Registro, "-1", "")
            '********************************************************
        
        End If 'Cierra Pertence
        
        '====================================================================================
        TiempoAcumulado = GetTickCount
        'Progreso = Progreso + IncPorc
        Progreso = Progreso + IncPorcAux
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        '====================================================================================
        rsEmp.MoveNext
    Loop
Else
    Flog.writeline "No se encontraron datos a procesar"
End If



End Sub

Public Sub GetSitdeRevista(ByRef CodSitRevista As String, ByRef CodSitRevista1 As String, ByRef Aux_diainisr1 As String, ByRef CodSitRevista2 As String, ByRef Aux_diainisr2 As String, ByRef CodSitRevista3 As String, ByRef Aux_diainisr3 As String, ByRef NroError, ByRef MsgErr, ByRef MsgInfo)
    Dim Aux_Cod
    'Flog.Writeline "Buscar Situacion de Revista Actual"
    ' ----------------------------------------------------------------
    'Buscar Situacion de Revista Actual
    StrSql = " SELECT * FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 30 AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(DatosPeriodo.pliqhasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(DatosPeriodo.pliqdesde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"

    
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    OpenRecordset StrSql, rs_Estructura
    'Flog.Writeline "inicializo"
    CodSitRevista1 = ""
    Aux_diainisr1 = ""
    CodSitRevista2 = ""
    Aux_diainisr2 = ""
    CodSitRevista3 = ""
    Aux_diainisr3 = ""
    Select Case rs_Estructura.RecordCount
        Case 0:
                'Si no tiene situación de revista, busca si tiene alguna estructura 'REV' en el confrep
                NroError = ""
                MsgErr = ""
                MsgInfo = ""
                CodSitRevista1 = Buscar_SituacionRevistaConfig(confRep(42).confval, confRep(42).confval2)
                Aux_diainisr1 = 1
                If CStr(CodSitRevista1) <> "0" Then
                    MsgInfo = "Se asignó la situación de revista del confrep: " & CodSitRevista1
                    CodSitRevista = CodSitRevista1
                Else
                    NroError = "21"
                    MsgErr = "No hay situaciones de revista. Se setea en [0]."
                    CodSitRevista = CodSitRevista1
                End If
        Case 1:
            NroError = ""
            MsgErr = ""
            MsgInfo = ""
           
            'Aux_Cod_sitr1 = rs_Estructura!estrcodext
            StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
            StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
            OpenRecordset StrSql, rs_Estr_cod
            If Not rs_Estr_cod.EOF Then
                CodSitRevista1 = Left(CStr(rs_Estr_cod!nrocod), 2)
            Else
                NroError = "20"
                MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
                CodSitRevista1 = 1
            End If
            
            If rs_Estructura!htetdesde < DatosPeriodo.pliqdesde Then
                Aux_diainisr1 = 1
            Else
                Aux_diainisr1 = Day(rs_Estructura!htetdesde)
            End If
            If CInt(Aux_diainisr1) > Day(DatosPeriodo.pliqhasta) Then
                Aux_diainisr1 = CStr(Day(DatosPeriodo.pliqhasta))
            End If
            
            CodSitRevista = CodSitRevista1
            MsgInfo = "Hay 1 Situación de revista."
    Case 2:
        'Primer situacion
        'Aux_Cod_sitr1 = rs_Estructura!estrcodext
        NroError = ""
        MsgErr = ""
        MsgInfo = ""
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro =" & confRep(42).confval3
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            CodSitRevista1 = Left(CStr(rs_Estr_cod!nrocod), 2)
        Else
            NroError = "20"
            MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
            CodSitRevista1 = 1
        End If
        
        If rs_Estructura!htetdesde < DatosPeriodo.pliqdesde Then
            Aux_diainisr1 = 1
        Else
            Aux_diainisr1 = Day(rs_Estructura!htetdesde)
        End If
        
        If CInt(Aux_diainisr1) > Day(DatosPeriodo.pliqhasta) Then
            Aux_diainisr1 = CStr(Day(DatosPeriodo.pliqhasta))
        End If
        
        'Siguiente situacion
        rs_Estructura.MoveNext
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
        OpenRecordset StrSql, rs_Estr_cod
        If CodSitRevista1 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then
                If Not rs_Estr_cod.EOF Then
                    CodSitRevista2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                Else
                    NroError = "20"
                    MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
                    CodSitRevista2 = 1
                End If
                'Aux_Cod_sitr2 = rs_Estructura!estrcodext
                Aux_diainisr2 = Day(rs_Estructura!htetdesde)

                If CInt(Aux_diainisr2) > Day(DatosPeriodo.pliqhasta) Then
                    Aux_diainisr2 = CStr(Day(DatosPeriodo.pliqhasta))
                End If
                CodSitRevista = CodSitRevista2
                'Flog.Writeline "Hay 2 situaciones de revista"
                MsgInfo = "Hay 2 situaciones de revista."
        Else
            'Si es la misma sit de revista ==> le asigno la anterior
            CodSitRevista = CodSitRevista1
        End If
        
    Case 3:
        'Primer situacion (1)
        NroError = ""
        MsgErr = ""
        MsgInfo = ""
        'Aux_Cod_sitr1 = rs_Estructura!estrcodext
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            CodSitRevista1 = Left(CStr(rs_Estr_cod!nrocod), 2)
        Else
            NroError = "22"
            MsgErr = "No se encontró el código interno para la Situación de Revista.Se setea default [1]."
            CodSitRevista1 = 1
        End If
        
        If rs_Estructura!htetdesde < DatosPeriodo.pliqdesde Then
            Aux_diainisr1 = 1
        Else
            Aux_diainisr1 = Day(rs_Estructura!htetdesde)
        End If
        'FGZ - 08/07/2005
        If CInt(Aux_diainisr1) > Day(DatosPeriodo.pliqhasta) Then
            Aux_diainisr1 = CStr(Day(DatosPeriodo.pliqhasta))
        End If
        
        'siguiente situacion (2)
        rs_Estructura.MoveNext
        'Aux_Cod_sitr2 = rs_Estructura!estrcodext
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            If CodSitRevista1 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then 'Agregado ver 1.31
                If Not rs_Estr_cod.EOF Then
                    CodSitRevista2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                Else
                    NroError = "22"
                    MsgErr = "No se encontró el código interno para la Situación de Revista.Se setea default [1]."
                    CodSitRevista2 = 1
                End If
                Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                
                If CInt(Aux_diainisr2) > Day(DatosPeriodo.pliqhasta) Then
                    Aux_diainisr2 = CStr(Day(DatosPeriodo.pliqhasta))
                End If
            Else
                'Si es la misma sit de revista ==> le asigno la anterior
                CodSitRevista = CodSitRevista1
            End If
        Else
            NroError = "22"
            MsgErr = "No se encontró el código interno para la Situación de Revista.Se setea default [1]."
            CodSitRevista2 = 1
            Aux_diainisr2 = Day(rs_Estructura!htetdesde)
            If CInt(Aux_diainisr2) > Day(DatosPeriodo.pliqhasta) Then
                Aux_diainisr2 = CStr(Day(DatosPeriodo.pliqhasta))
            End If
        End If
        'siguiente situacion (3)
        rs_Estructura.MoveNext
        'Aux_Cod_sitr3 = rs_Estructura!estrcodext
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            If CodSitRevista2 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then 'Agregado ver 1.31
                If Not rs_Estr_cod.EOF Then
                    If CodSitRevista2 <> "" Then
                        CodSitRevista3 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        CodSitRevista2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    End If
                Else
                    NroError = "22"
                    MsgErr = "No se encontró el código interno para la Situación de Revista.Se setea default [1]."
                    CodSitRevista3 = 1
                End If
                If CodSitRevista3 <> "" Then
                    Aux_diainisr3 = Day(rs_Estructura!htetdesde)
                Else
                    Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                End If
                'FGZ - 08/07/2005
                If Aux_diainisr3 <> "" Then
                    If CInt(Aux_diainisr3) > Day(DatosPeriodo.pliqhasta) Then
                        Aux_diainisr3 = CStr(Day(DatosPeriodo.pliqhasta))
                    End If
                    CodSitRevista = CodSitRevista3
                Else
                    If CInt(Aux_diainisr2) > Day(DatosPeriodo.pliqhasta) Then
                        Aux_diainisr2 = CStr(Day(DatosPeriodo.pliqhasta))
                    End If
                    CodSitRevista = CodSitRevista2
                End If
                MsgInfo = "Hay 3 situaciones de revista."
            Else
                'Si es la misma sit de revista ==> le asigno la anterior
                If CodSitRevista2 <> "" Then
                    CodSitRevista = CodSitRevista2
                Else
                    CodSitRevista = CodSitRevista1
                End If
            End If
        Else
            NroError = "22"
            MsgErr = "No se encontró el código interno para la Situación de Revista.Se setea default [1]."
            CodSitRevista3 = 1
            If CodSitRevista3 <> "" Then
                Aux_diainisr3 = Day(rs_Estructura!htetdesde)
            Else
                Aux_diainisr2 = Day(rs_Estructura!htetdesde)
            End If
            'FGZ - 08/07/2005
            If Aux_diainisr3 <> "" Then
                If CInt(Aux_diainisr3) > Day(DatosPeriodo.pliqhasta) Then
                    Aux_diainisr3 = CStr(Day(DatosPeriodo.pliqhasta))
                End If
                CodSitRevista = CodSitRevista3
            Else
                If CInt(Aux_diainisr2) > Day(DatosPeriodo.pliqhasta) Then
                    Aux_diainisr2 = CStr(Day(DatosPeriodo.pliqhasta))
                End If
                CodSitRevista = CodSitRevista2
            End If
            MsgInfo = "Hay 3 situaciones de revista."
        End If
        
        Case Else 'mas de tres situaciones ==> toma las ultimas tres pero verifica que no haya situaciones iguales en dif periodos
            NroError = ""
            MsgErr = ""
            MsgInfo = ""
             If Not rs_Estructura.EOF Then
                Dim k
                k = 0
                Do While Not rs_Estructura.EOF
                   If (k = 0) Then
                        ReDim Preserve Arrcod(k)
                        ReDim Preserve Arrdia(k)
                        'Arrcod(k) = rs_Estructura!estrcodext 'NG - V 1.7
                        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                        StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
                        OpenRecordset StrSql, rs_Estr_cod
                        If Not rs_Estr_cod.EOF Then
                            If EsNulo(rs_Estr_cod!nrocod) = False Then
                                Arrcod(k) = Left(CStr(rs_Estr_cod!nrocod), 2)
                            Else
                                NroError = "20"
                                MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
                                Arrcod(k) = 1
                            End If
                        Else
                            NroError = "20"
                            MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
                            Arrcod(k) = 1
                        End If

                       
                        If rs_Estructura!htetdesde < DatosPeriodo.pliqdesde Then
                           Arrdia(k) = 1
                        Else
                           Arrdia(k) = Day(rs_Estructura!htetdesde)
                        End If
                        'fin
                        k = k + 1
                   Else
                        
                        'NG - v 1.7-- ------------------------------------------------------------------------
                        Aux_Cod = ""
                        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                        StrSql = StrSql & " AND tcodnro = " & confRep(42).confval3
                        OpenRecordset StrSql, rs_Estr_cod
                        If Not rs_Estr_cod.EOF Then
                            If EsNulo(rs_Estr_cod!nrocod) = False Then
                                Aux_Cod = Left(CStr(rs_Estr_cod!nrocod), 2)
                            Else
                                NroError = "20"
                                MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
                                Aux_Cod = 1
                            End If
                        Else
                            NroError = "20"
                            MsgErr = "No se encontró el codigo interno para la Situación de Revista. Se setea default [1]."
                            Aux_Cod = 1
                        End If
                        'NG - v 1.7-- ------------------------------------------------------------------------

                   
                       If Arrcod(k - 1) <> Aux_Cod Then
                            ReDim Preserve Arrcod(k)
                            ReDim Preserve Arrdia(k)
                            'Arrcod(k) = rs_Estructura!estrcodext - NG v 1.7--
                            Arrcod(k) = Aux_Cod
                            If rs_Estructura!htetdesde < DatosPeriodo.pliqdesde Then
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
                CodSitRevista3 = Arrcod(UBound(Arrcod))
                CodSitRevista2 = Arrcod(UBound(Arrcod) - 1)
                CodSitRevista1 = Arrcod(UBound(Arrcod) - 2)
                Aux_diainisr3 = Arrdia(UBound(Arrdia))
                Aux_diainisr2 = Arrdia(UBound(Arrdia) - 1)
                Aux_diainisr1 = Arrdia(UBound(Arrdia) - 2)
             End If
            
             If UBound(Arrcod) = 1 Then
                CodSitRevista3 = ""
                CodSitRevista2 = Arrcod(UBound(Arrcod))
                CodSitRevista1 = Arrcod(UBound(Arrcod) - 1)
                Aux_diainisr3 = ""
                Aux_diainisr2 = Arrdia(UBound(Arrdia))
                Aux_diainisr1 = Arrdia(UBound(Arrdia) - 1)
             End If
            
             If UBound(Arrcod) = 0 Then
                CodSitRevista3 = ""
                CodSitRevista2 = ""
                CodSitRevista1 = Arrcod(UBound(Arrcod))
                Aux_diainisr3 = ""
                Aux_diainisr2 = ""
                Aux_diainisr1 = Arrdia(UBound(Arrdia))
             End If
             
             CodSitRevista = Arrcod(UBound(Arrcod))
            'fin
             MsgInfo = "Hay más de 3 situaciones de revista."
        End Select
        
        'FGZ - 28/12/2004
        'No puede haber situaciones de revista iguales consecutivas.
        'Antes ese caso, me quedo con la primera de las iguales y consecutivas
        If CodSitRevista3 = CodSitRevista2 Then
            'Elimino la situacion de revista 3
            CodSitRevista3 = ""
            CodSitRevista3 = ""
        End If
        If CodSitRevista2 = CodSitRevista1 Then
            'Elimino la situacion de revista 2 y la 3 la pongo en la 2
            CodSitRevista2 = CodSitRevista3
            CodSitRevista2 = CodSitRevista3
            
            CodSitRevista3 = ""
            Aux_diainisr3 = ""
        End If
       
End Sub

Public Sub InsReg(ByVal NroError As String, ByVal Registro As String, ByVal CampoNom As String, ByVal CampoVal As String)
   
    '----------------------------------------------------------
    'CONTROLO EL TIPO DE ERROR PARA MODIFICAR EN LA CABECERA
    Call ErrCabecera
    '----------------------------------------------------------
   'INSERTA WARNING Y ERRORES
    StrSql = "INSERT INTO rep_ar_libsuedigital_det (bpronro,ternro,replibsuedigreg,replibsuedigcol,replibsuedigerr,replibsuedignom,replibsuedigval) "
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & NroProcesoBatch
    StrSql = StrSql & "," & IIf(Ternro = 0, "NULL", Ternro)
    StrSql = StrSql & "," & "'" & Registro & "'"
    StrSql = StrSql & "," & NroCol
    StrSql = StrSql & "," & IIf(EsNulo(NroError), "NULL", "'" & NroError & "'")
    StrSql = StrSql & "," & IIf(EsNulo(CampoNom), "NULL", "'" & CampoNom & "'")
    StrSql = StrSql & "," & IIf(EsNulo(CampoVal), "NULL", "'" & CampoVal & "'")
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub



Public Sub ArmoDatosConfrep()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtiene configuración del confrep
' Autor      : Gonzalez Nicolás
' Fecha      : 18/06/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim contador As Long
    
StrSql = "SELECT confetiq,confnrocol,conftipo,conftipo2,conftipo3,conftipo4,conftipo5"
StrSql = StrSql & ",confval,confval2,confval3,confval4,confval5 "
StrSql = StrSql & " FROM confrepAdv "
StrSql = StrSql & " WHERE repnro = 491 "
'StrSql = StrSql & " AND confnrocol >= 10 AND confnrocol <=" & topeArrCrep
StrSql = StrSql & " ORDER BY confrepAdv.confnrocol ASC"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Texto = "No se encontro configuracion del reporte"
    HuboError = True
Else
    Do While Not rs_Confrep.EOF
        contador = CLng(rs_Confrep("confnrocol"))
        If contador > topeArreglo Then
           Flog.writeline " ERROR:La columna del confrep supera el limite establecido."
        Else
            'Guarda el resto de las columnas
            confRep(contador).Etiqueta = rs_Confrep("confetiq")
            'TIPOS
            If Not EsNulo(rs_Confrep("conftipo")) Then
                confRep(contador).tipo = rs_Confrep("conftipo")
            Else
                confRep(contador).tipo = ""
            End If
            
            If Not EsNulo(rs_Confrep("conftipo2")) Then
                confRep(contador).tipo2 = rs_Confrep("conftipo2")
            Else
                confRep(contador).tipo2 = ""
            End If
            
            If Not EsNulo(rs_Confrep("conftipo3")) Then
                confRep(contador).tipo3 = rs_Confrep("conftipo3")
            Else
                confRep(contador).tipo3 = ""
            End If

            If Not EsNulo(rs_Confrep("conftipo4")) Then
                confRep(contador).tipo4 = rs_Confrep("conftipo4")
            Else
                confRep(contador).tipo4 = ""
            End If

            If Not EsNulo(rs_Confrep("conftipo5")) Then
                confRep(contador).tipo5 = rs_Confrep("conftipo5")
            Else
                confRep(contador).tipo5 = ""
            End If

            '----------------------------------------------------
            'VALORES
            '----------------------------------------------------
            If Not EsNulo(rs_Confrep("confval")) Then
                confRep(contador).confval = rs_Confrep("confval")
            Else
                confRep(contador).confval = ""
            End If
            If Not EsNulo(rs_Confrep("confval2")) Then
                confRep(contador).confval2 = rs_Confrep("confval2")
            Else
                confRep(contador).confval2 = ""
            End If
            
            If Not EsNulo(rs_Confrep("confval3")) Then
                confRep(contador).confval3 = rs_Confrep("confval3")
            Else
                confRep(contador).confval3 = ""
            End If
            
            If Not EsNulo(rs_Confrep("confval4")) Then
                confRep(contador).confval4 = rs_Confrep("confval4")
            Else
                confRep(contador).confval4 = ""
            End If

            If Not EsNulo(rs_Confrep("confval5")) Then
                confRep(contador).confval5 = rs_Confrep("confval5")
            Else
                confRep(contador).confval5 = ""
            End If

            
        End If
        rs_Confrep.MoveNext
    Loop
End If
rs_Confrep.Close
End Sub

Public Sub ControlFases(ByVal Empleg As Long, ByRef Fecha_Inicio_Fase As Date, ByRef Fecha_Fin_Fase As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Controla las fases del empleado
' Autor      : Gonzalez Nicolás
' Fecha      : 24/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
   
   'Flog.writeline String(50, "-")
   Flog.writeline " CONTROLO FASES CONTRA PERDIODO SIJP"
   Flog.writeline String(50, "-")
   Dim FechaAuxAsignada  As Boolean
    StrSql = " SELECT * FROM fases WHERE real = -1 AND empleado = " & Ternro
    StrSql = StrSql & " AND (altfec <= " & ConvFecha(DatosPeriodo.pliqhasta) & ") "
    StrSql = StrSql & " AND ((bajfec >= " & ConvFecha(DatosPeriodo.pliqdesde) & ")"
    StrSql = StrSql & " OR (bajfec is null ))"
    StrSql = StrSql & " ORDER BY altfec"
    OpenRecordset StrSql, rs_fases
    ' Creo el Select para verificar si el empleado tiene un Contrato de tipo 11 (Afip)
    StrSql = "SELECT * FROM empleado e,his_estructura he,estructura es,estr_cod ec, tipocod tc"
    StrSql = StrSql & " WHERE e.empleg =" & Empleg
    StrSql = StrSql & " AND nrocod = '11' AND es.tenro = 18"
    StrSql = StrSql & " AND he.ternro = e.ternro"
    StrSql = StrSql & " AND es.estrnro = he.estrnro"
    StrSql = StrSql & " AND es.estrnro = ec.estrnro"
    StrSql = StrSql & " AND ec.tcodnro = tc.tcodnro"
    OpenRecordset StrSql, rs
    'Si dicho empleado tiene dicha estructura, asigno como fecha de inicio de Fase, el inicio del periodo, más allá de que el empleado tenga una fase abierta en dicho mes
    'Según resolución AFIP para el SICORE
    If Not rs.EOF Then
         Fecha_Inicio_Fase = DatosPeriodo.pliqdesde
         Fecha_Fin_Fase = DatosPeriodo.pliqhasta
    Else
         If rs_fases.RecordCount > 1 Then rs_fases.MoveFirst
         If rs_fases.RecordCount > 0 Then
                 Flog.writeline UCase("Comienza proceso de comparación de fechas de fases con las del período del SIJP")
                 Do While Not rs_fases.EOF
                     'Asigno la fecha de alta de la fase si es mayor a la del periodo
                     Fecha_Inicio_Fase = IIf(rs_fases!altfec > DatosPeriodo.pliqdesde, rs_fases!altfec, DatosPeriodo.pliqdesde)
                     Flog.writeline UCase("Asigno a fecha de inicio de fase el valor ") & Fecha_Inicio_Fase
                     If Not EsNulo(rs_fases!bajfec) Then
                         Flog.writeline UCase("El valor de fecha de baja no es nulo")
                         'Asigno la fecha de baja de la fase si es menor a la del periodo
                         Fecha_Fin_Fase = IIf(rs_fases!bajfec < DatosPeriodo.pliqhasta, rs_fases!bajfec, DatosPeriodo.pliqhasta)
                         Flog.writeline UCase("Asigno a fecha de fin de fase el valor ") & Fecha_Fin_Fase
                     Else
                         Fecha_Fin_Fase = DatosPeriodo.pliqhasta
                         Flog.writeline UCase("El valor de fecha de baja es nulo")
                         Flog.writeline UCase("El valor asignado a la Fecha de Fin de Fase es ") & Fecha_Fin_Fase
                     End If
                     'Aux_fecha = Fecha_Fin_Fase
                     'Flog.writeline "Valor de Aux_Fecha: " & Aux_fecha
                     StrSql = " SELECT * FROM his_estructura "
                     StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                     StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro & " AND "
                     StrSql = StrSql & " his_estructura.estrnro = " & ListParam.Empresa
                     StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                     Flog.writeline UCase("Hago la consulta sobre los históricos de estructura")
                     'Flog.writeline "Consulta: " & StrSql
                     Set rs_Estructura = New ADODB.Recordset
                     OpenRecordset StrSql, rs_Estructura
                     If Err.Number <> 0 Then
                         Flog.writeline UCase("Error: ") & Err.Number & " - Desc: " & Err.Description
                         Err.Clear
                     End If
                     If rs_Estructura.RecordCount > 1 Then rs_Estructura.MoveFirst
                     Do While Not rs_Estructura.EOF
                         If rs_Estructura!htetdesde <= Fecha_Inicio_Fase Then
                             If rs_Estructura!htethasta = Fecha_Fin_Fase Then
                                 Fecha_Inicio_Fase = DatosPeriodo.pliqdesde
                                 'Aux_fecha = Fecha_Fin_Fase
                                 'Flog.writeline
                                 Flog.writeline UCase("Fecha de alta de fase definitiva: ") & Fecha_Inicio_Fase
                                 Flog.writeline UCase("Fecha de baja de fase definitiva: ") & Fecha_Fin_Fase
                                 'Flog.writeline "Valor de Aux_Fecha: " & Aux_fecha
                                 'Flog.writeline
                                 FechaAuxAsignada = True
                                 Exit Do
                             End If
                         End If
                         rs_Estructura.MoveNext
                     Loop
                     If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                     Set rs_Estructura = Nothing
                     If Not FechaAuxAsignada Then
                         rs_fases.MoveNext
                     Else
                         Exit Do
                     End If
             Loop
         Else 'Si no encuentra fases para el empleado sigue con el proximo registro
         End If
    End If '------
    Flog.writeline ""
    If rs_fases.State = adStateOpen Then rs_fases.Close
    Set rs_fases = Nothing
End Sub



Public Sub ErrCabecera()
    If NroErrCab <> 1 And NroError <> "" Then
        StrSql = "SELECT repsuedigtiperr FROM rep_ar_libsuedigital_err WHERE repsuedigerrnro =" & NroError
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            NroErrCab = CLng(rs!repsuedigtiperr)
        End If
        rs.Close
    End If
 End Sub
 


